"""
FedEx Pickup Manifest monitor -- runs on a schedule (every 15 minutes) and posts a
Slack alert when a pickup is about to miss its close time.

Trigger (exactly as specified): for each stop on today's Pickup Manifest --
    PU Closed is empty      (pickup not yet closed out / "delivered time" still blank)
    AND Reas Code is empty  (no reason on file for not picking up)
    AND Close time is <= 60 minutes away (or has already passed)

WHAT THIS SCRIPT DOES EACH RUN
  1. Logs in. Two modes, in priority order:
       - Auto-login: if FEDEX_USERNAME and FEDEX_PASSWORD are set (as GitHub secrets,
         or local env vars), it fills and submits FedEx's login form itself every run.
         Can't get past an MFA/OTP challenge if FedEx ever shows one -- that still
         needs a human.
       - Saved session: otherwise it loads cookies from ./storage_state.json, created
         once by running fedex_login.py yourself (a real browser window you log into).
     Either way, this script's own code never prints, logs, or stores your password
     anywhere -- it only reads it from an environment variable at run time.
  2. Navigates to the Pickup Manifest tab, Work Area = ALL, Date = today, clicks Search.
  3. Clicks the Excel export icon and captures the download.
  4. Runs pickup_risk.find_at_risk() against it.
  5. Looks up ./state.db (SQLite) to see which of those stops it already alerted on, so
     the SAME stop isn't re-sent every 15 minutes. It alerts again only when a stop's
     status escalates (AT_RISK -> PAST_CLOSE) or a new stop enters the 60-minute window.
  6. Posts a Slack alert -- one TABLE listing every stop currently within 60 minutes
     of its close time (or past it), resent only when that set actually changes (a
     stop drops off once picked up, a new one is added once it enters the window) --
     via an Incoming Webhook (SLACK_WEBHOOK_URL).
  7. Separately, if SLACK_BOT_TOKEN and SLACK_DAILY_CHANNEL are set: posts (once per
     day) and then edits in place every 30 minutes a second Slack message, in its own
     channel, listing EVERY stop on today's manifest (Station, Shipper, Address,
     Ready, Close, PU Closed, Reas Code), with rows flagged 🟡 within 1hr of close /
     🔴 past close with nothing on file. See the SLACK_BOT_TOKEN comment above for
     how to get that token -- this is a separate credential from SLACK_WEBHOOK_URL
     because editing a message needs the Slack Web API, not an Incoming Webhook.

This version is built to run on an ephemeral runner (GitHub Actions) that has no disk
of its own between runs -- see fedex_login.py for the one-time login step, and
.github/workflows/fedex-monitor.yml for the schedule. It also runs fine on a normal
machine if you'd rather use cron/Task Scheduler instead; see README.md either way.

LIVE-VERIFIED SELECTORS
  Confirmed against a real logged-in session with real data (53-row manifest): the
  homepage's "FCC Links" tile is its own iframe (title="FCC Links") containing a
  "FedEx Customer Connection" link that opens the app in a new tab; Date/Work
  Area/Search on the Pickup Manifest tab; and the Excel export control, which is
  `<input type="image" id="manifestForm:buttonGenerateExcel">` -- not a normal
  button/link. If FedEx changes the page later and `export_and_download` can't find
  that id, it falls back to a couple of looser guesses, and if even those fail it
  prints every icon/button it found on screen instead of guessing blind.
"""
from __future__ import annotations

import argparse
import os
import re
import signal
import sqlite3
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from pickup_risk import load_manifest, find_at_risk, format_at_risk_table, build_daily_rows

HERE = Path(__file__).resolve().parent
STORAGE_STATE_PATH = HERE / "storage_state.json"
DOWNLOAD_DIR = HERE / "downloads"
STATE_DB = HERE / "state.db"

# Delaware is Eastern time -- change this if your station is elsewhere.
STATION_TZ = ZoneInfo("America/New_York")
WARN_MINUTES = 60  # back to 1 hour -- was widened to 120 for testing on 2026-10-19
FAILURE_ALERT_COOLDOWN_MIN = 120  # don't spam Slack more than once every 2h on failure
LOGIN_RETRY_BACKOFF_SEC = 120  # gap between retried logins in --watch mode (see
# login_with_retries() -- was 60s; widened 2026-10-19 after console diagnostics
# showed signs of a lingering server-side session colliding with the next login

# Pickups only happen in this window -- --watch mode logs out and goes fully idle
# (no FedEx traffic at all, not even a keep-alive click) outside it, and logs back in
# right at the start of the next window. Only read by run_watch(); run_once() always
# runs when called, so a plain cron/systemd-timer setup controls its own hours via
# OnCalendar/schedule instead.
ACTIVE_WINDOW_START_HOUR = 11  # 11:00 AM Delaware time
ACTIVE_WINDOW_END_HOUR = 19    # 7:00 PM Delaware time (exclusive)


def _next_active_window_start(now: datetime) -> datetime:
    """If `now` is already inside the active window, returns `now` unchanged.
    Otherwise returns the next moment the window opens (today's or tomorrow's
    ACTIVE_WINDOW_START_HOUR)."""
    start_today = now.replace(hour=ACTIVE_WINDOW_START_HOUR, minute=0, second=0, microsecond=0)
    end_today = now.replace(hour=ACTIVE_WINDOW_END_HOUR, minute=0, second=0, microsecond=0)
    if now < start_today:
        return start_today
    if now >= end_today:
        return start_today + timedelta(days=1)
    return now

class GracefulShutdown(BaseException):
    """Raised from a SIGTERM/SIGINT handler so `systemctl stop`/`restart` unwinds
    through run_watch's normal try/finally (which logs out of FedEx before exiting)
    instead of killing the process mid-step. Deliberately a BaseException, not
    Exception: this codebase has many broad `except Exception: pass` guards for
    resilience against page/network flakiness, and those must NOT swallow a
    shutdown request -- only KeyboardInterrupt/SystemExit normally skip them, and
    this needs the same treatment."""


def _install_shutdown_handler() -> None:
    def _handler(signum, frame):
        raise GracefulShutdown()
    signal.signal(signal.SIGTERM, _handler)
    signal.signal(signal.SIGINT, _handler)


HOMEPAGE_URL = "https://mybizaccount.fedex.com/"
SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL", "")

# "Daily Manifest" channel -- a separate Slack channel that gets ONE message per day
# (all of today's pickups, not just at-risk ones), edited in place every
# DAILY_UPDATE_INTERVAL_MIN minutes instead of re-posted. Editing an existing message
# needs the Slack Web API (chat.postMessage / chat.update) with a real Bot Token --
# an Incoming Webhook (SLACK_WEBHOOK_URL above) can only post new messages, never
# edit one, so this is intentionally a second, separate credential:
#   1. On the Slack App used for SLACK_WEBHOOK_URL (or a new one), add the
#      "chat:write" Bot Token Scope under OAuth & Permissions, reinstall the app,
#      and copy the "Bot User OAuth Token" (starts with xoxb-).
#   2. Invite that bot into the target channel (/invite @YourAppName there).
#   3. Set SLACK_BOT_TOKEN to that token and SLACK_DAILY_CHANNEL to the channel's ID
#      (right-click the channel in Slack -> View channel details -> copy Channel ID,
#      looks like C0123456789 -- NOT the #channel-name).
# Both unset == this feature is simply skipped (nothing breaks).
SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN", "")
SLACK_DAILY_CHANNEL = os.environ.get("SLACK_DAILY_CHANNEL", "")
DAILY_UPDATE_INTERVAL_MIN = 20  # matches the At Risk channel's full-check cadence
# Optional: a healthchecks.io (or similar) ping URL. Pinged once per run -- the plain
# URL on success, "<url>/fail" on a failed run -- so that service can show a run
# history/dashboard and alert if a scheduled run doesn't check in at all (meaning the
# systemd timer didn't fire, the script crashed before finishing, or the VPS is down).
# Entirely optional: if unset, this is just skipped.
HEALTHCHECK_URL = os.environ.get("HEALTHCHECK_URL", "").rstrip("/")


def ping_healthcheck(log, ok: bool) -> None:
    if not HEALTHCHECK_URL:
        return
    url = HEALTHCHECK_URL if ok else f"{HEALTHCHECK_URL}/fail"
    try:
        requests.get(url, timeout=10)
    except Exception as e:
        # Never let a monitoring ping itself break the actual monitor.
        log(f"Healthcheck ping failed (non-fatal): {e!r}")


# --------------------------------------------------------------------------- state db

def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(STATE_DB)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS alerted (
               pu_list INTEGER NOT NULL,
               puid INTEGER NOT NULL,
               status TEXT NOT NULL,
               first_alerted_at TEXT NOT NULL,
               last_alerted_at TEXT NOT NULL,
               PRIMARY KEY (pu_list, puid)
           )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS meta (
               key TEXT PRIMARY KEY,
               value TEXT NOT NULL
           )"""
    )
    conn.commit()
    return conn


def filter_new_or_escalated(conn: sqlite3.Connection, stops, now_iso: str):
    """Returns the subset of `stops` that should actually be sent to Slack: ones never
    alerted before, or ones whose status just got worse (AT_RISK -> PAST_CLOSE)."""
    out = []
    for stop in stops:
        pu_list, puid = stop.key
        row = conn.execute(
            "SELECT status FROM alerted WHERE pu_list=? AND puid=?", (pu_list, puid)
        ).fetchone()
        if row is None:
            conn.execute(
                "INSERT INTO alerted VALUES (?,?,?,?,?)",
                (pu_list, puid, stop.status, now_iso, now_iso),
            )
            out.append(stop)
        elif row[0] != stop.status:
            conn.execute(
                "UPDATE alerted SET status=?, last_alerted_at=? WHERE pu_list=? AND puid=?",
                (stop.status, now_iso, pu_list, puid),
            )
            out.append(stop)
        else:
            conn.execute(
                "UPDATE alerted SET last_alerted_at=? WHERE pu_list=? AND puid=?",
                (now_iso, pu_list, puid),
            )
    conn.commit()
    return out


def should_send_failure_alert(conn: sqlite3.Connection, now: datetime) -> bool:
    """Caps failure alerts (e.g. 'session expired') to one per cooldown window, since a
    broken session fails every single 15-minute run until someone fixes it."""
    row = conn.execute("SELECT value FROM meta WHERE key='last_failure_alert'").fetchone()
    if row is not None:
        last = datetime.fromisoformat(row[0])
        if now - last < timedelta(minutes=FAILURE_ALERT_COOLDOWN_MIN):
            return False
    conn.execute(
        "INSERT INTO meta(key, value) VALUES ('last_failure_alert', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (now.isoformat(),),
    )
    conn.commit()
    return True


# ------------------------------------------------------------------------------ slack

def send_slack(text: str, dry_run: bool = False, blocks: list[dict] | None = None) -> None:
    """text is always sent as Slack's required 'fallback' (shown in notifications/
    previews, and the whole message if blocks is omitted). blocks, when given, is
    the actual rich Block Kit layout rendered in the channel."""
    if dry_run or not SLACK_WEBHOOK_URL:
        print("[dry-run / no SLACK_WEBHOOK_URL set] would post to Slack:\n" + text + "\n")
        return
    payload = {"text": text}
    if blocks:
        payload["blocks"] = blocks
    resp = requests.post(SLACK_WEBHOOK_URL, json=payload, timeout=15)
    if resp.status_code >= 300:
        print(f"Slack post failed ({resp.status_code}): {resp.text}", file=sys.stderr)


# ------------------------------------------------------------------ slack (Web API)
# Only used by the "Daily Manifest" channel below, which needs to EDIT a message
# in place -- something an Incoming Webhook (send_slack, above) cannot do.

def slack_api_post(channel: str, text: str, blocks: list[dict], dry_run: bool) -> Optional[str]:
    """Posts a new message via chat.postMessage. Returns its ts (needed to edit it
    later), or None if it wasn't actually sent (dry-run, not configured, or failed)."""
    if dry_run or not SLACK_BOT_TOKEN or not channel:
        print(f"[dry-run / Daily Manifest not configured] would POST to {channel or '(no channel set)'}:\n{text}\n")
        return None
    resp = requests.post(
        "https://slack.com/api/chat.postMessage",
        headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        json={"channel": channel, "text": text, "blocks": blocks},
        timeout=15,
    )
    data = resp.json()
    if not data.get("ok"):
        print(f"Slack chat.postMessage failed: {data}", file=sys.stderr)
        return None
    return data.get("ts")


def slack_api_update(channel: str, ts: str, text: str, blocks: list[dict], dry_run: bool) -> bool:
    """Edits an existing message (by channel + ts) via chat.update."""
    if dry_run or not SLACK_BOT_TOKEN or not channel or not ts:
        print(f"[dry-run / Daily Manifest not configured] would UPDATE {channel or '(no channel set)'} ts={ts}:\n{text}\n")
        return False
    resp = requests.post(
        "https://slack.com/api/chat.update",
        headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        json={"channel": channel, "ts": ts, "text": text, "blocks": blocks},
        timeout=15,
    )
    data = resp.json()
    if not data.get("ok"):
        print(f"Slack chat.update failed: {data}", file=sys.stderr)
        return False
    return True


_STATUS_ICON = {"warn": "🟡", "overdue": "🔴", "pending": "🔵", "done": "🟢"}


def _chunk_table_lines(header: str, lines: list[str], max_chars: int = 2800) -> list[str]:
    """Slack's section text is capped at 3000 chars, and a single day's manifest
    (50-90 rows) can run well past that -- splits into several code-block chunks,
    each repeating the header, rather than truncating rows."""
    chunks, cur, cur_len = [], [header], len(header)
    for line in lines:
        if cur_len + len(line) + 1 > max_chars:
            chunks.append("\n".join(cur))
            cur, cur_len = [header], len(header)
        cur.append(line)
        cur_len += len(line) + 1
    if len(cur) > 1:
        chunks.append("\n".join(cur))
    return chunks


def build_daily_manifest_message(rows: list[dict], now: datetime) -> tuple[str, list[dict]]:
    n_warn = sum(1 for r in rows if r["status"] == "warn")
    n_overdue = sum(1 for r in rows if r["status"] == "overdue")
    summary = (f"*📦 FedEx Daily Pickup Manifest — {now:%Y-%m-%d}*\n"
               f"Updated {now:%I:%M %p %Z} · next update in ~{DAILY_UPDATE_INTERVAL_MIN} min · "
               f"{len(rows)} stop(s) · 🟡 {n_warn} within 1hr  🔴 {n_overdue} overdue")

    header = (f"{'':<3}{'STN':<7}{'WA':<8}{'SHIPPER':<22}{'ADDRESS':<28}{'READY':<8}"
              f"{'CLOSE':<8}{'PUCLSD':<9}{'REAS'}")
    lines = []
    for r in rows:
        icon = _STATUS_ICON.get(r["status"], "  ")
        shipper = (r["shipper"] or "")[:19]
        address = (r["address"] or "")[:25]
        station = r["station"] if r["station"] is not None else ""
        wa = r.get("wa") or ""
        lines.append(f"{icon:<3}{station:<7}{wa:<8}{shipper:<22}{address:<28}{r['ready']:<8}"
                      f"{r['close']:<8}{r['pu_closed']:<9}{r['reas_code']}")

    blocks = [{"type": "section", "text": {"type": "mrkdwn", "text": summary}}]
    for chunk in _chunk_table_lines(header, lines):
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": f"```{chunk}```"}})
    text = f"Daily Pickup Manifest — {now:%Y-%m-%d} — {len(rows)} stop(s)"
    return text, blocks


def update_daily_manifest_channel(df, now: datetime, log, dry_run: bool) -> None:
    """Posts today's full manifest table once (first successful check of the day),
    then edits that SAME message in place every DAILY_UPDATE_INTERVAL_MIN minutes
    for the rest of the day. No-ops entirely if SLACK_BOT_TOKEN/SLACK_DAILY_CHANNEL
    aren't set."""
    if not SLACK_BOT_TOKEN or not SLACK_DAILY_CHANNEL:
        return

    conn = db_connect()

    def get_meta(key: str) -> Optional[str]:
        row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return row[0] if row else None

    def set_meta(key: str, value: str) -> None:
        conn.execute(
            "INSERT INTO meta(key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )

    today_str = now.date().isoformat()
    stored_date = get_meta("daily_manifest_date")
    stored_ts = get_meta("daily_manifest_ts")
    stored_last_update = get_meta("daily_manifest_last_update")

    rows = build_daily_rows(df, now, warn_minutes=WARN_MINUTES, manifest_day=now)
    text, blocks = build_daily_manifest_message(rows, now)

    if stored_date != today_str or not stored_ts:
        ts = slack_api_post(SLACK_DAILY_CHANNEL, text, blocks, dry_run)
        if ts or dry_run:
            set_meta("daily_manifest_date", today_str)
            set_meta("daily_manifest_ts", ts or "")
            set_meta("daily_manifest_last_update", now.isoformat())
            conn.commit()
            log(f"Daily Manifest: posted today's message ({len(rows)} stop(s)).")
    else:
        last_update = datetime.fromisoformat(stored_last_update) if stored_last_update else None
        if last_update is None or (now - last_update) >= timedelta(minutes=DAILY_UPDATE_INTERVAL_MIN):
            ok = slack_api_update(SLACK_DAILY_CHANNEL, stored_ts, text, blocks, dry_run)
            if ok or dry_run:
                set_meta("daily_manifest_last_update", now.isoformat())
                conn.commit()
                log(f"Daily Manifest: updated today's message in place ({len(rows)} stop(s)).")
    conn.close()


# --------------------------------------------------------------------------- browser

def is_logged_out(page) -> bool:
    """FedEx's login form has a password field; if we see one after navigating to the
    homepage, the saved session has expired and nothing past this point will work."""
    try:
        return page.locator('input[type="password"]').count() > 0
    except Exception:
        return False


_LOGIN_FIELD_SELECTOR = (
    'input[type="password"], input[name="identifier"], input[name="username"], '
    '#okta-signin-username'
)


def find_login_scope(page, log):
    """Returns whichever Page or Frame actually contains the login fields. Okta often
    renders its sign-in widget inside an iframe (seen live: #account-chooser-iframe)
    rather than at the top level, so a top-level-only search can silently find
    nothing and skip the login attempt entirely."""
    try:
        if page.locator(_LOGIN_FIELD_SELECTOR).count() > 0:
            return page
    except Exception:
        pass
    for frame in page.frames:
        try:
            if frame.locator(_LOGIN_FIELD_SELECTOR).count() > 0:
                return frame
        except Exception:
            continue
    log(f"No login fields found at top level or in any iframe. Frames present: "
        + str([f.url for f in page.frames]))
    return page  # fall back so the caller's own diagnostics still fire


def attempt_credential_login(page, log) -> bool:
    """If FEDEX_USERNAME / FEDEX_PASSWORD are set, tries to fill and submit FedEx's
    login form. Returns True if that got past the password screen. These selectors
    are best-effort, standard-pattern guesses -- I don't fill in a real login form
    myself (that stays the user's own action), so this hasn't been confirmed against
    FedEx's actual markup the way the Pickup Manifest selectors were. If it's wrong,
    it prints every input field it found so the one-line fix is obvious. It also
    can't click through an MFA/OTP challenge if FedEx shows one -- that still needs a
    human, same as before."""
    username = os.environ.get("FEDEX_USERNAME")
    password = os.environ.get("FEDEX_PASSWORD")
    if not username or not password:
        return False

    # The actual login form lives on FedEx's Okta SSO page (purpleid.okta.com), reached
    # after clicking Sign In -- often inside an iframe, and Okta frequently splits this
    # into two steps (username -> Next -> password) rather than one combined form.
    try:
        scope = find_login_scope(page, log)

        user_field = scope.locator(
            'input[name="identifier"], input[name="username"], '
            '#okta-signin-username, input[type="email"], '
            'input[name*="user" i], input[id*="user" i]'
        ).first
        if user_field.count() == 0:
            user_field = scope.get_by_label(re.compile(r"user\s*(id|name)|email", re.I)).first
        user_field.fill(username, timeout=8000)

        if scope.locator('input[type="password"]').count() == 0:
            # Two-step flow: submit the username first and wait for the password step.
            next_btn = scope.get_by_role("button", name=re.compile(r"next|continue|submit", re.I)).first
            if next_btn.count() == 0:
                next_btn = scope.locator('input[type="submit"]').first
            if next_btn.count() > 0:
                next_btn.click(timeout=8000)
                page.wait_for_timeout(2500)
                scope = find_login_scope(page, log)  # the password step may be a new iframe

        pass_field = scope.locator('input[type="password"], #okta-signin-password').first
        pass_field.fill(password, timeout=8000)

        submit = scope.get_by_role("button", name=re.compile(r"sign\s*in|log\s*in|verify", re.I)).first
        if submit.count() == 0:
            submit = scope.locator('input[type="submit"], #okta-signin-submit').first
        submit.click(timeout=8000)
        page.wait_for_timeout(4000)
    except Exception as e:
        log(f"Automatic login attempt raised an error: {e!r}")
        return False

    # Still on Okta's domain, or still looking at a password field, means it didn't
    # actually get through -- most likely an MFA/OTP step this script can't complete
    # unattended, or a selector above didn't match. Either way, print what's on screen.
    if "okta.com" in page.url or is_logged_out(page):
        log(f"Login did not complete. URL: {page.url!r}, title: {page.title()!r}. "
            "Visible <input> fields, for debugging:")
        try:
            fields = page.locator("input").evaluate_all(
                "els => els.map(e => ({type: e.type, name: e.name, id: e.id, "
                "placeholder: e.placeholder}))"
            )
            log(str(fields))
        except Exception:
            pass
        return False

    log("Automatic login succeeded.")
    return True


def attempt_logout(page, log) -> None:
    """Best-effort sign-out before abandoning a session. ctx.close() only drops our
    end of the TCP connection -- it never tells FedEx's PeopleSoft/ICEfaces app that
    this session is done, so the server-side session can stay "alive" for its own
    timeout. Console diagnostics on 2026-10-19 showed FedEx's own JS refusing a new
    session with "Blocking connection cannot be shared among multiple web-contexts"
    right after a day of many logins in a row without ever logging out -- one
    plausible, well-supported contributor to that pile-up. Never raises; a missed or
    failed logout just means the old session expires on its own timeout as before."""
    if page is None or page.is_closed():
        return
    try:
        signout = page.get_by_role("link", name=re.compile(r"sign\s*out|log\s*out", re.I)).first
        if signout.count() == 0:
            signout = page.locator(
                'a:has-text("Sign Out"), a:has-text("Log Out"), '
                'button:has-text("Sign Out"), button:has-text("Log Out")'
            ).first
        if signout.count() > 0:
            signout.click(timeout=5000)
            page.wait_for_timeout(1500)
            log("Clicked Sign Out before abandoning this session.")
        else:
            log("No Sign Out control found on this page -- leaving the session to "
                "expire on its own timeout.")
    except Exception as e:
        log(f"Best-effort logout failed (not fatal, continuing): {e!r}")


def robust_click(locator, log, description: str) -> bool:
    """Tries a raw JS .click() FIRST, then a normal click, then a forced click.
    FedEx serves this script a MOBILE layout (ice.mobi.* handlers) where several
    controls sit outside the viewport -- a real (coordinate-based) click, forced or
    not, can silently land nowhere on an off-screen element, while el.click() calls
    the element's own click handler directly regardless of position or visibility,
    which is both more reliable here and avoids scrolling tricks. The two
    coordinate-based clicks are kept only as a fallback in case a handler relies on
    real mouse-event properties the JS method doesn't set. Returns False (after
    logging) only if all three fail. Also logs which element actually matched --
    a click can "succeed" against the wrong element if the locator matched more
    than one thing or something decorative."""
    try:
        info = locator.evaluate(
            "el => ({tag: el.tagName, id: el.id, type: el.type || null, "
            "disabled: !!el.disabled, cls: el.className})"
        )
    except Exception:
        info = None
    log(f"About to click {description}. Matched element: {info!r}")

    err1 = err2 = err3 = None
    try:
        locator.evaluate("el => el.click()")
        log(f"Clicked {description} via raw JS el.click().")
        return True
    except Exception as e:
        err1 = repr(e)
    try:
        locator.click(timeout=8000)
        log(f"Clicked {description} normally (JS click failed: {err1}).")
        return True
    except Exception as e:
        err2 = repr(e)
    try:
        locator.click(timeout=4000, force=True)
        log(f"Clicked {description} with force=True (JS and normal both failed: "
            f"{err1}, {err2}).")
        return True
    except Exception as e:
        err3 = repr(e)
        log(f"Could not click {description} by any method: {err3}")
        return False


def goto_pickup_manifest(ctx, page, log):
    """Navigation from the FedEx homepage to the Pickup Manifest tab, Work Area = ALL,
    today's date, Search already clicked. Returns the Page the manifest ended up on
    (FCC opens in a new tab), "LOGGED_OUT" if the saved session has expired, or None
    if some other step failed."""
    page.goto(HOMEPAGE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)

    # Not logged in yet, this lands on a pre-login "Welcome to MyBizAccount" page with
    # just a "Sign In" button -- no password field until that's clicked, which is why
    # is_logged_out() alone doesn't catch this state. Click it to reveal the real form.
    if not is_logged_out(page):
        signin = page.get_by_role("button", name=re.compile(r"sign\s*in", re.I)).first
        if signin.count() == 0:
            signin = page.locator(
                'input[type="submit"][value*="Sign In" i], '
                'input[type="button"][value*="Sign In" i], '
                'a:has-text("Sign In")'
            ).first
        if signin.count() > 0:
            try:
                signin.click(timeout=8000)
                page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(2000)
                log(f"Clicked Sign In. URL: {page.url!r}, title: {page.title()!r}")
            except Exception as e:
                log(f"Found a Sign In control but clicking it failed: {e!r}")

    # Being on okta.com at all means we're not actually logged in yet, even if the
    # top-level page has no <input type="password"> -- Okta's widget is often inside
    # an iframe (find_login_scope handles that), so don't rely on is_logged_out()
    # alone to decide whether a login attempt is needed.
    if is_logged_out(page) or "okta.com" in page.url:
        if attempt_credential_login(page, log):
            page.wait_for_timeout(1000)
        else:
            log("Saved session appears to be logged out, and no automatic login "
                "succeeded. If you've set FEDEX_USERNAME/FEDEX_PASSWORD, check the "
                "log output above for why the login attempt failed; otherwise "
                "re-run fedex_login.py locally and update the storage_state secret.")
            return "LOGGED_OUT"

    log(f"Appears logged in. URL: {page.url!r}, title: {page.title()!r}")

    # A fresh login (new IP/device, as a CI runner always is) can land on an F5
    # access-policy interstitial (URL contains "my.policy") instead of straight on the
    # PeopleSoft homepage -- not something seen during live testing with an already-
    # established session. Best-effort click through it, and log what's actually on
    # the page either way, since the exact wording here hasn't been confirmed yet.
    if "my.policy" in page.url:
        try:
            visible_links = page.locator("a, button, input[type='submit']").all_text_contents()
        except Exception:
            visible_links = []
        log(f"Landed on an F5 policy/portal page ({page.title()!r}). "
            f"Visible links/buttons: {[t.strip() for t in visible_links if t.strip()]}")
        try:
            candidate = page.get_by_role(
                "link", name=re.compile("mybizaccount|my biz|continue|portal", re.I)
            ).first
            if candidate.count() == 0:
                candidate = page.locator("a, button").filter(
                    has_text=re.compile("mybizaccount|my biz|continue|portal", re.I)
                ).first
            if candidate.count() > 0:
                candidate.click(timeout=8000)
                page.wait_for_load_state("domcontentloaded")
                page.wait_for_timeout(2000)
                log(f"After clicking through the policy page -- URL: {page.url!r}, "
                    f"title: {page.title()!r}")
        except Exception as e:
            log(f"Could not click through the policy page: {e!r}")

    fcc_page = page
    # Collected from whichever page fcc_page ends up being (attached below, once we
    # know which one that is) so a stuck-render diagnostic can show real JS errors
    # instead of just a timeout with no explanation.
    console_messages: list[str] = []

    def _capture_console(target_page):
        target_page.on("console", lambda m: console_messages.append(f"[{m.type}] {m.text}"))
        target_page.on("pageerror", lambda e: console_messages.append(f"[pageerror] {e}"))
        # Stashed as a plain attribute (Playwright Page objects allow this) so any
        # OTHER function holding this same page object -- export_and_download, in
        # particular -- can read the live, growing console log too, without needing
        # its own listener wired up or the list threaded through every call site.
        target_page._console_log = console_messages

    if "cpc-mi" in page.url:
        _capture_console(fcc_page)  # already on the FCC app page, nothing more to open
    else:
        # The homepage's tiles are each their own iframe and load asynchronously, so
        # wait for the specific one we need rather than a fixed sleep -- a headless CI
        # runner can be slower to render these than a real browser.
        try:
            page.wait_for_selector('iframe[title="FCC Links"]', timeout=20000)
        except PWTimeout:
            log(f"'FCC Links' iframe never appeared. Current URL: {page.url!r}, "
                f"title: {page.title()!r}. iframes present: "
                + str(page.locator("iframe").evaluate_all(
                    "els => els.map(e => ({id: e.id, title: e.title, name: e.name}))"
                )))
            return None

        # The homepage's tiles are each their own iframe; "FedEx Customer Connection"
        # lives inside the one titled "FCC Links" and opens the app in a NEW tab.
        try:
            frame = page.frame_locator('iframe[title="FCC Links"]')
            with ctx.expect_page(timeout=15000) as new_page_info:
                frame.get_by_text(re.compile("FedEx Customer Connection", re.I)).first.click(timeout=10000)
            fcc_page = new_page_info.value
            _capture_console(fcc_page)
            fcc_page.wait_for_load_state("domcontentloaded")
            # The new tab goes through an intermediate redirect (its title briefly
            # reads "Loading https://...") before the actual app renders -- wait for a
            # real element instead of a fixed delay, since that redirect's duration
            # isn't consistent (fine live, too slow on a CI runner in practice).
            try:
                fcc_page.wait_for_selector("text=P&D Manifests", timeout=20000)
            except PWTimeout:
                pass  # diagnostics below will show whatever actually loaded
            log(f"Opened FCC. URL: {fcc_page.url!r}, title: {fcc_page.title()!r}")
        except PWTimeout:
            log("Could not find/click 'FedEx Customer Connection' inside the FCC "
                "Links iframe (the iframe itself did render). Links visible in it: "
                + str(frame.locator("a, [role='link']").all_text_contents()))
            return None

    # Make sure the Pickup Manifest (P&D Manifests) tab is selected. The bottom-nav
    # items are plain <li> elements (role="listitem", not button/link), which can
    # make Playwright's normal actionability checks (occlusion/stability) time out
    # even though the element is really there and a real click would land fine.
    pdm = fcc_page.get_by_text(re.compile(r"^P&D Manifests$", re.I)).first
    pdm_ok = robust_click(pdm, log, "'P&D Manifests' tab")
    if not pdm_ok:
        log("Bottom-nav items visible: "
            + str(fcc_page.locator("text=/Messaging|P&D Manifests|Service Area Status|"
                                    "Pickup Alerts|SCH PU Mgmt|P&D Planning/").all_text_contents()))
    log(f"Clicked 'P&D Manifests' tab (ok={pdm_ok}).")

    # The desktop UI loads noticeably slower than the mobile one we were previously
    # (accidentally) served -- wait for the actual content pane (the Search button,
    # confirmed live as id="manifestForm:search") rather than a fixed 1s delay.
    try:
        fcc_page.wait_for_selector('input[name="manifestForm:search"]', timeout=30000)
        log("Pickup Manifest form (Date/Work Area/Search) has rendered.")
    except PWTimeout:
        log("Timed out waiting for the Pickup Manifest form to render after clicking "
            f"the tab. Page text so far: {fcc_page.locator('body').inner_text()[:500]!r}")
        # This exact failure has repeated many times (2026-10-17, 2026-10-19) with no
        # change in the text dumps. A full_page screenshot turned out to be a dead
        # end -- chrome-headless-shell (what this VPS actually runs, confirmed via
        # `systemctl status`) is known to produce solid-black full_page screenshots
        # when the frame is captured mid-paint, which is exactly what we got: a plain
        # black image with no diagnostic value. Dump real DOM/HTML/console evidence
        # instead, which headless-shell doesn't corrupt. Overwrites the same files
        # each time on purpose -- only the latest occurrence matters for this.
        try:
            html_path = HERE / "debug_stuck_render.html"
            html_path.write_text(fcc_page.content(), encoding="utf-8")
            log(f"Saved the page's live HTML to {html_path} ({html_path.stat().st_size} bytes)")
        except Exception as e:
            log(f"Could not save debug HTML: {e!r}")
        try:
            iframes = fcc_page.locator("iframe").evaluate_all(
                "els => els.map(e => ({id: e.id, title: e.title, name: e.name, "
                "src: e.src, w: e.clientWidth, h: e.clientHeight}))")
            log(f"iframes on the page right now: {iframes}")
        except Exception as e:
            log(f"Could not inspect iframes: {e!r}")
        try:
            scroll_info = fcc_page.locator("body").evaluate("""body => {
                const all = [...document.querySelectorAll("*")];
                return all.filter(el => {
                    const s = getComputedStyle(el);
                    return el.scrollWidth > el.clientWidth + 20 && el.clientWidth > 0 &&
                           (s.overflowX === "auto" || s.overflowX === "scroll");
                }).map(el => ({tag: el.tagName, id: el.id,
                                className: String(el.className).slice(0, 120),
                                clientWidth: el.clientWidth, scrollWidth: el.scrollWidth}))
                .slice(0, 20);
            }""")
            log(f"Horizontal-scroll containers found: {scroll_info}")
        except Exception as e:
            log(f"Could not run the scroll-container diagnostic: {e!r}")
        try:
            manifest_els = fcc_page.locator(
                '[id*="manifestForm"], [id*="GenerateExcel"], table, [role="grid"]'
            ).evaluate_all(
                """els => els.map(e => ({tag: e.tagName, id: e.id,
                    visible: !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length),
                    display: getComputedStyle(e).display})).slice(0, 50)"""
            )
            log(f"manifestForm/grid/export elements anywhere in the DOM (even if hidden): {manifest_els}")
        except Exception as e:
            log(f"Could not inspect manifestForm/grid elements: {e!r}")
        if console_messages:
            log(f"Browser console/page errors captured since this page loaded "
                f"({len(console_messages)} total, showing up to 30): {console_messages[:30]}")
        else:
            log("No browser console or page errors were captured.")

        # The console diagnostics above (2026-10-19, reproduced on every attempt that
        # day) consistently show FedEx's own ICEfaces JS logging "singleSubmit failed
        # cannot find enclosing form" right as the tab click happens -- that's ICEfaces'
        # click handler erroring out specifically on our *scripted* el.click() (fired
        # via locator.evaluate("el => el.click()") in robust_click), which never
        # throws so robust_click reports success without ever trying a real click.
        # One retry with an actual trusted, OS-level mouse click -- a fundamentally
        # different kind of event than a scripted one -- before giving up on this
        # attempt entirely.
        log("Retrying the 'P&D Manifests' tab with a real (trusted) mouse click, "
            "since the scripted click may be what's tripping ICEfaces' handler.")
        try:
            pdm.click(timeout=8000)
            log("Sent a real mouse click on the tab. Waiting again for the form...")
            fcc_page.wait_for_selector('input[name="manifestForm:search"]', timeout=20000)
            log("Pickup Manifest form rendered after the real-click retry.")
        except PWTimeout:
            log("Still timed out after the real-click retry -- treating this attempt "
                "as a genuinely stuck render.")
        except Exception as e:
            log(f"Real-click retry itself failed: {e!r}")

    # Work Area and Date are BOTH confirmed live (screenshot, 2026-10-17) to already
    # default to ALL / today on their own -- the script used to spend up to 30s per
    # attempt explicitly re-setting Work Area to "ALL" for no reason, which also
    # timed out on nearly every run today. Removed entirely: nothing left to set
    # before Search, so go straight to clicking it.

    # Click Search -- confirmed live as id="manifestForm:search", so target that
    # directly instead of relying only on a generic role/name search.
    search_btn = fcc_page.locator('input[name="manifestForm:search"]').first
    if search_btn.count() == 0:
        search_btn = fcc_page.get_by_role("button", name=re.compile("^Search$", re.I)).first
    search_ok = robust_click(search_btn, log, "the Search button")
    log(f"Clicked Search ok={search_ok}.")
    if not search_ok:
        log(f"URL: {fcc_page.url!r}, title: {fcc_page.title()!r}. Buttons visible: "
            + str(fcc_page.get_by_role("button").all_text_contents())
            + ". Page text (first 500 chars): "
            + fcc_page.locator("body").inner_text()[:500])
        return None

    # A full manifest (dozens of rows, seen live) can take noticeably longer than a
    # few seconds to actually render after Search -- wait for the "N records found"
    # text FedEx shows once results are in, rather than guessing a fixed delay.
    try:
        results_el = fcc_page.wait_for_selector("text=/records found/i", timeout=25000)
        log(f"Search results rendered: {results_el.inner_text()!r}")
    except PWTimeout:
        log("Timed out waiting for '...records found' text after clicking Search. "
            f"Current URL: {fcc_page.url!r}, title: {fcc_page.title()!r}. "
            "Page text (first 800 chars): " + fcc_page.locator("body").inner_text()[:800])
    fcc_page.wait_for_timeout(1000)
    return fcc_page


def dismiss_popups(page, log, max_rounds: int = 3) -> None:
    """Closes any modal/notification dialogs sitting on top of the page (seen live:
    the only clickable elements on screen were several 'close' buttons, blocking the
    actual manifest controls underneath). Best-effort -- clicks anything that looks
    like a close control, up to a few rounds in case closing one reveals another."""
    close_pattern = re.compile(r"^(close|×|x|dismiss|ok|got it)$", re.I)
    dismissed_total = 0
    for _ in range(max_rounds):
        closed_this_round = 0
        try:
            # Broad on purpose: the live page showed plain text 'close' on screen,
            # which could be a <button>, <a>, <span> or <div> -- not necessarily
            # something with an ARIA "button" role, so match by visible text too.
            candidates = page.get_by_role("button", name=close_pattern).all()
            candidates += page.get_by_role("link", name=close_pattern).all()
            candidates += page.locator(
                "[aria-label='Close' i], [title='Close' i], .close, .ui-dialog-titlebar-close"
            ).all()
            candidates += page.get_by_text(close_pattern, exact=True).all()
        except Exception:
            candidates = []
        for c in candidates:
            try:
                if c.is_visible():
                    c.click(timeout=2000)
                    closed_this_round += 1
                    page.wait_for_timeout(500)
            except Exception:
                continue
        dismissed_total += closed_this_round
        if closed_this_round == 0:
            break
    if dismissed_total:
        log(f"Dismissed {dismissed_total} popup dialog(s) before continuing.")


def export_and_download(page, log, ctx=None) -> Path | None:
    """Clicks the Excel export control and returns the path of the downloaded file,
    or None if there was nothing to export (zero open stops is a normal outcome)."""
    # A fuller manifest (dozens of rows, seen live) can take longer to finish
    # rendering than a fixed short sleep -- wait for the actual export control to
    # show up rather than guessing a delay, before trying anything else.
    try:
        page.wait_for_selector('[id="manifestForm:buttonGenerateExcel"]', timeout=25000)
    except PWTimeout:
        pass  # fall through to dismiss_popups/candidates/diagnostics below either way

    dismiss_popups(page, log)
    # Confirmed live: it's `<input type="image" id="manifestForm:buttonGenerateExcel">`.
    # The id is the one actually confirmed live -- id and name usually match on a JSF/
    # PrimeFaces component like this, but that was never directly confirmed, so the id
    # selector goes first and the name-based one stays only as a fallback.
    candidates = [
        lambda: page.locator('[id="manifestForm:buttonGenerateExcel"]').first,
        lambda: page.locator('input[name="manifestForm:buttonGenerateExcel"]').first,
        lambda: page.locator('input[type="image"][src*="excel" i]').first,
        lambda: page.get_by_role("link", name=re.compile("excel|export", re.I)).first,
        lambda: page.get_by_role("button", name=re.compile("excel|export", re.I)).first,
        lambda: page.locator("img[alt*='xcel' i], img[title*='xcel' i]").first,
        lambda: page.locator("[id*='xcel' i], [id*='export' i]").first,
    ]
    for make_locator in candidates:
        try:
            loc = make_locator()
            if loc.count() == 0:
                continue
        except Exception:
            continue

        dest = DOWNLOAD_DIR / f"manifest_{datetime.now():%Y%m%d_%H%M%S}.xls"

        # Confirmed live behavior: clicking this opens a NEW window that downloads
        # the file and then closes itself automatically -- often too fast to reliably
        # grab a reference to that popup page and attach a download listener to it
        # after the fact. A context-level "download" listener catches the event no
        # matter which page (this one, or a popup that's already gone) fired it, so
        # that's used here instead of page.expect_download()/ctx.expect_page().
        downloader = ctx if ctx is not None else page.context
        try:
            with downloader.expect_event("download", timeout=25000) as dl_info:
                if not robust_click(loc, log, "the Excel export control"):
                    raise RuntimeError("click itself failed")
            dl_info.value.save_as(dest)
            return dest
        except Exception as e:
            # The click matching and firing successfully but no download event ever
            # arriving (confirmed live 2026-10-20, ~30+ min into a session) needs
            # more than "timed out" to actually diagnose -- so capture what's
            # happening around the failed attempt: any extra tabs/popups sitting
            # open (the export is documented to open one and self-close; one stuck
            # open, or several piling up across retries, is a real clue), and
            # whatever the page's own JS has logged most recently (reuses the
            # console/pageerror capture goto_pickup_manifest already wired up --
            # see _capture_console's `_console_log` stash -- no new listener here).
            try:
                open_pages = [{"url": p.url, "title": p.title()} for p in downloader.pages]
            except Exception:
                open_pages = "(could not enumerate open pages)"
            recent_console = getattr(page, "_console_log", [])[-15:]
            log(f"Download did not complete for this candidate: {e!r}\n"
                f"  Open tabs/popups right now ({len(open_pages) if isinstance(open_pages, list) else '?'}): {open_pages}\n"
                f"  Most recent console/page-error entries: {recent_console}")
            # The export is documented to open a popup that downloads and
            # self-closes -- if one failed to self-close, it's another browser
            # context sharing this same FedEx session, which is exactly the shape
            # of the "Blocking connection cannot be shared among multiple
            # web-contexts" collision seen before. Close anything stray before the
            # next candidate/attempt, rather than letting them pile up.
            try:
                for p in downloader.pages:
                    if p is not page and not p.is_closed():
                        stray_url = p.url
                        p.close()
                        log(f"  Closed a stray popup/tab left open: {stray_url!r}")
            except Exception as cleanup_err:
                log(f"  (stray-popup cleanup itself failed, non-fatal: {cleanup_err!r})")
            continue

    # Nothing worked -- dump what IS on screen so this is a quick fix, not a guess.
    try:
        elements = page.locator("img, a, button, span, div").evaluate_all(
            "els => els.filter(e => e.innerText && e.innerText.trim().length && "
            "e.innerText.trim().length < 40 && e.offsetParent !== null)"
            ".slice(0, 40).map(e => ({tag: e.tagName, text: e.innerText.trim(), "
            "id: e.id, cls: e.className}))"
        )
        log("Could not find the Excel export control. Visible elements with short text "
            "on this screen were:\n  " + "\n  ".join(str(e) for e in elements))
    except Exception as e:
        log(f"Could not find the Excel export control, and the diagnostic dump itself "
            f"failed: {e!r}")
    return None


# ---------------------------------------------------------------------------- cycle

def process_manifest(fcc_page, ctx, log, dry_run: bool) -> bool:
    """Runs one real check (export -> risk scan -> Slack alerts) against an
    already-logged-in, already-searched Pickup Manifest page. Shared by run_once
    (which only ever does this once) and run_watch (which calls it repeatedly on the
    same session). Returns False only when something about the page itself looks
    broken -- callers treat that as 'this session probably needs a fresh login' --
    not when there's simply nothing to export (that's a normal, healthy outcome)."""
    now = datetime.now(STATION_TZ)

    try:
        xls_path = export_and_download(fcc_page, log, ctx=ctx)
    except Exception as e:
        log(f"export_and_download raised an error (treating as a dead session): {e!r}")
        ping_healthcheck(log, ok=False)
        return False

    if xls_path is None:
        # "No export button found" is ambiguous on its own -- it's the normal, healthy
        # shape of a genuinely empty manifest, but it's ALSO exactly what a stuck/
        # broken render looks like (Search got clicked but the results pane never
        # actually came up -- the original failure mode this was built to catch).
        # Confirmed live: FedEx's "nothing matched" wording isn't consistent --
        # "0 records found" sometimes, "No results match your search request" other
        # times (seen 2026-10-17) -- so don't match either phrase specifically. What's
        # actually reliable is whether the manifestForm itself is still on the page:
        # that's only there once the tab has genuinely rendered, regardless of which
        # wording FedEx used for an empty result.
        looks_rendered = False
        try:
            looks_rendered = fcc_page.locator('input[name="manifestForm:search"]').count() > 0
        except Exception:
            pass

        if looks_rendered:
            # The form being present isn't enough on its own to call this healthy --
            # 2026-10-20 15:45 showed exactly why: form rendered, "99 records found"
            # on screen, the export control clicked successfully 4 times in a row,
            # but the download itself never fired. The OLD version of this check
            # treated "form present + no export" as always meaning a genuine 0-result
            # day and returned healthy=True unconditionally -- which silently
            # skipped alerting AND skipped updating both Slack channels for the rest
            # of the day, with nothing in Slack to show anything was wrong. Now it
            # actually reads the on-screen result count and only calls this healthy
            # when it genuinely says 0/no results; a real count (or one we can't
            # parse) is treated as a real failure, not assumed benign.
            result_text = ""
            try:
                result_text = fcc_page.locator(
                    '[id="manifestForm:outputResultCount"], [id="manifestForm:panelResultCount"]'
                ).first.inner_text()
            except Exception:
                pass
            m = re.search(r"(\d+)\s*records?\s*found", result_text, re.I)
            genuinely_empty = ("no results" in result_text.lower()
                                or (m is not None and int(m.group(1)) == 0))
            if genuinely_empty:
                log(f"No export captured this cycle -- genuinely 0 results "
                    f"({result_text.strip()!r}). Nothing to export, healthy.")
                ping_healthcheck(log, ok=True)
                return True
            log(f"No export captured this cycle, but the page shows "
                f"{result_text.strip() or '(no result-count text found)'!r} -- "
                "NOT a genuine empty result, so treating this as a real export "
                "failure (not silently healthy). This also means run_watch will "
                "log back in fresh, which is what cleared this exact symptom "
                "the last time it happened.")
            ping_healthcheck(log, ok=False)
            return False

        log("No export AND the manifest form itself isn't on the page anymore -- "
            "this looks like a stuck/broken render (columns never came up), not a "
            "genuine answer from FedEx. Treating this session as dead.")
        ping_healthcheck(log, ok=False)
        return False

    log(f"Downloaded {xls_path.name}")

    df = load_manifest(xls_path)
    risky = find_at_risk(df, now=now, warn_minutes=WARN_MINUTES, manifest_day=now)
    log(f"{len(risky)} stop(s) currently within the {WARN_MINUTES}-minute warning window.")

    # "At Risk" channel: one table of EVERY currently-at-risk/past-close stop, resent
    # as a new message only when that set actually changes (a stop drops off once
    # picked up, a new one is added once it enters the window) -- not one message
    # per stop, and not a resend every cycle if nothing changed.
    current_keys = ",".join(f"{k}:{v}" for k, v in sorted(s.key for s in risky))
    conn = db_connect()
    row = conn.execute("SELECT value FROM meta WHERE key='last_at_risk_keys'").fetchone()
    prev_keys = row[0] if row else None
    if current_keys != (prev_keys or ""):
        if risky:
            text, blocks = format_at_risk_table(risky, now)
            send_slack(text, dry_run=dry_run, blocks=blocks)
            log(f"At-risk set changed -- sent a fresh table ({len(risky)} stop(s)).")
        elif prev_keys:  # had stops before, now zero -- say so once
            send_slack("✅ All previously at-risk FedEx pickups have been closed out "
                        "-- none currently at risk.", dry_run=dry_run)
            log("At-risk set changed -- now empty, sent all-clear message.")
        conn.execute(
            "INSERT INTO meta(key, value) VALUES ('last_at_risk_keys', ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (current_keys,),
        )
        conn.commit()
    else:
        log("At-risk set unchanged since last check -- not resending the table.")
    conn.close()

    update_daily_manifest_channel(df, now, log, dry_run)

    ping_healthcheck(log, ok=True)
    return True


def keepalive_probe(fcc_page, log) -> bool:
    """A cheap 'are we still logged in' check for the gaps between full checks in
    --watch mode. Re-clicking Search is the whole point, not just the test: it's
    also what resets PeopleSoft's own server-side idle timeout, so the session stays
    usable without a brand new Okta login every few minutes. Nothing is exported or
    parsed here -- only process_manifest() does that."""
    try:
        search_btn = fcc_page.locator('input[name="manifestForm:search"]').first
        if search_btn.count() == 0:
            search_btn = fcc_page.get_by_role("button", name=re.compile("^Search$", re.I)).first
        if not robust_click(search_btn, log, "the Search button (keep-alive probe)"):
            return False
        fcc_page.wait_for_selector("text=/records found/i", timeout=20000)
        log("Keep-alive probe: session still alive.")
        return True
    except Exception as e:
        log(f"Keep-alive probe failed (treating as a dead session): {e!r}")
        return False


def run_once(headless: bool, dry_run: bool) -> int:
    now = datetime.now(STATION_TZ)

    def log(msg):
        # Timestamp at call time, not run_once's start time -- otherwise every line
        # in a run shows the same instant and timing issues (like a navigation step
        # outrunning the page) are invisible in the logs.
        print(f"[{datetime.now(STATION_TZ):%Y-%m-%d %H:%M:%S %Z}] {msg}")

    have_creds = bool(os.environ.get("FEDEX_USERNAME")) and bool(os.environ.get("FEDEX_PASSWORD"))
    if not STORAGE_STATE_PATH.exists() and not have_creds:
        log(f"Neither {STORAGE_STATE_PATH.name} nor FEDEX_USERNAME/FEDEX_PASSWORD are "
            "set -- run fedex_login.py once locally (saved-session mode), or set "
            "those two env vars / secrets (auto-login mode).")
        return 1

    DOWNLOAD_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        # channel="chromium" forces the full Chromium binary's headless mode instead of
        # Playwright's default "chromium_headless_shell" -- confirmed 2026-10-20: manual
        # logins in a real browser worked every time, while headless_shell consistently
        # failed at the exact same ICEfaces AJAX step ("singleSubmit failed cannot find
        # enclosing form" / "Blocking connection cannot be shared among multiple
        # web-contexts") on every attempt for 2+ hours regardless of session cleanup.
        # headless_shell is a stripped-down binary with known gaps vs real Chrome on
        # older/complex JS apps -- this ICEfaces app (a 2000s-era JSF AJAX framework)
        # is exactly that kind of app. Requires `playwright install chromium` once
        # (headless_shell is downloaded by default; the full chromium channel is not).
        browser = p.chromium.launch(headless=headless, channel="chromium")
        # FedEx serves a MOBILE layout to this script (onclick handlers under
        # "ice.mobi.*") -- forcing a desktop UA/viewport was tried and made things
        # worse (the content pane stopped rendering within any reasonable wait), so
        # that's reverted. The mobile layout itself renders content fine; the only
        # real problem was clicking elements that sit off-screen, which robust_click
        # now handles by trying a coordinate-free JS click first (see below).
        ctx_kwargs = {"accept_downloads": True}
        # Only fall back to the saved cookies when there's no username/password to do
        # a clean login with -- see the matching comment in run_watch()'s login() for
        # why: an old, effectively-dead storage_state.json presented alongside a fresh
        # credential login looks like a plausible contributor to the session-collision
        # symptoms seen 2026-10-19.
        if not have_creds and STORAGE_STATE_PATH.exists():
            ctx_kwargs["storage_state"] = str(STORAGE_STATE_PATH)
        ctx = browser.new_context(**ctx_kwargs)
        page = ctx.new_page()

        fcc_page = goto_pickup_manifest(ctx, page, log)

        if fcc_page in (None, "LOGGED_OUT"):
            browser.close()
            conn = db_connect()
            if should_send_failure_alert(conn, now):
                msg = ("⚠️ FedEx Pickup Manifest check failed this cycle"
                       + (" -- saved login session has expired, needs a fresh login."
                          if fcc_page == "LOGGED_OUT" else
                          " -- see Action run logs for details.")
                       + " (this alert is sent at most once every "
                       f"{FAILURE_ALERT_COOLDOWN_MIN} minutes while it stays broken)")
                send_slack(msg, dry_run=dry_run)
            conn.close()
            ping_healthcheck(log, ok=False)
            return 1

        ok = process_manifest(fcc_page, ctx, log, dry_run)
        browser.close()

    return 0 if ok else 1


def run_watch(dry_run: bool, headless: bool, full_interval_min: int = 20,
              keepalive_min: int = 4, max_consecutive_failures: int = 3) -> int:
    """Alternative to run_once for long-running use (systemd Type=simple instead of a
    timer firing run_once every cycle). Logs in ONCE and keeps that same browser
    session alive for as long as it can, instead of a brand new Okta SSO handshake
    every 10-15 minutes:
      - every `keepalive_min` minutes: a cheap keep-alive probe (re-click Search,
        nothing exported) that also resets PeopleSoft's own idle timeout.
      - every `full_interval_min` minutes: the real check (export, risk scan, Slack
        alerts) via process_manifest(), same logic run_once uses.
    A fresh login only happens when a probe or full check finds the session actually
    dead -- not on a fixed schedule -- so a login that was flaky with the old
    every-cycle approach gets far fewer chances to land on whatever FedEx-side
    variance (mobile vs. desktop UI, F5 interstitial, etc.) causes that flakiness.
    After `max_consecutive_failures` dead-session logins in a row, gives up and
    exits non-zero so systemd (Restart=always) starts over with a clean browser."""
    def log(msg):
        print(f"[{datetime.now(STATION_TZ):%Y-%m-%d %H:%M:%S %Z}] {msg}")

    have_creds = bool(os.environ.get("FEDEX_USERNAME")) and bool(os.environ.get("FEDEX_PASSWORD"))
    if not STORAGE_STATE_PATH.exists() and not have_creds:
        log(f"Neither {STORAGE_STATE_PATH.name} nor FEDEX_USERNAME/FEDEX_PASSWORD are "
            "set -- run fedex_login.py once locally (saved-session mode), or set "
            "those two env vars / secrets (auto-login mode).")
        return 1

    DOWNLOAD_DIR.mkdir(exist_ok=True)
    full_interval = timedelta(minutes=full_interval_min)
    keepalive_interval = timedelta(minutes=keepalive_min)
    consecutive_failures = 0
    # Tracks full-check failures that happen AFTER a successful login (export stuck /
    # columns never render) -- separate from login failures above, because those were
    # already alerted on but this case previously was not: login can keep succeeding
    # while Search/export keeps failing, which looped silently forever with zero Slack
    # alert. EXPORT_STUCK_ALERT_THRESHOLD full checks in a row (not counting keep-alive
    # probes) is enough to rule out a one-off FedEx blip before paging the user.
    EXPORT_STUCK_ALERT_THRESHOLD = 2
    # Hard cap on how many times in a row a full-check failure (export stuck, or a
    # genuinely dead session) is allowed to trigger an immediate fresh relogin.
    # Without this, a PERSISTENT problem (not a one-off blip) would relogin every
    # single cycle with zero backoff between attempts -- repeated rapid-fire Okta
    # logins, which is exactly the kind of thing that caused the session-collision
    # mess on 2026-10-19. After this many in a row, give up and exit instead (same
    # "3 strikes" exit-1 pattern login_with_retries already uses) -- systemd
    # (Restart=always, RestartSec=30) then restarts with a real pause and a
    # completely clean browser, rather than this process hammering FedEx forever.
    MAX_CONSECUTIVE_EXPORT_FAILURES = 3
    consecutive_export_failures = 0
    today_confirmed_date = None  # date() of the last "today's check is live" ping sent

    def send_failure_alert(reason):
        conn = db_connect()
        if should_send_failure_alert(conn, datetime.now(STATION_TZ)):
            if reason == "LOGGED_OUT":
                detail = " -- saved login session has expired, needs a fresh login."
            elif reason == "EXPORT_STUCK":
                detail = (" -- logged in fine, but Search/Excel export keeps failing to "
                           "render even after re-login (this looks like the FedEx page "
                           "itself getting stuck, not a login problem). See "
                           "debug_stuck_render.png and the service logs on the server.")
            else:
                detail = " -- see service logs for details."
            msg = ("⚠️ FedEx Pickup Manifest check failed this cycle" + detail
                   + " (this alert is sent at most once every "
                   f"{FAILURE_ALERT_COOLDOWN_MIN} minutes while it stays broken)")
            send_slack(msg, dry_run=dry_run)
        conn.close()
        ping_healthcheck(log, ok=False)

    _install_shutdown_handler()

    with sync_playwright() as p:
        # See the matching comment in run_once() -- channel="chromium" uses the full
        # Chromium binary's headless mode instead of the stripped-down
        # chromium_headless_shell default, which consistently failed to render this
        # ICEfaces app's AJAX-driven manifest form while manual logins in a real
        # browser worked every time (confirmed 2026-10-20).
        browser = p.chromium.launch(headless=headless, channel="chromium")
        ctx = None
        fcc_page = None

        def login(close_old: bool):
            nonlocal ctx
            if close_old and ctx is not None:
                attempt_logout(fcc_page, log)
                try:
                    ctx.close()
                except Exception:
                    pass
            ctx_kwargs = {"accept_downloads": True}
            # Only fall back to the saved cookies when there's no username/password to
            # do a clean login with. When FEDEX_USERNAME/PASSWORD are set (the normal
            # case here), every attempt was ALSO loading this old, effectively-dead
            # storage_state.json cookie on top of a fresh credential login (confirmed
            # by the logs always showing "Automatic login succeeded" -- the saved
            # cookie never actually skipped login, so it was never doing anything
            # useful). Console diagnostics on 2026-10-19 caught FedEx's own JS warning
            # about a session "cannot be shared among multiple web-contexts" right as
            # a render got stuck -- presenting a stale old session cookie alongside a
            # brand new login is exactly the kind of thing that could cause that.
            if not have_creds and STORAGE_STATE_PATH.exists():
                ctx_kwargs["storage_state"] = str(STORAGE_STATE_PATH)
            ctx = browser.new_context(**ctx_kwargs)
            page = ctx.new_page()
            result = goto_pickup_manifest(ctx, page, log)
            if result in (None, "LOGGED_OUT"):
                return None, result
            return result, None

        def login_with_retries(close_old: bool) -> bool:
            """Tries to log in, retrying up to max_consecutive_failures times before
            giving up -- the SAME resilience a mid-day session death already gets,
            applied here too so one bad render right at window-open doesn't take the
            whole day down on a single failed attempt.
            Backoff between attempts (LOGIN_RETRY_BACKOFF_SEC) is deliberately longer
            than a quick retry: console diagnostics on 2026-10-19 showed FedEx's own
            ICEfaces JS warning "Blocking connection cannot be shared among multiple
            web-contexts" right as a render got stuck -- i.e. the previous (failed/
            killed) session's server-side connection hadn't fully torn down yet when
            the next login started. Giving it more breathing room between attempts
            reduces that collision."""
            nonlocal fcc_page, consecutive_failures
            for attempt in range(1, max_consecutive_failures + 1):
                fcc_page, fail_reason = login(close_old=close_old)
                if fcc_page is not None:
                    consecutive_failures = 0
                    return True
                consecutive_failures += 1
                send_failure_alert(fail_reason)
                if attempt < max_consecutive_failures:
                    log(f"Login attempt {attempt}/{max_consecutive_failures} failed -- "
                        f"retrying in {LOGIN_RETRY_BACKOFF_SEC}s.")
                    time.sleep(LOGIN_RETRY_BACKOFF_SEC)
                close_old = True  # clean up the failed attempt's own context before retrying
            log(f"{max_consecutive_failures} login attempts in a row failed -- "
                "giving up for now.")
            return False

        try:
            now = datetime.now(STATION_TZ)
            window_start = _next_active_window_start(now)
            if window_start > now:
                wait_s = (window_start - now).total_seconds()
                log(f"Outside the {ACTIVE_WINDOW_START_HOUR}:00-{ACTIVE_WINDOW_END_HOUR}:00 "
                    f"{now:%Z} active window -- not logging in; idle for "
                    f"{wait_s / 60:.0f} min until it opens.")
                time.sleep(wait_s)

            if not login_with_retries(close_old=False):
                # Every attempt failed -- exit non-zero so systemd (Restart=always)
                # starts the whole process over with a clean browser rather than
                # looping forever inside one already-struggling run.
                return 1

            next_full = datetime.now(STATION_TZ)  # run a full check right away
            next_keepalive = datetime.now(STATION_TZ) + keepalive_interval
            log(f"Watch mode started: full check every {full_interval_min} min, "
                f"keep-alive probe every {keepalive_min} min, active "
                f"{ACTIVE_WINDOW_START_HOUR}:00-{ACTIVE_WINDOW_END_HOUR}:00 {now:%Z}.")

            while True:
                now = datetime.now(STATION_TZ)

                if now.hour >= ACTIVE_WINDOW_END_HOUR or now.hour < ACTIVE_WINDOW_START_HOUR:
                    log(f"End of the active window ({ACTIVE_WINDOW_START_HOUR}:00-"
                        f"{ACTIVE_WINDOW_END_HOUR}:00 {now:%Z}) -- logging out and "
                        "going idle until it reopens (no FedEx traffic at all while idle).")
                    try:
                        ctx.close()
                    except Exception:
                        pass
                    ctx, fcc_page = None, None
                    window_start = _next_active_window_start(now)
                    wait_s = (window_start - now).total_seconds()
                    time.sleep(max(wait_s, 0))

                    if not login_with_retries(close_old=False):
                        return 1
                    next_full = datetime.now(STATION_TZ)
                    next_keepalive = datetime.now(STATION_TZ) + keepalive_interval
                    continue

                healthy = True

                if now >= next_full:
                    log("-- full check --")
                    healthy = process_manifest(fcc_page, ctx, log, dry_run)
                    if healthy:
                        consecutive_export_failures = 0
                        next_full = now + full_interval
                        next_keepalive = now + keepalive_interval
                        if today_confirmed_date != now.date():
                            send_slack(
                                f"✅ FedEx Pickup Manifest check is live for today "
                                f"({now:%Y-%m-%d}) -- logged in, Search/Excel export "
                                "confirmed working.", dry_run=dry_run)
                            today_confirmed_date = now.date()
                    else:
                        consecutive_export_failures += 1
                        if consecutive_export_failures >= EXPORT_STUCK_ALERT_THRESHOLD:
                            send_failure_alert("EXPORT_STUCK")
                        if consecutive_export_failures >= MAX_CONSECUTIVE_EXPORT_FAILURES:
                            log(f"{MAX_CONSECUTIVE_EXPORT_FAILURES} full checks in a row have "
                                "failed (export stuck, or session dying) even after relogins "
                                "in between -- giving up for now instead of continuing to "
                                "relogin every cycle with no backoff. Exiting; systemd "
                                "(Restart=always, RestartSec=30) will restart with a real "
                                "pause and a completely fresh browser.")
                            return 1
                elif now >= next_keepalive:
                    log("-- keep-alive probe --")
                    healthy = keepalive_probe(fcc_page, log)
                    if healthy:
                        next_keepalive = now + keepalive_interval

                if not healthy:
                    log("Session looks dead -- logging back in from scratch (this is "
                        "the only time a fresh Okta login happens in --watch mode).")
                    if not login_with_retries(close_old=True):
                        return 1
                    next_full = datetime.now(STATION_TZ)
                    next_keepalive = datetime.now(STATION_TZ) + keepalive_interval
                    continue

                time.sleep(15)
        except GracefulShutdown:
            # systemctl stop/restart (or Ctrl+C) -- fall through to the `finally`
            # below so FedEx gets a real sign-out before this process dies, instead
            # of SIGTERM killing it mid-step and leaving a stale server-side session
            # for the next login to fight through (the exact "blocking connection"
            # collision debugged 2026-10-19).
            log("Received shutdown signal -- logging out of FedEx before exiting.")
        finally:
            if ctx is not None:
                attempt_logout(fcc_page, log)
                try:
                    ctx.close()
                except Exception:
                    pass
            browser.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--headful", action="store_true", help="show the browser window for this run")
    ap.add_argument("--dry-run", action="store_true", help="print alerts instead of posting to Slack")
    ap.add_argument("--watch", action="store_true",
                     help="log in once and keep running: a full check every "
                          "--full-interval minutes, with cheap keep-alive Search "
                          "clicks every --keepalive-interval minutes in between, "
                          "instead of a brand new Okta login every cycle")
    ap.add_argument("--full-interval", type=int, default=20,
                     help="minutes between full checks in --watch mode (default: 20)")
    ap.add_argument("--keepalive-interval", type=int, default=4,
                     help="minutes between keep-alive Search clicks in --watch mode (default: 4)")
    ap.add_argument("--max-failures", type=int, default=3,
                     help="consecutive failed re-logins in --watch mode before giving "
                          "up and exiting (default: 3)")
    args = ap.parse_args()

    if args.watch:
        sys.exit(run_watch(dry_run=args.dry_run, headless=not args.headful,
                            full_interval_min=args.full_interval,
                            keepalive_min=args.keepalive_interval,
                            max_consecutive_failures=args.max_failures))
    sys.exit(run_once(headless=not args.headful, dry_run=args.dry_run))
