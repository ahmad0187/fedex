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
  6. Posts new/escalated alerts to Slack via an Incoming Webhook (SLACK_WEBHOOK_URL).

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
import sqlite3
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from pickup_risk import load_manifest, find_at_risk, format_alert

HERE = Path(__file__).resolve().parent
STORAGE_STATE_PATH = HERE / "storage_state.json"
DOWNLOAD_DIR = HERE / "downloads"
STATE_DB = HERE / "state.db"

# Delaware is Eastern time -- change this if your station is elsewhere.
STATION_TZ = ZoneInfo("America/New_York")
WARN_MINUTES = 60
FAILURE_ALERT_COOLDOWN_MIN = 120  # don't spam Slack more than once every 2h on failure

# Pickups only happen in this window -- --watch mode logs out and goes fully idle
# (no FedEx traffic at all, not even a keep-alive click) outside it, and logs back in
# right at the start of the next window. Only read by run_watch(); run_once() always
# runs when called, so a plain cron/systemd-timer setup controls its own hours via
# OnCalendar/schedule instead.
ACTIVE_WINDOW_START_HOUR = 8   # 8:00 AM Delaware time
ACTIVE_WINDOW_END_HOUR = 20    # 8:00 PM Delaware time (exclusive)


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

HOMEPAGE_URL = "https://mybizaccount.fedex.com/"
SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL", "")
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

def send_slack(text: str, dry_run: bool = False) -> None:
    if dry_run or not SLACK_WEBHOOK_URL:
        print("[dry-run / no SLACK_WEBHOOK_URL set] would post to Slack:\n" + text + "\n")
        return
    resp = requests.post(SLACK_WEBHOOK_URL, json={"text": text}, timeout=15)
    if resp.status_code >= 300:
        print(f"Slack post failed ({resp.status_code}): {resp.text}", file=sys.stderr)


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
    if "cpc-mi" not in page.url:
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

    # Work Area -> ALL (it's usually the default, but set it explicitly).
    wa_ok = False
    try:
        wa = fcc_page.locator("select").filter(has_text="ALL").first
        wa.select_option(label="ALL")
        wa_ok = True
    except Exception as e:
        log(f"Could not explicitly set Work Area to ALL (leaving default): {e!r}")

    # Diagnostic only (no behavior change): the script has never actually set the Date
    # field -- the docstring assumed it defaults to today, but that's never been
    # confirmed, and an error-styled "No results match your search request" (seen
    # 2026-10-17) is exactly what a stale/blank/wrong Date would produce. Logging
    # whatever it holds right now, so the next time that happens we can tell a real
    # empty manifest apart from a bad default without guessing.
    try:
        date_fields = fcc_page.locator('input[id*="date" i], input[name*="date" i]').evaluate_all(
            "els => els.map(e => ({id: e.id, name: e.name, value: e.value}))"
        )
        if date_fields:
            log(f"Date field(s) on the form right now: {date_fields}")
    except Exception:
        pass

    # Click Search -- confirmed live as id="manifestForm:search", so target that
    # directly instead of relying only on a generic role/name search.
    search_btn = fcc_page.locator('input[name="manifestForm:search"]').first
    if search_btn.count() == 0:
        search_btn = fcc_page.get_by_role("button", name=re.compile("^Search$", re.I)).first
    search_ok = robust_click(search_btn, log, "the Search button")
    log(f"Work Area set ok={wa_ok}; clicked Search ok={search_ok}.")
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
        page.wait_for_selector('input[name="manifestForm:buttonGenerateExcel"]', timeout=25000)
    except PWTimeout:
        pass  # fall through to dismiss_popups/candidates/diagnostics below either way

    dismiss_popups(page, log)
    # Confirmed live: it's `<input type="image" id="manifestForm:buttonGenerateExcel">`,
    # not a normal button/link -- that's tried first; the rest are fallbacks only.
    candidates = [
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
            log(f"Download did not complete for this candidate: {e!r}")
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
    now_iso = now.isoformat()

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
            log("No export captured this cycle (0 results for this search, or the "
                "export control needs a one-line selector fix -- see any message "
                "above).")
            ping_healthcheck(log, ok=True)
            return True
        log("No export AND the manifest form itself isn't on the page anymore -- "
            "this looks like a stuck/broken render (columns never came up), not a "
            "genuine answer from FedEx. Treating this session as dead.")
        ping_healthcheck(log, ok=False)
        return False

    log(f"Downloaded {xls_path.name}")

    df = load_manifest(xls_path)
    risky = find_at_risk(df, now=now, warn_minutes=WARN_MINUTES, manifest_day=now)
    log(f"{len(risky)} stop(s) currently within the {WARN_MINUTES}-minute warning window.")

    conn = db_connect()
    to_send = filter_new_or_escalated(conn, risky, now_iso)
    conn.close()

    for stop in to_send:
        send_slack(format_alert(stop), dry_run=dry_run)
    log(f"Sent {len(to_send)} new/escalated Slack alert(s).")
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
        browser = p.chromium.launch(headless=headless)
        # FedEx serves a MOBILE layout to this script (onclick handlers under
        # "ice.mobi.*") -- forcing a desktop UA/viewport was tried and made things
        # worse (the content pane stopped rendering within any reasonable wait), so
        # that's reverted. The mobile layout itself renders content fine; the only
        # real problem was clicking elements that sit off-screen, which robust_click
        # now handles by trying a coordinate-free JS click first (see below).
        ctx_kwargs = {"accept_downloads": True}
        if STORAGE_STATE_PATH.exists():
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


def run_watch(dry_run: bool, headless: bool, full_interval_min: int = 15,
              keepalive_min: int = 2, max_consecutive_failures: int = 3) -> int:
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

    def send_failure_alert(reason):
        conn = db_connect()
        if should_send_failure_alert(conn, datetime.now(STATION_TZ)):
            msg = ("⚠️ FedEx Pickup Manifest check failed this cycle"
                   + (" -- saved login session has expired, needs a fresh login."
                      if reason == "LOGGED_OUT" else
                      " -- see service logs for details.")
                   + " (this alert is sent at most once every "
                   f"{FAILURE_ALERT_COOLDOWN_MIN} minutes while it stays broken)")
            send_slack(msg, dry_run=dry_run)
        conn.close()
        ping_healthcheck(log, ok=False)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        ctx = None
        fcc_page = None

        def login(close_old: bool):
            nonlocal ctx
            if close_old and ctx is not None:
                try:
                    ctx.close()
                except Exception:
                    pass
            ctx_kwargs = {"accept_downloads": True}
            if STORAGE_STATE_PATH.exists():
                ctx_kwargs["storage_state"] = str(STORAGE_STATE_PATH)
            ctx = browser.new_context(**ctx_kwargs)
            page = ctx.new_page()
            result = goto_pickup_manifest(ctx, page, log)
            if result in (None, "LOGGED_OUT"):
                return None, result
            return result, None

        def login_with_retries(close_old: bool) -> bool:
            """Tries to log in, retrying up to max_consecutive_failures times (60s
            apart) before giving up -- the SAME resilience a mid-day session death
            already gets, applied here too so one bad render right at window-open
            doesn't take the whole day down on a single failed attempt."""
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
                        "retrying in 60s.")
                    time.sleep(60)
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
                        next_full = now + full_interval
                        next_keepalive = now + keepalive_interval
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
        finally:
            if ctx is not None:
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
    ap.add_argument("--full-interval", type=int, default=15,
                     help="minutes between full checks in --watch mode (default: 15)")
    ap.add_argument("--keepalive-interval", type=int, default=2,
                     help="minutes between keep-alive Search clicks in --watch mode (default: 2)")
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
