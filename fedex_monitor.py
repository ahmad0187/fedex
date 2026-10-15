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

HOMEPAGE_URL = "https://mybizaccount.fedex.com/"
SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL", "")


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
    """Tries a normal click, then a forced click (skips Playwright's occlusion/
    visibility actionability checks), then a raw JS .click() as a last resort. Some of
    FedEx's nav elements are plain non-interactive tags (<li>, <span>) with a JS click
    handler rather than a real <button>/<a>, which can make Playwright's normal
    actionability wait time out even though the element is genuinely there and
    clickable. Returns False (after logging) only if all three fail."""
    try:
        locator.click(timeout=8000)
        return True
    except Exception:
        pass
    try:
        locator.click(timeout=4000, force=True)
        return True
    except Exception:
        pass
    try:
        locator.evaluate("el => el.click()")
        return True
    except Exception as e:
        log(f"Could not click {description} by any method: {e!r}")
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
    if not robust_click(pdm, log, "'P&D Manifests' tab"):
        log("Bottom-nav items visible: "
            + str(fcc_page.locator("text=/Messaging|P&D Manifests|Service Area Status|"
                                    "Pickup Alerts|SCH PU Mgmt|P&D Planning/").all_text_contents()))
    fcc_page.wait_for_timeout(1000)

    # Work Area -> ALL (it's usually the default, but set it explicitly).
    try:
        wa = fcc_page.locator("select").filter(has_text="ALL").first
        wa.select_option(label=re.compile(r"^ALL$"))
    except Exception:
        pass  # leave whatever it defaults to

    # Click Search.
    search_btn = fcc_page.get_by_role("button", name=re.compile("^Search$", re.I)).first
    if not robust_click(search_btn, log, "the Search button"):
        log(f"URL: {fcc_page.url!r}, title: {fcc_page.title()!r}. Buttons visible: "
            + str(fcc_page.get_by_role("button").all_text_contents())
            + ". Page text (first 500 chars): "
            + fcc_page.locator("body").inner_text()[:500])
        return None

    fcc_page.wait_for_timeout(3000)
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
            candidates = page.get_by_role("button", name=close_pattern).all()
            candidates += page.locator(
                "[aria-label='Close' i], [title='Close' i], .close, .ui-dialog-titlebar-close"
            ).all()
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


def export_and_download(page, log) -> Path | None:
    """Clicks the Excel export control and returns the path of the downloaded file,
    or None if there was nothing to export (zero open stops is a normal outcome)."""
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
            with page.expect_download(timeout=15000) as dl_info:
                loc.click()
            download = dl_info.value
            dest = DOWNLOAD_DIR / f"manifest_{datetime.now():%Y%m%d_%H%M%S}.xls"
            download.save_as(dest)
            return dest
        except Exception:
            continue

    # Nothing worked -- dump what IS on screen so this is a quick fix, not a guess.
    try:
        icons = page.locator("img, a, button").all_text_contents()
        log("Could not find the Excel export control. Visible clickable elements on "
            "this screen were:\n  " + "\n  ".join(repr(t) for t in icons[:40] if t.strip()))
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------- cycle

def run_once(headless: bool, dry_run: bool) -> int:
    now = datetime.now(STATION_TZ)
    now_iso = now.isoformat()

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
            return 1

        xls_path = export_and_download(fcc_page, log)
        browser.close()

        if xls_path is None:
            log("No export captured this cycle (could genuinely be zero open pickups "
                "right now, or the export control needs a one-line selector fix -- see "
                "any message above).")
            return 0

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
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--headful", action="store_true", help="show the browser window for this run")
    ap.add_argument("--dry-run", action="store_true", help="print alerts instead of posting to Slack")
    args = ap.parse_args()

    sys.exit(run_once(headless=not args.headful, dry_run=args.dry_run))
