"""
GroundCloud vs FedEx Pickup Manifest reconciliation.

Runs ONCE per invocation (triggered hourly by a systemd timer -- see
systemd/groundcloud-monitor.timer in this folder). Deliberately has NO shared
code with fedex_monitor.py/pickup_risk.py one directory up -- no imports from
them, no shared state.db, no shared browser session -- so nothing here can
break that service. The only thing this script reads from the FedEx side is
the most recent .xls file FedEx's own --watch service already downloads every
cycle into ../downloads/ (read-only), so this script never logs into FedEx
at all -- only into GroundCloud.

WHAT THIS SCRIPT DOES EACH RUN
  1. Logs into GroundCloud (GROUNDCLOUD_USERNAME/GROUNDCLOUD_PASSWORD) -- a
     single-page username+password form, not a multi-step SSO flow like
     FedEx's Okta.
  2. Goes to P&D Time Commits, opens Stop Filters, clears whatever's
     pre-selected by default, selects "Pickups" only, applies.
  3. Scrapes every row across every page (Address, Status, ...) -- there's no
     Excel export button on this screen, just a paginated on-screen table, so
     this builds the equivalent of one itself.
  4. Loads the most recent FedEx Pickup Manifest .xls already sitting in
     ../downloads/ for today.
  5. Matches each GroundCloud pickup row to a FedEx manifest row by
     normalized address (no shared ID exists between the two systems).
  6. Flags the ONE-WAY mismatch only: GroundCloud status is Completed, but
     the matching FedEx row's "PU Closed" is still blank. (The reverse --
     FedEx PU Closed filled in but GroundCloud still Pending -- is NOT
     flagged; that's a normal, expected lag, not a gap.)
  7. Posts the result to its own Slack channel (SLACK_GC_CHANNEL) via the
     Slack Web API, reusing the same SLACK_BOT_TOKEN the Daily Manifest
     channel uses -- a fresh message every run (a table of flagged
     mismatches, or a one-line all-clear), not edited in place.

LIVE-VERIFIED SELECTORS: NONE YET. This was written from screenshots only --
the sandbox this was built in has no network access to groundcloud.io to test
against. Every DOM lookup below tries a couple of reasonable selectors and,
on total failure, dumps what it actually found on screen instead of failing
silently -- same pattern fedex_monitor.py uses -- so the first real run on
the VPS should be a quick one-line selector fix, not a mystery, if anything
doesn't match first try.
"""
from __future__ import annotations

import argparse
import difflib
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

HERE = Path(__file__).resolve().parent
FEDEX_DOWNLOADS_DIR = Path(os.environ.get("FEDEX_DOWNLOADS_DIR", str(HERE.parent / "downloads")))

# Same assumption fedex_monitor.py makes -- Delaware is Eastern time.
STATION_TZ = ZoneInfo("America/New_York")

GC_BASE = "https://groundcloud.io"
LOGIN_URL = f"{GC_BASE}/dashboard/login/"
TIME_COMMITS_URL = f"{GC_BASE}/dashboard/stop_service_times/"

GROUNDCLOUD_USERNAME = os.environ.get("GROUNDCLOUD_USERNAME", "")
GROUNDCLOUD_PASSWORD = os.environ.get("GROUNDCLOUD_PASSWORD", "")
SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN", "")
SLACK_GC_CHANNEL = os.environ.get("SLACK_GC_CHANNEL", "")

ADDRESS_MATCH_THRESHOLD = 0.85  # difflib ratio -- see match_address()


def log(msg: str) -> None:
    print(f"[{datetime.now(STATION_TZ):%Y-%m-%d %H:%M:%S %Z}] {msg}")


# --------------------------------------------------------------------- Slack

def post_to_slack(text: str, blocks: list[dict], dry_run: bool) -> bool:
    """One-shot post (chat.postMessage) -- this channel gets a fresh message
    every run, never edited in place, so there's no ts/state to track."""
    if dry_run or not SLACK_BOT_TOKEN or not SLACK_GC_CHANNEL:
        print(f"[dry-run / not configured] would POST to {SLACK_GC_CHANNEL or '(no channel set)'}:\n{text}\n")
        return False
    resp = requests.post(
        "https://slack.com/api/chat.postMessage",
        headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        json={"channel": SLACK_GC_CHANNEL, "text": text, "blocks": blocks},
        timeout=15,
    )
    data = resp.json()
    if not data.get("ok"):
        print(f"Slack chat.postMessage failed: {data}", file=sys.stderr)
        return False
    return True


# ------------------------------------------------------------- address match

_ABBR = {
    "STREET": "ST", "ROAD": "RD", "AVENUE": "AVE", "DRIVE": "DR", "LANE": "LN",
    "COURT": "CT", "PARKWAY": "PKWY", "BOULEVARD": "BLVD", "HIGHWAY": "HWY",
    "PLACE": "PL", "CIRCLE": "CIR", "TRAIL": "TRL", "SUITE": "STE",
    "APARTMENT": "APT", "BUILDING": "BLDG", "NORTH": "N", "SOUTH": "S",
    "EAST": "E", "WEST": "W",
}


def normalize_address(addr: str) -> str:
    """Uppercase, strip punctuation/zip/extra whitespace, normalize a handful
    of common street-suffix spellings -- GroundCloud and FedEx format
    addresses slightly differently (e.g. 'Rd' vs 'Road'), so an exact string
    match would miss plenty of genuine matches."""
    if not addr:
        return ""
    s = addr.upper()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\b\d{5}(\d{4})?\b$", "", s)  # trailing zip / zip+4 (no dash, already stripped)
    s = re.sub(r"\s+", " ", s).strip()
    words = [_ABBR.get(w, w) for w in s.split()]
    return " ".join(words)


def _score(gc_norm: str, fx_norm: str) -> float:
    """difflib ratio, or token containment when the house number is identical.
    GroundCloud often appends a business name to the street ('6716 CHURCH HILL RD
    RAMSEY PONTIAC, CHESTERTOWN') that FedEx omits, which drops the plain ratio
    to ~0.7-0.8 even though every FedEx token is present."""
    ratio = difflib.SequenceMatcher(None, gc_norm, fx_norm).ratio()
    g, f = gc_norm.split(), fx_norm.split()
    if g and f and g[0] == f[0]:
        contain = len(set(g) & set(f)) / min(len(set(g)), len(set(f)))
        ratio = max(ratio, contain)
    return ratio


def best_match(gc_address: str, fedex_rows: list[dict]) -> tuple[Optional[dict], float]:
    """Closest FedEx row (any score) and its similarity -- used by match_address()
    and by the per-row diagnostics logged on every run."""
    gc_norm = normalize_address(gc_address)
    if not gc_norm:
        return None, 0.0
    best, best_ratio = None, 0.0
    for row in fedex_rows:
        ratio = _score(gc_norm, row["_addr_norm"])
        if ratio > best_ratio:
            best, best_ratio = row, ratio
    return best, best_ratio


def match_address(gc_address: str, fedex_rows: list[dict]) -> Optional[dict]:
    """Best fuzzy match of a GroundCloud address against every FedEx row's
    normalized address. Returns None below ADDRESS_MATCH_THRESHOLD -- an
    unmatched GroundCloud row is simply skipped, not flagged (per spec: only
    flag when a FedEx match IS found and PU Closed is blank there)."""
    best, best_ratio = best_match(gc_address, fedex_rows)
    return best if best is not None and best_ratio >= ADDRESS_MATCH_THRESHOLD else None


# ------------------------------------------------------------- fedex import

def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() == ""


def find_latest_fedex_export(now: datetime) -> Optional[Path]:
    """FedEx's own --watch service writes manifest_YYYYMMDD_HHMMSS.xls into
    ../downloads/ every full-check cycle (every 15 min by default) -- this
    just picks today's most recent one. Read-only: never writes there."""
    pattern = f"manifest_{now:%Y%m%d}_*.xls"
    matches = sorted(FEDEX_DOWNLOADS_DIR.glob(pattern))
    return matches[-1] if matches else None


def load_fedex_rows(path: Path) -> list[dict]:
    """Minimal, independent read of the same 'Stop Details' sheet
    pickup_risk.py reads -- duplicated here on purpose (a few lines) rather
    than imported, to keep this script's only dependency on the FedEx side
    being the downloaded file itself, not its code."""
    df = pd.read_excel(path, sheet_name="Stop Details", engine="xlrd", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, row in df.iterrows():
        address = ", ".join(x for x in (
            str(row.get("Address Line 1") or "").strip(),
            str(row.get("City") or "").strip(),
            str(row.get("State") or "").strip(),
        ) if x)
        rows.append({
            "station": row.get("Station"),
            "shipper": str(row.get("Shipper Name") or "").strip(),
            "address": address,
            "pu_closed_blank": _is_blank(row.get("PU Closed")),
            "_addr_norm": normalize_address(address),
        })
    return rows


# ---------------------------------------------------------------- scraping

def login(page, log) -> bool:
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    # This looks like a client-rendered (JS) app -- domcontentloaded alone can
    # fire before the login form itself has actually rendered. Give it a real
    # chance to finish, then wait specifically for ANY input to show up
    # before even trying the specific field selectors below.
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except PWTimeout:
        log("Page didn't reach networkidle within 20s (ok if it's just a slow "
            "background request) -- continuing anyway.")
    try:
        page.wait_for_selector("input", timeout=15000)
    except PWTimeout:
        log("No <input> elements appeared within 15s of loading the login page -- "
            "see the diagnostic dump below for what IS on the page.")

    user_candidates = [
        lambda: page.locator('input[name="username"]').first,
        lambda: page.locator('input[type="email"]').first,
        lambda: page.locator('input[id*="user" i]').first,
        lambda: page.locator('input[placeholder*="user" i], input[placeholder*="email" i]').first,
        lambda: page.locator('input[type="text"]').first,
    ]
    pass_candidates = [
        lambda: page.locator('input[type="password"]').first,
    ]

    user_input = next((loc() for loc in user_candidates if loc().count() > 0), None)
    pass_input = next((loc() for loc in pass_candidates if loc().count() > 0), None)

    if user_input is None or pass_input is None:
        _dump_visible_elements(page, log, "Could not find the username/password fields on the login page.")
        return False

    user_input.fill(GROUNDCLOUD_USERNAME)
    pass_input.fill(GROUNDCLOUD_PASSWORD)

    submit_candidates = [
        lambda: page.get_by_role("button", name=re.compile(r"next|log\s*in|sign\s*in", re.I)).first,
        lambda: page.locator('button[type="submit"]').first,
    ]
    submit = next((loc() for loc in submit_candidates if loc().count() > 0), None)
    if submit is None:
        _dump_visible_elements(page, log, "Could not find the login submit button.")
        return False

    submit.click()
    try:
        page.wait_for_url(re.compile(r"/dashboard/?(?!.*login)"), timeout=20000)
    except PWTimeout:
        if "/login" in page.url:
            log(f"Still on the login page after submitting -- login failed "
                f"(wrong credentials, or the page needs a different selector). URL: {page.url}")
            return False
        # Some other URL than expected but not the login page -- probably fine, keep going.
    log(f"Logged into GroundCloud. Landed on: {page.url}")
    return True


def apply_pickups_filter(page, log) -> bool:
    page.goto(TIME_COMMITS_URL, wait_until="domcontentloaded")
    if "/login" in page.url:
        log("Redirected back to login when opening P&D Time Commits -- session isn't valid.")
        return False

    filters_btn = page.get_by_role("button", name=re.compile(r"stop filters", re.I)).first
    if filters_btn.count() == 0:
        _dump_visible_elements(page, log, "Could not find the 'Stop Filters' button.")
        return False
    filters_btn.click()

    try:
        page.wait_for_selector("text=/stop filters/i", timeout=10000)
    except PWTimeout:
        log("Stop Filters panel didn't visibly open -- continuing anyway in case it's just slow.")

    clear_btn = page.get_by_role("button", name=re.compile(r"^clear$", re.I)).first
    if clear_btn.count() > 0:
        clear_btn.click()
    else:
        log("No 'Clear' button found in Stop Filters -- continuing without clearing "
            "(filters may already be at defaults).")

    pickups_btn = page.get_by_role("button", name=re.compile(r"^pickups$", re.I)).first
    if pickups_btn.count() == 0:
        _dump_visible_elements(page, log, "Could not find the 'Pickups' filter toggle.")
        return False
    pickups_btn.click()
    log("Stop Filters: clicked Clear, then Pickups.")

    done_btn = page.get_by_role("button", name=re.compile(r"^done$", re.I)).first
    if done_btn.count() > 0:
        done_btn.click()
    else:
        log("No 'Done' button found -- the filter panel may auto-apply, continuing.")

    log("Stop Filters: clicked Done (or auto-applied).")
    page.wait_for_timeout(1500)  # let the table re-render
    return True


def _dump_visible_elements(page, log, context: str) -> None:
    try:
        log(f"{context} Page title: {page.title()!r} -- URL: {page.url}")
        counts = page.evaluate(
            "() => ({input: document.querySelectorAll('input').length, "
            "button: document.querySelectorAll('button').length, "
            "iframe: document.querySelectorAll('iframe').length, "
            "bodyTextLen: (document.body.innerText||'').length})"
        )
        log(f"{context} Raw element counts on the page: {counts}")
        if counts.get("iframe"):
            log(f"{context} NOTE: page has {counts['iframe']} iframe(s) -- the login "
                "form may be INSIDE one of those, not on the top-level page.")
        # Don't filter by text/name/placeholder here -- an <input> with none of
        # those is exactly the case we need to see, not hide.
        elements = page.locator("button, a, input, span[role], div[role]").evaluate_all(
            "els => els.filter(e => e.offsetParent !== null).slice(0, 50).map(e => ({"
            "tag: e.tagName, text: (e.innerText||'').trim().slice(0,40), "
            "type: e.type||'', name: e.name||'', id: e.id||'', "
            "placeholder: e.placeholder||''}))"
        )
        log(f"{context} Visible interactive elements on screen "
            f"({len(elements)} shown, up to 50):\n  " + "\n  ".join(str(e) for e in elements))
    except Exception as e:
        log(f"{context} (diagnostic dump itself failed: {e!r})")


def scrape_pickups(page, log) -> list[dict]:
    """Scrapes every row across every page of the (already Pickups-filtered)
    P&D Time Commits table. Column order assumed from the live screenshot:
    Route, Stop Type, SID, Address, Open, Close, Stop Sequence, Time, Status,
    Actions -- adjusted here if the real table differs once this runs live."""
    rows: list[dict] = []
    seen_pages = 0
    max_pages = 200  # safety valve against an infinite Next loop

    while seen_pages < max_pages:
        seen_pages += 1
        try:
            page.wait_for_selector("table tbody tr", timeout=15000)
        except PWTimeout:
            log(f"No table rows found on page {seen_pages} -- stopping pagination here.")
            break

        page_rows = page.locator("table tbody tr")
        count = page_rows.count()
        for i in range(count):
            cells = page_rows.nth(i).locator("td").all_inner_texts()
            if len(cells) < 9:
                continue  # header/spacer row or layout surprise -- skip rather than guess
            rows.append({
                "route": cells[0].strip(),
                "stop_type": cells[1].strip(),
                "sid": cells[2].strip(),
                "address": cells[3].strip(),
                "open": cells[4].strip(),
                "close": cells[5].strip(),
                "stop_sequence": cells[6].strip(),
                "time": cells[7].strip(),
                "status": cells[8].strip(),
            })
        log(f"Scraped {count} row(s) from page {seen_pages} ({len(rows)} total so far).")

        # Live DOM (screenshot 2026-10-21): <li class="page-item page-next"><button
        # role="menuitem" aria-label="Go to next page" class="page-link">Next</button></li>.
        # role=menuitem, so get_by_role("button") never matched it. On the last page the
        # <li> gains class "disabled" (same as page-previous on page 1).
        next_btn = page.locator('li.page-next button.page-link, button[aria-label="Go to next page"]').first
        if next_btn.count() == 0:
            log("No 'Next' button found -- treating this as the last page.")
            break
        li_disabled = page.locator("li.page-next.disabled").count() > 0
        if li_disabled or next_btn.is_disabled():
            log("'Next' is disabled -- that was the last page.")
            break
        first_before = page.locator("table tbody tr").first.inner_text()
        next_btn.click()
        page.wait_for_timeout(1500)
        if page.locator("table tbody tr").first.inner_text() == first_before:
            page.wait_for_timeout(2500)  # slow render -- give it one more chance
            if page.locator("table tbody tr").first.inner_text() == first_before:
                log("Clicked Next but the table didn't change -- stopping to avoid a loop.")
                break

    return rows


# ------------------------------------------------------------------- report

# GroundCloud's table says "Delivered" for a finished pickup (seen live 2026-10-21;
# the original code only looked for "Completed" so it could never flag anything).
GC_DONE_STATUSES = {"DELIVERED", "COMPLETED"}


def find_mismatches(gc_rows: list[dict], fedex_rows: list[dict]) -> list[dict]:
    """Per FedEx address: GroundCloud shows N finished pickups there, FedEx has C
    rows with PU Closed filled and O still blank. Many stops share an address
    (e.g. 14 pickups at one street address), so matching one row to one row by
    address would pick an arbitrary FedEx row. Counting is order-independent:
    flag when GroundCloud has finished more than FedEx has closed AND FedEx still
    has open rows there; the number flagged is min(N - C, O)."""
    done_by_addr: dict[str, list[dict]] = {}
    for gc in gc_rows:
        if gc["status"].strip().upper() not in GC_DONE_STATUSES:
            continue
        m = match_address(gc["address"], fedex_rows)
        if m is not None:
            done_by_addr.setdefault(m["_addr_norm"], []).append(gc)

    flagged = []
    for key, gcs in done_by_addr.items():
        fx = [r for r in fedex_rows if r["_addr_norm"] == key]
        closed = sum(1 for r in fx if not r["pu_closed_blank"])
        open_rows = [r for r in fx if r["pu_closed_blank"]]
        n = min(len(gcs) - closed, len(open_rows))
        if n > 0:
            flagged.append({
                "address": gcs[0]["address"],
                "fedex_shipper": next((r["shipper"] for r in open_rows if r["shipper"]), ""),
                "fedex_station": open_rows[0]["station"],
                "gc_done": len(gcs), "fx_closed": closed, "fx_open": len(open_rows),
                "n": n,
            })
    flagged.sort(key=lambda f: f["address"])
    return flagged


def build_report(gc_rows: list[dict], fedex_rows: list[dict], now: datetime) -> tuple[str, list[dict]]:
    flagged = find_mismatches(gc_rows, fedex_rows)
    total = sum(f["n"] for f in flagged)

    summary = (f"*🔀 GroundCloud vs FedEx — Pickup Reconciliation*\n"
               f"Checked {now:%I:%M %p %Z} · {len(gc_rows)} GroundCloud pickup(s) scanned · "
               f"{total} mismatch(es) at {len(flagged)} address(es)")

    if not flagged:
        text = "GroundCloud/FedEx reconciliation: no mismatches this run."
        blocks = [{"type": "section", "text": {"type": "mrkdwn", "text": summary + "\n✅ No mismatches found."}}]
        return text, blocks

    header = f"{'STN':<7}{'FEDEX SHIPPER':<24}{'ADDRESS':<32}{'GC DONE':<9}{'FX CLOSED':<11}{'OPEN'}"
    lines = [header]
    for f in flagged:
        shipper = (f["fedex_shipper"] or "Unknown")[:22]
        address = (f["address"] or "")[:30]
        stn = str(int(f["fedex_station"])) if str(f["fedex_station"]).replace(".0", "").isdigit() else ""
        lines.append(f"{stn:<7}{shipper:<24}{address:<32}{f['gc_done']:<9}{f['fx_closed']:<11}{f['fx_open']}")
    table = "\n".join(lines)

    text = f"GroundCloud/FedEx reconciliation: {total} mismatch(es)"
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": summary}},
        {"type": "section", "text": {"type": "mrkdwn", "text": f"```{table}```"}},
        {"type": "context", "elements": [{"type": "mrkdwn",
            "text": "GroundCloud shows these pickups *Delivered*, but FedEx's *PU Closed* is still blank for stops at that address."}]},
    ]
    return text, blocks


# --------------------------------------------------------------------- main

def run(headless: bool, dry_run: bool) -> int:
    now = datetime.now(STATION_TZ)

    if not GROUNDCLOUD_USERNAME or not GROUNDCLOUD_PASSWORD:
        log("GROUNDCLOUD_USERNAME/GROUNDCLOUD_PASSWORD not set -- nothing to do.")
        return 1

    fedex_path = find_latest_fedex_export(now)
    if fedex_path is None:
        log(f"No FedEx export found for today ({now:%Y-%m-%d}) in {FEDEX_DOWNLOADS_DIR} -- "
            "skipping this run (the FedEx monitor may not have completed a check yet).")
        return 1
    log(f"Using FedEx export: {fedex_path.name}")
    fedex_rows = load_fedex_rows(fedex_path)

    with sync_playwright() as p:
        # groundcloud.io answers plain curl with 200 but returned 403 to Playwright's
        # default headless browser (2026-10-21), whose user agent says "HeadlessChrome".
        # Present as a normal desktop Chrome instead (full chromium build, real UA,
        # normal viewport).
        browser = p.chromium.launch(headless=headless, channel="chromium")
        ctx = browser.new_context(
            user_agent=("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/131.0.0.0 Safari/537.36"),
            viewport={"width": 1366, "height": 768},
        )
        page = ctx.new_page()
        try:
            if not login(page, log):
                return 1
            if not apply_pickups_filter(page, log):
                return 1
            gc_rows = scrape_pickups(page, log)
        finally:
            browser.close()

    log(f"Scraped {len(gc_rows)} GroundCloud pickup row(s) total.")
    # Per-row evidence (log only, never Slack): a silent "no mismatches" is only
    # trustworthy if the addresses really matched, so show every row's best match.
    log(f"FedEx manifest rows loaded: {len(fedex_rows)}")
    from collections import Counter
    log(f"GC stop types scraped: {dict(Counter(r['stop_type'] for r in gc_rows))}; "
        f"statuses: {dict(Counter(r['status'] for r in gc_rows))}")
    for gc in gc_rows:
        m, ratio = best_match(gc["address"], fedex_rows)
        log(f"  GC sid={gc['sid']} status={gc['status']!r} addr={gc['address']!r} -> "
            f"best ratio={ratio:.2f} "
            + (f"FedEx addr={m['address']!r} PU-closed-blank={m['pu_closed_blank']}" if m else "no candidate"))
    text, blocks = build_report(gc_rows, fedex_rows, now)
    post_to_slack(text, blocks, dry_run)
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--headful", action="store_true", help="show the browser window for this run")
    ap.add_argument("--dry-run", action="store_true", help="print the report instead of posting to Slack")
    args = ap.parse_args()
    sys.exit(run(headless=not args.headful, dry_run=args.dry_run))
