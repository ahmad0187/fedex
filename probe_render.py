"""
One-off diagnostic (NOT used by the service): tries several browser configurations
against the FedEx Pickup Manifest page, one at a time, and reports which ones get the
manifest form to render. Created 2026-10-21 after the form failed to render on every
automated attempt (including fresh logins after 10+ idle hours) while a manual login in
a real browser worked every time.

Usage (stop the service first so nothing else is logging in as this account):
    systemctl stop fedex-monitor.service
    cd ~/fedex && set -a && source .env && set +a
    ~/fedex/venv/bin/python3 probe_render.py            # headless variants
    xvfb-run -a ~/fedex/venv/bin/python3 probe_render.py --headed   # + headed variant

Each variant signs out and waits before the next one so they don't collide.
"""
import argparse
import time
from datetime import datetime

from playwright.sync_api import sync_playwright

import fedex_monitor as fm

DESKTOP_UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/131.0.0.0 Safari/537.36")
HIDE_WEBDRIVER = "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"

VARIANTS = [
    # name, headless, launch args, context kwargs, init script
    ("A baseline (current setup)", True, [], {}, None),
    ("B desktop UA + 1366x768 viewport", True, [],
     {"user_agent": DESKTOP_UA, "viewport": {"width": 1366, "height": 768}}, None),
    ("C desktop UA + viewport + hide automation flags", True,
     ["--disable-blink-features=AutomationControlled"],
     {"user_agent": DESKTOP_UA, "viewport": {"width": 1366, "height": 768}}, HIDE_WEBDRIVER),
]
HEADED_VARIANT = ("D headed (real window via xvfb) + desktop UA + viewport", False,
                  ["--disable-blink-features=AutomationControlled"],
                  {"user_agent": DESKTOP_UA, "viewport": {"width": 1366, "height": 768}},
                  HIDE_WEBDRIVER)


def log(msg):
    print(f"[{datetime.now(fm.STATION_TZ):%H:%M:%S}] {msg}", flush=True)


def run_variant(p, name, headless, args, ctx_kwargs, init_script):
    log(f"===== VARIANT {name} =====")
    browser = p.chromium.launch(headless=headless, channel="chromium", args=args)
    ctx = browser.new_context(accept_downloads=True, **ctx_kwargs)
    if init_script:
        ctx.add_init_script(init_script)
    page = ctx.new_page()
    try:
        ua = page.evaluate("navigator.userAgent")
        log(f"user agent in use: {ua}")
    except Exception:
        pass
    result = fm.goto_pickup_manifest(ctx, page, log)
    ok = result not in (None, "LOGGED_OUT")
    log(f"RESULT for {name}: {'FORM RENDERED, Search worked' if ok else 'FAILED'}")
    for pg in list(ctx.pages):
        fm.attempt_logout(pg, log)
    browser.close()
    return ok


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--headed", action="store_true", help="also try a headed variant (run under xvfb-run)")
    ap.add_argument("--only", type=int, default=0, help="run just the first N variants (1 = a single login)")
    ap.add_argument("--pause", type=int, default=90, help="seconds to wait between variants")
    args = ap.parse_args()

    variants = list(VARIANTS) + ([HEADED_VARIANT] if args.headed else [])
    if args.only:
        variants = variants[:args.only]
    results = {}
    with sync_playwright() as p:
        for i, v in enumerate(variants):
            if i:
                log(f"waiting {args.pause}s so the previous session can wind down...")
                time.sleep(args.pause)
            results[v[0]] = run_variant(p, *v)

    print("\n==== SUMMARY ====")
    for name, ok in results.items():
        print(f"{'OK  ' if ok else 'FAIL'}  {name}")
