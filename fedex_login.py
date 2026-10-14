"""
Run this LOCALLY (on your own computer), not in GitHub Actions. It opens a real,
visible browser so YOU log in to FedEx yourself -- this script never sees, asks for,
or stores your password. Once you're logged in, it saves the session (cookies only)
to ./storage_state.json, which is what fedex_monitor.py loads on every scheduled run.

Usage
    pip install playwright
    playwright install chromium
    python fedex_login.py

Then follow README.md to put storage_state.json's contents into a GitHub Actions
secret (FEDEX_STORAGE_STATE) so the scheduled workflow can use it.

Re-run this whenever fedex_monitor.py reports the session has expired.
"""
from pathlib import Path

from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
STORAGE_STATE_PATH = HERE / "storage_state.json"
HOMEPAGE_URL = "https://mybizaccount.fedex.com/"

print("Opening a browser window. Log in to FedEx yourself (handle MFA if asked), then "
      "click through to FedEx Customer Connection -> P&D Manifests -> Pickup Manifest, "
      "set Work Area to ALL, pick today's date and click Search once so the session is "
      "fully established. Then come back here and press Enter.")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    ctx = browser.new_context()
    page = ctx.new_page()
    page.goto(HOMEPAGE_URL)
    input("\nPress Enter once you're logged in and on the Pickup Manifest page... ")
    ctx.storage_state(path=str(STORAGE_STATE_PATH))
    browser.close()

print(f"Saved session to {STORAGE_STATE_PATH}.")
print("Next: follow README.md to add its contents as the FEDEX_STORAGE_STATE GitHub secret.")
