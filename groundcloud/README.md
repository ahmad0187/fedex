# GroundCloud / FedEx reconciliation

Catches the gap where a driver marks a pickup **Completed** on GroundCloud but
never closes it out on the FedEx side (`PU Closed` stays blank). Runs
**hourly**, completely independent of the FedEx `--watch` service one
directory up: separate script, separate systemd service, no shared code, no
shared state. The only thing it touches from the FedEx side is the most
recent `.xls` the FedEx service already downloads every 15 minutes into
`../downloads/` (read-only) -- so this never needs its own FedEx login.

## What it does each run

1. Logs into GroundCloud (single-page username+password form).
2. Opens **P&D Time Commits** → **Stop Filters** → Clear → **Pickups** → Done.
3. Scrapes every row across every page of that table (no Excel export button
   on this screen, so this builds the equivalent of one itself).
4. Loads today's latest FedEx manifest export.
5. Matches each GroundCloud pickup to a FedEx row by normalized address (fuzzy
   match -- the two systems format addresses slightly differently).
6. Flags it **only** when GroundCloud says Completed and the matched FedEx
   row's `PU Closed` is still blank. (Not the reverse.)
7. Edits ONE Slack message in place every run (all-clear or table).
   flagged mismatches, or a one-line all-clear if there are none.

## Setup

### 1. Add these to the SAME `.env` the FedEx service already reads
(`/root/fedex/.env` on the VPS -- no new secrets file needed):

```
GROUNDCLOUD_USERNAME=...
GROUNDCLOUD_PASSWORD=...
SLACK_GC_CHANNEL=C0CAS5GCS1Y
```

`SLACK_BOT_TOKEN` is reused from the Daily Manifest setup -- make sure the
same Slack bot is also invited into the GroundCloud comparison channel
(`/invite @YourAppName` there), since a Bot Token only works in channels the
bot has actually been added to.

### 2. Install the systemd timer

```bash
cp /root/fedex/groundcloud/systemd/groundcloud-monitor.service /etc/systemd/system/
cp /root/fedex/groundcloud/systemd/groundcloud-monitor.timer /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now groundcloud-monitor.timer
```

Check it's scheduled: `systemctl list-timers groundcloud-monitor.timer`

### 3. Test it manually first

```bash
cd /root/fedex/groundcloud
python3 groundcloud_monitor.py --dry-run --headful
```

`--dry-run` prints the report instead of posting to Slack. `--headful` shows
the actual browser window (if you're on a desktop session / VNC) -- drop it
to run headless, which is what the timer uses.

## Known gap: untested selectors

This was built from screenshots only -- the environment it was written in
has no network access to groundcloud.io, so none of the login/filter/scrape
steps have been run against the real site yet. Every DOM lookup has a
fallback that prints what it actually found on screen if a selector doesn't
match, so if the first real run fails, send me that log output and it should
be a quick, targeted fix rather than a guessing game -- the same process that
got the FedEx side working originally.

## Tuning the address match

`ADDRESS_MATCH_THRESHOLD` in `groundcloud_monitor.py` (default `0.85`, a
`difflib` similarity ratio on normalized addresses) controls how close two
addresses need to be to count as the same stop. If real runs show missed
matches (too strict) or wrong matches (too loose), that's the one number to
adjust -- send me a couple of real examples and I'll tune it.
