# FedEx Pickup Manifest monitor (GitHub Actions)

Checks the Pickup Manifest every 15 minutes and posts to Slack when a stop is about
to miss its pickup close time: `PU Closed` (delivered time) empty AND `Reas Code`
empty AND close time within 60 minutes. Runs for free on GitHub Actions — nothing
needs to stay open on your end.

## Files

- `pickup_risk.py` — the rule itself (parses the manifest export, decides which stops
  are at risk). Tested against a real exported manifest.
- `fedex_monitor.py` — what each scheduled run does: loads the saved login session,
  exports the manifest, calls `pickup_risk.py`, tracks what's already been alerted
  (`state.db`), posts to Slack. All selectors in it were confirmed against a real
  logged-in FedEx session with real data.
- `fedex_login.py` — run this **locally only**, once, to log in and save the session.
- `.github/workflows/fedex-monitor.yml` — the schedule (every 15 min) and the steps
  each run follows.

## Setup

### 1. Create a private GitHub repo and push these files

Keep it **private** — not because the code is sensitive, but because the saved
session secret and dedup state live alongside it. Push everything in this folder,
including `.github/workflows/fedex-monitor.yml` and `.gitignore`, to the repo root
(not nested in a subfolder).

### 2. Get a Slack webhook URL (2 minutes, free)

In Slack: **Apps → search "Incoming WebHooks" → Add to Slack → choose the channel →
Copy the Webhook URL.**

### 3. Choose how the script logs in

**Option A — auto-login (fully hands-off, recommended if you're fine with it holding
a login secret):** the script logs in itself every run using a username/password you
give it. Add these as GitHub secrets — **paste them directly into GitHub's own secret
form, never into a chat with me or anyone else**:

- `FEDEX_USERNAME`
- `FEDEX_PASSWORD`

Caveat: if FedEx ever shows an MFA/one-time-code screen on a run (new IP, unfamiliar
login, etc.), the script can't click through that — it'll fail that cycle and alert
you once (see "When login fails" below), and you'd log in manually once to clear it.
I built the login-form selectors from standard patterns, not by watching FedEx's
actual login page live (I don't fill in real login forms myself, including for
testing) — so the **first real run matters**: check its logs, and if it can't find
the right field it prints every input field on screen so the fix is a one-line
selector change, not guesswork.

**Option B — saved session (no password stored anywhere):** on your own computer
(not in GitHub Actions):

```bash
pip install playwright
playwright install chromium
python fedex_login.py
```

A real browser window opens. Log in to FedEx yourself (handle MFA if asked), click
through to **FedEx Customer Connection → P&D Manifests → Pickup Manifest**, set
Work Area to **ALL**, pick today's date, click **Search** once, then come back to the
terminal and press **Enter**. This saves `storage_state.json` next to the script —
cookies only, never your password. Add its contents as the `FEDEX_STORAGE_STATE`
secret. Trade-off: the session eventually expires and needs a manual refresh (see
"When login fails" below); nothing auto-recovers on its own.

You can set up both — the script tries the saved session first, falls back to
auto-login if the session's expired and credentials are set.

### 4. Add the Slack secret

In your repo: **Settings → Secrets and variables → Actions → New repository secret.**

- `SLACK_WEBHOOK_URL` — the webhook URL from step 2.

### 5. Test it manually before trusting the schedule

Repo → **Actions** tab → **FedEx Pickup Manifest Monitor** → **Run workflow**. Watch
it run, then check the logs and your Slack channel. If it fails at the export step,
the logs will print every icon/button it found on that screen instead of guessing —
send me that output and it's a one-line fix.

### 6. Let the schedule take over

Once a manual run works, the `*/15 * * * *` cron in the workflow takes it from there —
nothing to start or leave running. GitHub's scheduler is best-effort (can slip a few
minutes under load), which doesn't matter for a 60-minute warning window.

## How re-alerting works

A stop is alerted once when it first enters the 60-minute window, and again only if
it escalates from 🟠 AT_RISK to ⛔ PAST_CLOSE — not every 15 minutes while it sits
unresolved. That memory (`state.db`) is committed back to the repo by the workflow
after each run, since GitHub Actions runners don't keep any disk between runs.

## When login fails

If auto-login is set up (`FEDEX_USERNAME`/`FEDEX_PASSWORD`) and it can't get past the
login screen — wrong selector, or FedEx showing an MFA/OTP challenge — or if you're
on the saved-session approach and that session's expired, every run fails the same
way, and the workflow posts **one** Slack warning (capped to once per 2 hours, not
every 15 minutes) saying so. Check the Action run's logs for specifics (missing
field names get printed there). Fix depends on which mode: auto-login usually means
a selector needs a tweak (send me the printed fields) or you need to log in manually
once to clear an MFA prompt; saved-session means re-running `fedex_login.py` and
updating the `FEDEX_STORAGE_STATE` secret.

## Things worth deciding later

- **Timezone**: assumes Delaware/Eastern (`America/New_York`) — the manifest's
  Ready/Close times have no timezone marker, so this is an assumption, not something
  FedEx tells us.
- **"PU Closed" = "delivered time"** — the field your brother described. If that's
  not quite right, the trigger logic lives in one place (`pickup_risk.py`,
  `find_at_risk()`) and is easy to point at a different column.
- Native **Pickup Alerts** tab in FCC was checked — it only covers SDOC/transferred
  pickups, not close-time risk, so it doesn't replace this.
- GitHub disables a repo's scheduled workflows after 60 days with **zero** repo
  activity — the bot's own `state.db` commits every run count as activity, so this
  shouldn't actually happen in practice, but worth knowing if it ever goes quiet.
