#!/usr/bin/env bash
# One-shot setup of the FedEx monitor + GroundCloud reconciliation on a fresh Ubuntu server.
#   1. git clone <repo> /root/fedex      (the repo is private: use a GitHub token when asked)
#   2. bash /root/fedex/deploy/setup.sh   -> creates .env from the template and stops
#   3. nano /root/fedex/.env              -> fill in the secrets
#   4. bash /root/fedex/deploy/setup.sh   -> installs everything and starts the services
# Safe to re-run. Does not touch Docker/EMR or anything outside /root/fedex and two systemd units.
set -euo pipefail
APP=/root/fedex
cd "$APP"

echo "== system packages"
apt-get update -y
apt-get install -y git python3 python3-venv python3-pip

echo "== python environment"
[ -d venv ] || python3 -m venv venv
venv/bin/pip install --upgrade pip
venv/bin/pip install -r requirements.txt

echo "== chromium for playwright"
venv/bin/playwright install --with-deps chromium

if [ ! -f .env ]; then
  cp deploy/env.example .env
  chmod 600 .env
  echo
  echo "Created $APP/.env from the template. Fill it in (nano $APP/.env), then run this script again."
  exit 0
fi
chmod 600 .env
if grep -q -E '^(FEDEX_USERNAME|FEDEX_PASSWORD)=$' .env; then
  echo "FEDEX_USERNAME / FEDEX_PASSWORD are still empty in .env -- fill them in and re-run."
  exit 1
fi

mkdir -p downloads

echo "== systemd units"
cp deploy/fedex-monitor.service /etc/systemd/system/
cp groundcloud/systemd/groundcloud-monitor.service groundcloud/systemd/groundcloud-monitor.timer /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now fedex-monitor.service
systemctl enable --now groundcloud-monitor.timer

echo
systemctl is-active fedex-monitor.service groundcloud-monitor.timer
echo "Watch it: journalctl -u fedex-monitor.service -f | cut -c1-200 | grep --line-buffered -v Set-Cookie"
