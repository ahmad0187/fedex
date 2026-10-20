"""
Parses a FedEx "Pickup Manifest" export (the .xls downloaded from FedEx Customer
Connection, FCC Links > FedEx Customer Connection > Pickup Manifest > Excel export)
and flags stops at risk of missing their pickup close time.

Trigger (exactly as specified):
    PU Closed is empty              -- pickup not yet closed out
    AND Reas Code is empty          -- no reason on file for not picking up
    AND Close time is <= 1 hour away (and hasn't already passed)

This file has NO network/browser code and no alert-sending code on purpose: it is
the deterministic part, easy to test against a real export, independent of how the
file gets downloaded or where the alert gets sent.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, time as dtime
from pathlib import Path
from typing import Optional

import pandas as pd

SHEET = "Stop Details"

# Daily Manifest row order: at-risk (🟡) first since those need action soonest,
# then past-close/overdue (🔴), then not-yet-due pending (🔵), then completed
# (🟢) at the very bottom -- so the rows needing attention stay pinned to the
# top of the table on every re-sort instead of drifting by close time.
_STATUS_SORT_RANK = {"warn": 0, "overdue": 1, "pending": 2, "done": 3}

# A stop's unique key on one day's manifest. PU List is the pickup-group id and PUID
# is the line number within it (both repeat day to day), so the pair is what to
# dedupe/track alerts against -- not just PUID alone (PUID resets to 1 for every list).
StopKey = tuple[int, int]


@dataclass
class RiskyStop:
    key: StopKey
    station: int
    wa: str
    shipper_name: str
    address: str
    city: str
    state: str
    close: dtime
    ready: Optional[dtime]
    minutes_to_close: float
    status: str  # "AT_RISK" (<=60 min away) or "PAST_CLOSE" (close time already passed)


def _parse_hhmm(v) -> Optional[dtime]:
    """Ready/Close/PU Closed come through as 'HH:MM' strings (or blank/NaN)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if not s:
        return None
    h, m = s.split(":")
    return dtime(int(h), int(m))


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() == ""


def load_manifest(path: str | Path) -> pd.DataFrame:
    """Reads the 'Stop Details' sheet. Works on both the legacy .xls FedEx exports
    and .xlsx, in case the export format ever changes."""
    path = Path(path)
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(path, sheet_name=SHEET, engine=engine, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_at_risk(df: pd.DataFrame, now: datetime, warn_minutes: float = 60.0,
                 manifest_day: Optional[datetime] = None) -> list[RiskyStop]:
    """
    now            aware or naive datetime, in the SAME wall-clock time as the
                   Ready/Close columns (FedEx shows station-local time; see note below)
    warn_minutes   how far ahead of Close to start flagging (default 60, i.e. "1 hour prior")
    manifest_day   the calendar date the manifest is for (defaults to now's date) --
                   pass this explicitly when polling after midnight for a manifest
                   pulled before midnight, so Close doesn't get compared against the wrong day
    """
    day = (manifest_day or now).date()
    out: list[RiskyStop] = []

    for _, row in df.iterrows():
        pu_closed = row.get("PU Closed")
        reas_code = row.get("Reas Code")
        if not (_is_blank(pu_closed) and _is_blank(reas_code)):
            continue  # already closed out, or already has a reason on file

        close_t = _parse_hhmm(row.get("Close"))
        if close_t is None:
            continue  # no close time on this row at all -- nothing to check
        close_dt = datetime.combine(day, close_t, tzinfo=now.tzinfo)

        minutes_left = (close_dt - now).total_seconds() / 60.0
        if minutes_left > warn_minutes:
            continue  # not within the warning window yet
        status = "PAST_CLOSE" if minutes_left < 0 else "AT_RISK"

        out.append(RiskyStop(
            key=(int(row["PU List"]), int(row["PUID"])),
            station=int(row["Station"]),
            wa=str(row["WA"]),
            shipper_name=str(row.get("Shipper Name") or "").strip(),
            address=", ".join(x for x in (
                str(row.get("Address Line 1") or "").strip(),
                str(row.get("City") or "").strip(),
                str(row.get("State") or "").strip(),
            ) if x),
            city=str(row.get("City") or "").strip(),
            state=str(row.get("State") or "").strip(),
            close=close_t,
            ready=_parse_hhmm(row.get("Ready")),
            minutes_to_close=round(minutes_left, 1),
            status=status,
        ))

    out.sort(key=lambda s: s.minutes_to_close)
    return out


def format_alert(stop: RiskyStop) -> str:
    """Plain-text version -- used as the Slack notification's required fallback
    'text' (shown in previews/notifications where blocks don't render) and as the
    --dry-run printout. format_alert_blocks() below is the actual rich message."""
    icon = "⛔ PAST CLOSE" if stop.status == "PAST_CLOSE" else "🟠 PICKUP AT RISK"
    when = stop.close.strftime("%I:%M %p").lstrip("0")
    if stop.status == "PAST_CLOSE":
        timing = f"close time {when} has passed"
    else:
        timing = f"closes {when} ({stop.minutes_to_close:.0f} min left)"
    lines = [
        f"{icon}: {stop.shipper_name or 'Unknown shipper'} (Station {stop.station}, WA {stop.wa})",
        f"Pickup window {timing}. No pickup closed, no reason on file.",
    ]
    if stop.address:
        lines.append(f"📍 {stop.address}")
    return "\n".join(lines)


def format_at_risk_table(stops: list[RiskyStop], now: datetime) -> tuple[str, list[dict]]:
    """Slack message for the 'At Risk' channel: ONE table listing every stop
    currently in the warning window (not one message per stop). Caller re-sends
    this whenever the set of at-risk stops actually changes (a stop drops off once
    picked up, a new one is added once it enters the window) -- see
    fedex_monitor.py's at-risk-set diffing."""
    header = f"{'':<2}{'STN':<5}{'WA':<6}{'SHIPPER':<22}{'CLOSE':<7}{'LEFT':<9}"
    lines = [header]
    for s in stops:
        icon = "⛔" if s.status == "PAST_CLOSE" else "🟠"
        shipper = (s.shipper_name or "Unknown")[:21]
        close_s = s.close.strftime("%H:%M")
        left = (f"-{abs(s.minutes_to_close):.0f}m" if s.status == "PAST_CLOSE"
                else f"{s.minutes_to_close:.0f}m")
        lines.append(f"{icon:<2}{s.station:<5}{s.wa:<6}{shipper:<22}{close_s:<7}{left:<9}")
    table = "\n".join(lines)

    n_past = sum(1 for s in stops if s.status == "PAST_CLOSE")
    n_risk = len(stops) - n_past
    summary = (f"*⚠️ FedEx Pickup Manifest — At Risk / Past Close*\n"
               f"{len(stops)} stop(s) within 60 min of close ({n_risk} at risk, "
               f"{n_past} past close) — {now:%I:%M %p %Z}")
    text = f"At Risk / Past Close: {len(stops)} stop(s) — see table"
    blocks = [
        {"type": "section", "text": {"type": "mrkdwn", "text": summary}},
        {"type": "section", "text": {"type": "mrkdwn", "text": f"```{table}```"}},
    ]
    return text, blocks


def build_daily_rows(df: pd.DataFrame, now: datetime, warn_minutes: float = 60.0,
                     manifest_day: Optional[datetime] = None) -> list[dict]:
    """Every row on today's manifest (not just at-risk ones), trimmed to the
    columns the 'Daily Manifest' Slack channel wants, with a per-row status:
      'done'    -- already closed out (PU Closed filled), or a reason code is
                   on file (cancelled/nothing to track) -- nothing more to do
      'pending' -- PU Closed empty, no reason code, but more than warn_minutes
                   from close -- not urgent yet
      'warn'    -- PU Closed empty, no reason code, <=warn_minutes to close
      'overdue' -- PU Closed empty, no reason code, close time already passed
    Sorted by status (warn, overdue, pending, done) so the rows needing
    attention stay pinned near the top and completed ones sink to the bottom
    -- see _STATUS_SORT_RANK below."""
    day = (manifest_day or now).date()
    rows: list[dict] = []

    for _, row in df.iterrows():
        close_t = _parse_hhmm(row.get("Close"))
        ready_t = _parse_hhmm(row.get("Ready"))
        pu_closed_t = _parse_hhmm(row.get("PU Closed"))
        reas_code = row.get("Reas Code")
        pu_closed_blank = _is_blank(row.get("PU Closed"))
        reas_blank = _is_blank(reas_code)

        if not pu_closed_blank:
            status = "done"
        elif not reas_blank:
            status = "done"  # cancelled / reason on file -- nothing left to track
        elif close_t is None:
            status = "pending"  # no close time to judge urgency by
        else:
            close_dt = datetime.combine(day, close_t, tzinfo=now.tzinfo)
            minutes_left = (close_dt - now).total_seconds() / 60.0
            if minutes_left < 0:
                status = "overdue"
            elif minutes_left <= warn_minutes:
                status = "warn"
            else:
                status = "pending"

        station = row.get("Station")
        rows.append({
            "station": int(station) if not pd.isna(station) else None,
            "wa": str(row.get("WA") or "").strip(),
            "shipper": str(row.get("Shipper Name") or "").strip(),
            "address": ", ".join(x for x in (
                str(row.get("Address Line 1") or "").strip(),
                str(row.get("City") or "").strip(),
                str(row.get("State") or "").strip(),
            ) if x),
            "ready": ready_t.strftime("%H:%M") if ready_t else "",
            "close": close_t.strftime("%H:%M") if close_t else "",
            "pu_closed": pu_closed_t.strftime("%H:%M") if pu_closed_t else "",
            "reas_code": "" if reas_blank else str(reas_code).strip(),
            "status": status,
            "_sort": (_STATUS_SORT_RANK[status], close_t or dtime(23, 59)),
        })

    rows.sort(key=lambda r: r["_sort"])
    for r in rows:
        r.pop("_sort", None)
    return rows


def format_alert_blocks(stop: RiskyStop) -> list[dict]:
    """Slack Block Kit version: a 2-column Metric/Value table (Alert Level, Account,
    Station ID, Address, Cutoff Time, Time Remaining, System Flag), matching the
    layout requested 2026-10-19. 'Account' is the Shipper Name column off the
    manifest -- FedEx's own term for the customer on that stop (e.g. "Dollar
    General"), not a billing/FedEx-account-number field."""
    alert_level = "⛔ Past Close" if stop.status == "PAST_CLOSE" else "🟠 At Risk"
    cutoff = stop.close.strftime("%I:%M %p").lstrip("0")
    if stop.status == "PAST_CLOSE":
        time_remaining = f"Past close by {abs(stop.minutes_to_close):.0f} min"
    else:
        time_remaining = f"{stop.minutes_to_close:.0f} minutes"

    def field(label: str, value: str) -> dict:
        return {"type": "mrkdwn", "text": f"*{label}*\n{value}"}

    return [{
        "type": "section",
        "fields": [
            field("Alert Level", alert_level),
            field("Account", stop.shipper_name or "Unknown"),
            field("Station ID", f"{stop.station} (WA {stop.wa})"),
            field("Address", stop.address or "N/A"),
            field("Cutoff Time", cutoff),
            field("Time Remaining", time_remaining),
            field("System Flag", "No pickup closed / No reason on file"),
        ],
    }]
