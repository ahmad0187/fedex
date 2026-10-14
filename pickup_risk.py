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
