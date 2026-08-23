import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

"""
breach_monitor.py — Weekly breach audit for ThreeUp CAO Agent.

Reads last 6 weeks of raw performance data per non-churned client,
recomputes Hook / Offer / Cancellation / Loyalty flags from scratch
(never trusts agent-written flag columns), identifies consecutive-week
streaks, writes results to the 'CAO Breach Monitor' tab in the AM
Google Sheet, and sends a weekly email summary.

Also detects data-coverage gaps: clients with missing weeks or lagging
behind the most recently entered week across all clients.

Entry points:
  python agent/main.py --breach-audit            (production)
  python agent/main.py --breach-audit --dry-run  (preview only)
"""

import json
import logging
import os
import re
import time
import traceback
from datetime import datetime, date, timedelta
from pathlib import Path

import pytz
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

logger = logging.getLogger("breach_monitor")

_HERE = Path(__file__).parent.resolve()
_ROOT = _HERE.parent

CREDENTIALS_PATH = str(_ROOT / "config" / "credentials.json")
SHEET_IDS_PATH   = _ROOT / "config" / "sheet_ids.json"
AM_SHEET_ID      = "12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA"
GHL_BASE         = "https://services.leadconnectorhq.com"
COUNTER_FID      = "nNx5vev4O2dBgLbqYNSh"
AMMAN_TZ         = pytz.timezone("Asia/Amman")
TAB_NAME         = "CAO Breach Monitor"
MASTER_TAB       = "Clients Master"

INACTIVITY_LOG_PATH  = _ROOT / "logs" / "inactivity_alerts.json"
THROTTLE_PATH        = _ROOT / "logs" / "last_whatsapp_send.json"
DRIP_DELAY           = 1800
INACTIVITY_MIN_WEEKS = 3
COL_STRATEGY_CHANGED = 12   # col L, 1-based (Strategy Changed On)
COL_STRATEGY_RESET   = 13   # col M, 1-based (Strategy Reset Done)

# ── Thresholds — must match config/thresholds.json and sheets_manager.py ──────
HOOK_MIN     = 0.60   # TL / NP
OFFER_MIN    = 0.60   # Reviews / TL
CANCEL_MAX   = 0.25
LOYALTY_DROP = 0.30

# ── Low-volume thresholds ─────────────────────────────────────────────────────
LOW_VOL_MIN  = 7   # New Patients This Week below this → flag fires
LOW_VOL_WARN = 2   # 2 consecutive weeks → WARNING
LOW_VOL_CRIT = 3   # 3+ consecutive weeks → CRITICAL
LOW_VOL_SKIP = 3   # skip entirely for clients with fewer data rows than this

# ── Column indices (0-based, Sheet1 data rows) ────────────────────────────────
C_DATE     =  0   # A  Date Range
C_REVIEWS  =  1   # B  Google Reviews Count
C_NEWPT    =  3   # D  New Patients This Week
C_TRIGGER  =  7   # H  Trigger Links Clicked
C_CANCEL   = 15   # P  Cancellation Rate %
C_LOY_THIS = 18   # S  Loyalty Points This Period
C_LOY_PRI  = 20   # U  Loyalty Points Prior Period

FLAG_KEYS   = ["hook", "offer", "cancel_flag", "loyalty"]
FLAG_LABELS = {
    "hook":        "Hook (TL/NP)",
    "offer":       "Offer (Rev/TL)",
    "cancel_flag": "Cancellation",
    "loyalty":     "Loyalty Drop",
    "low_volume":  "Low Volume (NP<7)",
}

# FLAG_KEYS drives GHL counter + existing breach alerts (unchanged).
# LOW_VOLUME is tracked separately so it never inflates the GHL strike counter.
ALL_FLAG_KEYS = FLAG_KEYS + ["low_volume"]

WEEKS_TO_READ = 8   # read last 8 raw rows; analyse last 6 with actual data
STREAK_WARN   = 3   # 3+ consecutive → warning
STREAK_CRIT   = 4   # 4+ consecutive → critical

# Contacts known to return HTTP 400 from GHL (invalid / deleted contact)
_KNOWN_400_CONTACT_IDS: set = set()


# ── Google Sheets helpers ─────────────────────────────────────────────────────

def _get_sheets_service(readonly: bool = True):
    scopes = (
        ["https://www.googleapis.com/auth/spreadsheets.readonly",
         "https://www.googleapis.com/auth/drive.readonly"]
        if readonly else
        ["https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive"]
    )
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_PATH, scopes=scopes
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _read_client_rows(svc, sheet_id: str, n: int = WEEKS_TO_READ):
    """Return last n non-empty data rows, each padded to 32 cols. Returns list or None."""
    try:
        res = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range="Sheet1!A2:AF",
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute()
        rows = res.get("values", [])
        rows = [r for r in rows if len(r) > 0 and str(r[0]).strip()]
        rows = [r + [''] * (32 - len(r)) for r in rows]
        return rows[-n:]
    except Exception as e:
        logger.error(f"[breach] sheet read error {sheet_id}: {e}")
        return None


# ── Flag computation ──────────────────────────────────────────────────────────

def _num(val) -> float:
    try:
        return float(val) if val not in (None, '', 'N/A') else 0.0
    except (ValueError, TypeError):
        return 0.0


def _compute_flags(row: list) -> dict:
    new_pt   = _num(row[C_NEWPT])
    trigger  = _num(row[C_TRIGGER])
    reviews  = _num(row[C_REVIEWS])
    cancel   = _num(row[C_CANCEL])
    loy_this = _num(row[C_LOY_THIS])
    loy_pri  = _num(row[C_LOY_PRI])

    hook_pct  = round(trigger / new_pt,  4) if new_pt   > 0 else None
    offer_pct = round(reviews / trigger, 4) if trigger  > 0 else None
    loy_chg   = round((loy_this - loy_pri) / loy_pri, 4) if loy_pri > 0 else None

    return {
        "date":        str(row[C_DATE]),
        "new_pt":      new_pt,
        "trigger":     trigger,
        "reviews":     reviews,
        "cancel":      cancel,
        "hook_pct":    hook_pct,
        "offer_pct":   offer_pct,
        "loy_chg":     loy_chg,
        "hook":        (hook_pct  < HOOK_MIN)     if hook_pct  is not None else False,
        "offer":       (offer_pct < OFFER_MIN)    if offer_pct is not None else False,
        "cancel_flag": cancel > CANCEL_MAX,
        "loyalty":     (loy_chg <= -LOYALTY_DROP) if loy_chg  is not None else False,
        "low_volume":  new_pt < LOW_VOL_MIN,
        "has_data":    any(v > 0 for v in [new_pt, trigger, reviews]),
    }


def _streak_current(rows: list, key: str) -> int:
    """Consecutive True from the END of the list (active streak)."""
    count = 0
    for r in reversed(rows):
        if r[key]:
            count += 1
        else:
            break
    return count


def _streak_max(rows: list, key: str) -> int:
    """Longest contiguous True run anywhere in the list."""
    best = cur = 0
    for r in rows:
        cur = (cur + 1) if r[key] else 0
        best = max(best, cur)
    return best


def _alert_level(current_streak: int) -> str:
    if current_streak >= STREAK_CRIT:
        return "critical"
    if current_streak >= STREAK_WARN:
        return "warning"
    if current_streak >= 1:
        return "watch"
    return "ok"


def _alert_level_low_vol(current_streak: int) -> str:
    """Low-volume uses a lower trigger (2/3) than the standard breach flags (3/4)."""
    if current_streak >= LOW_VOL_CRIT:
        return "critical"
    if current_streak >= LOW_VOL_WARN:
        return "warning"
    if current_streak >= 1:
        return "watch"
    return "ok"


def _overall_current_streak(rows: list) -> int:
    """Consecutive weeks from END where ANY breach flag was True."""
    count = 0
    for r in reversed(rows):
        if any(r[fk] for fk in FLAG_KEYS):
            count += 1
        else:
            break
    return count


def _ghl_auto_label(streak: int) -> str:
    if streak == 0: return "— (recovery)"
    if streak == 1: return "1 → first bad week alert"
    if streak == 2: return "2 → noted"
    if streak == 3: return "3 → note created"
    return f"{streak} → change-strategy alert"


def _fmt_flag_val(row: dict, flag_key: str) -> str:
    if flag_key == "hook":
        v = row.get("hook_pct")
        return f"{v*100:.0f}%" if v is not None else "n/a"
    if flag_key == "offer":
        v = row.get("offer_pct")
        return f"{v*100:.0f}%" if v is not None else "n/a"
    if flag_key == "cancel_flag":
        return f"{row.get('cancel', 0)*100:.0f}%"
    if flag_key == "loyalty":
        v = row.get("loy_chg")
        return f"{v*100:+.0f}%" if v is not None else "n/a"
    if flag_key == "low_volume":
        return str(int(_num(row.get("new_pt", 0))))
    return "n/a"


# ── GHL counter (uses ThreeUp main account key, not per-client key) ───────────

def _ghl_get_counter(threeup_api_key: str, contact_id: str):
    """
    Read the strike-counter field from GHL using the ThreeUp Solutions
    main-account PIT key (pit-5fe4c7a5...).  All contacts live in location
    ck8gYVHUEnLEIg0OVsLd — per-client sub-account keys CANNOT access them.
    """
    try:
        r = requests.get(
            f"{GHL_BASE}/contacts/{contact_id}",
            headers={
                "Authorization": f"Bearer {threeup_api_key}",
                "Version":       "2021-07-28",
                "User-Agent":    "ThreeUp-CAO-Agent/1.0",
            },
            timeout=10,
        )
        if r.status_code == 200:
            fields = r.json().get("contact", {}).get("customFields", [])
            for f in fields:
                if f.get("id") == COUNTER_FID:
                    v = f.get("value", 0)
                    return int(_num(v)) if v not in (None, "") else 0
            return 0
        logger.warning(f"[breach] GHL {r.status_code} for {contact_id}")
        return None
    except Exception as e:
        logger.warning(f"[breach] GHL counter exception for {contact_id}: {e}")
        return None


def _ghl_write_counter(threeup_api_key: str, contact_id: str, value: int) -> bool:
    """
    Write the current overall streak to the GHL strike-counter field.
    Uses ThreeUp main-account PIT key.  Returns True on 200.
    """
    try:
        r = requests.put(
            f"{GHL_BASE}/contacts/{contact_id}",
            headers={
                "Authorization": f"Bearer {threeup_api_key}",
                "Version":       "2021-07-28",
                "User-Agent":    "ThreeUp-CAO-Agent/1.0",
                "Content-Type":  "application/json",
            },
            json={"customFields": [{"id": COUNTER_FID, "field_value": value}]},
            timeout=10,
        )
        if r.status_code == 200:
            return True
        logger.warning(f"[breach] GHL write {r.status_code} for {contact_id}: {r.text[:120]}")
        return False
    except Exception as e:
        logger.warning(f"[breach] GHL write exception for {contact_id}: {e}")
        return False


_COUNTER_TIER = {
    0: "recovery",
    1: "first bad week",
    2: "noted",
    3: "note created",
    4: "change-strategy review required",
}


def _breach_note_body(written: int, true_streak: int) -> str:
    """Human-readable note text that travels with every counter write."""
    tier       = _COUNTER_TIER.get(written, f"counter {written}")
    streak_str = f"{true_streak}wk" if true_streak > 0 else "clean (0wk)"
    return f"CAO Breach Monitor — counter: {written} | {tier} | streak: {streak_str}"


def _ghl_add_note(threeup_api_key: str, contact_id: str, body: str) -> bool:
    """POST a note to a GHL contact (triggers 'note added' workflow). Returns True on 200/201."""
    try:
        r = requests.post(
            f"{GHL_BASE}/contacts/{contact_id}/notes",
            headers={
                "Authorization": f"Bearer {threeup_api_key}",
                "Version":       "2021-07-28",
                "User-Agent":    "ThreeUp-CAO-Agent/1.0",
                "Content-Type":  "application/json",
            },
            json={"body": body},
            timeout=10,
        )
        if r.status_code in (200, 201):
            return True
        logger.warning(f"[breach] note POST {r.status_code} for {contact_id}: {r.text[:120]}")
        return False
    except Exception as e:
        logger.warning(f"[breach] note exception for {contact_id}: {e}")
        return False


# ── Strategy-change helpers ───────────────────────────────────────────────────

def _parse_date_str(s) -> "date | None":
    """Parse DD/MM/YYYY or YYYY-MM-DD → date, or None."""
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except ValueError:
            continue
    return None


def _read_strategy_meta(svc) -> dict:
    """Read Clients Master A:M → {am_name: {changed_on, reset_done, sheet_row}}."""
    try:
        res = svc.spreadsheets().values().get(
            spreadsheetId=AM_SHEET_ID,
            range=f"'{MASTER_TAB}'!A:M",
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute()
        rows = res.get("values", [])
        meta = {}
        for i, row in enumerate(rows):
            if i < 2:
                continue
            if not row or not str(row[0]).strip():
                continue
            name = str(row[0]).strip()
            def _c(j, _r=row):
                return str(_r[j]).strip() if len(_r) > j and str(_r[j]).strip() else ""
            meta[name] = {
                "changed_on": _parse_date_str(_c(11)),  # col L (0-based index 11)
                "reset_done": _parse_date_str(_c(12)),  # col M (0-based index 12)
                "sheet_row":  i + 1,                    # 1-based spreadsheet row
            }
        return meta
    except Exception as e:
        logger.warning(f"[breach] _read_strategy_meta failed: {e}")
        return {}


def _filter_rows_after(rows: list, cutoff_date: date, today: date) -> list:
    """Return computed rows whose week-start is >= cutoff_date. Unparseable dates pass through."""
    result = []
    for r in rows:
        d = _parse_week_start(r["date"], today)
        if d is None or d >= cutoff_date:
            result.append(r)
    return result


def _write_strategy_reset_done(reset_writes: list) -> None:
    """Stamp col M (Strategy Reset Done) = change date string, for idempotency."""
    if not reset_writes:
        return
    try:
        svc = _get_sheets_service(readonly=False)
        data = [
            {"range": f"'{MASTER_TAB}'!M{row}", "values": [[date_str]]}
            for row, date_str in reset_writes
        ]
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=AM_SHEET_ID,
            body={"valueInputOption": "RAW", "data": data},
        ).execute()
        print(f"[breach] Strategy reset stamped in col M: {len(reset_writes)} client(s)")
    except Exception as e:
        print(f"[breach] !! _write_strategy_reset_done failed: {e}")


# ── Coverage gap detection ────────────────────────────────────────────────────

def _fmt_date(d: date) -> str:
    months = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    return f"{d.day} {months[d.month - 1]}"


def _parse_week_start(date_str: str, today: date):
    """
    Parse 'Sat 30/5 – Fri 5/6' → date(2026, 5, 30).
    Returns None if unparseable.
    """
    m = re.match(r'\w+\s+(\d{1,2})/(\d{1,2})', str(date_str).strip())
    if not m:
        return None
    day, month = int(m.group(1)), int(m.group(2))
    year = today.year
    try:
        d = date(year, month, day)
        # If the parsed date is more than 3 months in the future, use previous year
        if (d - today).days > 90:
            d = date(year - 1, month, day)
        return d
    except ValueError:
        return None


def _check_data_coverage(raw_data: dict, today: date) -> dict:
    """
    Phase 1 — Per-client data coverage check.

    Strategy: find the "frontier" = the most recent entry date across ALL clients.
    Any client whose last entry is 7+ days behind the frontier is flagged.
    This is robust to mid-week runs where everyone is equally behind
    (no one gets falsely flagged since they're all at the same date).

    Also detects interior gaps: missing weeks within a client's sequence
    (gaps > 9 days between consecutive entries).

    Returns:
    {
        frontier:        "30 May" | None,
        frontier_date:   date | None,
        behind_frontier: [(name, weeks_behind, last_entry_str), ...],
        gap_clients:     [(name, [missing_week_labels]), ...],
        per_client:      {name: {last_entry, last_entry_date, interior_gaps}},
    }
    """
    per_client = {}

    for name, data in raw_data.items():
        computed  = data.get("computed", [])
        data_rows = [r for r in computed if r["has_data"]]

        week_dates = set()
        for r in data_rows:
            d = _parse_week_start(r["date"], today)
            if d is not None:
                week_dates.add(d)

        last_entry_date = max(week_dates) if week_dates else None

        # Interior gaps: weeks missing between earliest and latest entry
        interior_gaps = []
        if len(week_dates) >= 2:
            sorted_dates = sorted(week_dates)
            for i in range(len(sorted_dates) - 1):
                diff = (sorted_dates[i + 1] - sorted_dates[i]).days
                if diff > 9:  # >9 days → at least one week missing
                    n_missed = (diff // 7) - 1
                    for j in range(n_missed):
                        gap_sat = sorted_dates[i] + timedelta(weeks=j + 1)
                        interior_gaps.append(_fmt_date(gap_sat))

        per_client[name] = {
            "last_entry_date": last_entry_date,
            "last_entry":      _fmt_date(last_entry_date) if last_entry_date else None,
            "interior_gaps":   interior_gaps,
            "data_count":      len(data_rows),
        }

    # Frontier = max last-entry date across all clients
    all_dates = [cd["last_entry_date"] for cd in per_client.values()
                 if cd["last_entry_date"] is not None]
    frontier_date = max(all_dates) if all_dates else None

    behind_frontier = []
    if frontier_date:
        for name, cd in per_client.items():
            ld = cd["last_entry_date"]
            if ld is None:
                behind_frontier.append((name, 99, "no data"))
            elif (frontier_date - ld).days >= 7:
                weeks_behind = (frontier_date - ld).days // 7
                behind_frontier.append((name, weeks_behind, cd["last_entry"]))
    behind_frontier.sort(key=lambda x: -x[1])

    gap_clients = [
        (name, cd["interior_gaps"])
        for name, cd in per_client.items()
        if cd["interior_gaps"]
    ]

    return {
        "frontier":        _fmt_date(frontier_date) if frontier_date else None,
        "frontier_date":   frontier_date,
        "behind_frontier": behind_frontier,
        "gap_clients":     gap_clients,
        "per_client":      per_client,
    }


# ── Core audit ────────────────────────────────────────────────────────────────

def _build_report(clients: list, sheet_ids: dict, threeup_api_key: str) -> dict:
    """
    Returns a report dict:
    {
        run_date:      str,
        clients:       { name: { flags, ghl_counter, data_weeks, coverage_gap,
                                 last_entry, interior_gaps } },
        alerts_3plus:  [(name, fk, label, streak, week_vals, week_dates), ...],
        alerts_4plus:  same but streak >= 4,
        coverage:      { frontier, behind_frontier, gap_clients, per_client },
        errors:        { name: reason },
    }
    """
    run_date = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d")
    today    = datetime.now(AMMAN_TZ).date()

    svc = _get_sheets_service(readonly=True)
    strategy_meta = _read_strategy_meta(svc)

    raw_data = {}
    errors   = {}

    print(f"[breach] Reading {len(clients)} client sheets...")
    for c in clients:
        name = c.get("name", "")
        sid  = sheet_ids.get(name)
        if not sid:
            errors[name] = "no sheet_id in sheet_ids.json"
            continue
        rows = _read_client_rows(svc, sid)
        if rows is None:
            errors[name] = "Sheets API error"
            continue
        raw_data[name] = {"raw": rows, "client": c}
        time.sleep(0.12)

    print(f"[breach] Sheets done: {len(raw_data)} ok, {len(errors)} errors")

    # GHL counter reads — must use threeup_api_key, NOT per-client ghl_api_key
    print(f"[breach] Reading GHL strike counters (threeup_api_key)...")
    for name, data in raw_data.items():
        contact_id = data["client"].get("contact_id", "")
        data["ghl_counter"] = (
            _ghl_get_counter(threeup_api_key, contact_id)
            if threeup_api_key and contact_id
            else None
        )
        time.sleep(0.08)

    clients_out  = {}
    alerts_3plus = []
    alerts_4plus = []
    alerts_lowvol = []

    for name, data in raw_data.items():
        computed  = [_compute_flags(r) for r in data["raw"]]
        data["computed"] = computed          # store for coverage check
        data_rows = [r for r in computed if r["has_data"]]
        last6     = data_rows[-6:]

        # Strategy-change: filter to post-change rows for current streak only.
        # MaxStreak and W1-W6 always use full history (last6).
        sc_info      = strategy_meta.get(name, {})
        sc_date      = sc_info.get("changed_on")
        sc_reset     = sc_info.get("reset_done")
        reset_needed = sc_date is not None and sc_reset != sc_date
        post_change  = _filter_rows_after(data_rows, sc_date, today) if sc_date else data_rows
        last6_curr   = post_change[-6:]

        # Data-quality gap note (structural: empty sheet or all-zero)
        coverage_gap = None
        if len(data_rows) == 0:
            coverage_gap = "no data rows with values"
        elif len(data_rows) < 3:
            coverage_gap = f"only {len(data_rows)} week(s) — streak analysis unreliable"
        if data_rows and all(
            r["new_pt"] == 0 and r["trigger"] == 0 for r in data_rows
        ):
            suffix = "all NP/TL = 0 — Hook/Offer uncomputable"
            coverage_gap = f"{coverage_gap} | {suffix}" if coverage_gap else suffix

        flags_out = {}
        for fk in FLAG_KEYS:
            cs = _streak_current(last6_curr, fk)   # post-change rows only
            ms = _streak_max(last6, fk)             # full history
            week_values = [_fmt_flag_val(r, fk) for r in last6]
            week_dates  = [r["date"][:14] for r in last6]
            flags_out[fk] = {
                "label":       FLAG_LABELS[fk],
                "current":     cs,
                "max":         ms,
                "alert":       _alert_level(cs),
                "week_values": week_values,
                "week_dates":  week_dates,
            }
            if cs >= STREAK_WARN:
                entry = (name, fk, FLAG_LABELS[fk], cs, week_values[-cs:], week_dates[-cs:])
                alerts_3plus.append(entry)
                if cs >= STREAK_CRIT:
                    alerts_4plus.append(entry)

        # ── Low-volume flag (separate; does not affect GHL counter) ──────────
        # Skip entirely for new clients with fewer than LOW_VOL_SKIP data rows.
        if len(post_change) >= LOW_VOL_SKIP:
            lv_cs = _streak_current(last6_curr, "low_volume")
            lv_ms = _streak_max(last6, "low_volume")
        else:
            lv_cs = lv_ms = 0
        lv_wv = [_fmt_flag_val(r, "low_volume") for r in last6]
        lv_wd = [r["date"][:14] for r in last6]
        flags_out["low_volume"] = {
            "label":       FLAG_LABELS["low_volume"],
            "current":     lv_cs,
            "max":         lv_ms,
            "alert":       _alert_level_low_vol(lv_cs),
            "week_values": lv_wv,
            "week_dates":  lv_wd,
        }
        if lv_cs >= LOW_VOL_WARN:
            alerts_lowvol.append(
                (name, "low_volume", FLAG_LABELS["low_volume"],
                 lv_cs, lv_wv[-lv_cs:], lv_wd[-lv_cs:])
            )

        last2_raw   = data_rows[-2:] if len(data_rows) >= 2 else []
        last2_np_tl = [(r["new_pt"], r["trigger"]) for r in last2_raw]

        overall_streak = _overall_current_streak(last6_curr)
        clients_out[name] = {
            "flags":               flags_out,
            "ghl_counter":         data.get("ghl_counter"),
            "data_weeks":          len(data_rows),
            "last2_np_tl":         last2_np_tl,
            "coverage_gap":        coverage_gap,
            "current_streak":      overall_streak,
            "strategy_changed_on": sc_date.isoformat() if sc_date else None,
            "reset_needed":        reset_needed,
            "sc_sheet_row":        sc_info.get("sheet_row"),
            # coverage fields filled in below after _check_data_coverage()
            "last_entry":          None,
            "interior_gaps":       [],
        }

    # Phase 1 — coverage check (uses raw_data["computed"] populated above)
    coverage = _check_data_coverage(raw_data, today)
    for name in clients_out:
        cv = coverage["per_client"].get(name, {})
        clients_out[name]["last_entry"]    = cv.get("last_entry")
        clients_out[name]["interior_gaps"] = cv.get("interior_gaps", [])

    alerts_3plus.sort(key=lambda x: -x[3])
    alerts_4plus.sort(key=lambda x: -x[3])
    alerts_lowvol.sort(key=lambda x: -x[3])

    return {
        "run_date":      run_date,
        "clients":       clients_out,
        "alerts_3plus":  alerts_3plus,
        "alerts_4plus":  alerts_4plus,
        "alerts_lowvol": alerts_lowvol,
        "coverage":      coverage,
        "errors":        errors,
        "strategy_meta": strategy_meta,
    }


# ── Sheet write ───────────────────────────────────────────────────────────────

def write_breach_to_sheet(report: dict) -> None:
    """
    Write breach results to the 'CAO Breach Monitor' tab in the AM sheet.
    Schema: 15 columns A:O
      A Client | B Flag | C Streak | D MaxStreak | E AlertLevel
      F-K W1–W6 | L CoverageGap | M Updated | N LastEntry | O MissingWeeks
    """
    try:
        svc     = _get_sheets_service(readonly=False)
        updated = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M")

        meta   = svc.spreadsheets().get(spreadsheetId=AM_SHEET_ID).execute()
        titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if TAB_NAME not in titles:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=AM_SHEET_ID,
                body={"requests": [{"addSheet": {"properties": {"title": TAB_NAME}}}]},
            ).execute()
            print(f"[breach] Created '{TAB_NAME}' tab")

        # 15 columns (A:O)
        header = [
            "Client", "Flag", "Streak", "MaxStreak", "AlertLevel",
            "W1", "W2", "W3", "W4", "W5", "W6",
            "CoverageGap", "Updated", "LastEntry", "MissingWeeks",
        ]
        rows = [header]

        coverage = report.get("coverage", {})

        for name, cdata in sorted(report["clients"].items()):
            cv      = coverage.get("per_client", {}).get(name, {})
            last_e  = cv.get("last_entry") or ""
            gaps    = ", ".join(cv.get("interior_gaps", []))

            for fk in ALL_FLAG_KEYS:
                f   = cdata["flags"][fk]
                wv  = f["week_values"]
                wv6 = ([""] * (6 - len(wv))) + wv  # left-pad to 6 cols
                rows.append([
                    name,
                    f["label"],
                    f["current"],
                    f["max"],
                    f["alert"],
                    *wv6,
                    cdata.get("coverage_gap") or "",
                    updated,
                    last_e,
                    gaps,
                ])

        for name, reason in sorted(report["errors"].items()):
            rows.append([name, "ERROR", 0, 0, "error",
                         "", "", "", "", "", "", reason, updated, "", ""])

        svc.spreadsheets().values().clear(
            spreadsheetId=AM_SHEET_ID,
            range=f"'{TAB_NAME}'!A:O",
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=AM_SHEET_ID,
            range=f"'{TAB_NAME}'!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        print(f"[breach] '{TAB_NAME}' written: {len(rows)-1} data rows")
    except Exception as e:
        print(f"[breach] write_breach_to_sheet failed: {e}")
        traceback.print_exc()


# ── Email composition ─────────────────────────────────────────────────────────

def _compose_email(report: dict) -> tuple:
    """Returns (subject, body) for the weekly breach email."""
    run_date     = report["run_date"]
    alerts_3plus = report["alerts_3plus"]
    alerts_4plus = report["alerts_4plus"]
    errors       = report["errors"]
    coverage     = report.get("coverage", {})

    n_clients = len({e[0] for e in alerts_3plus})

    subject = (
        f"ThreeUp Weekly Breach Report — {n_clients} client(s) struggling 3+ weeks"
        if n_clients
        else "ThreeUp Weekly Breach Report — No clients breaching 3+ weeks"
    )

    lines = [
        "ThreeUp Weekly Breach Report",
        f"Run date: {run_date}",
        "",
    ]

    # ── BREACH ALERTS (headline) ──────────────────────────────────────────────
    if alerts_4plus:
        lines += ["━━━ CRITICAL — 4+ Consecutive Weeks ━━━", ""]
        for name, fk, fl, streak, week_vals, week_dates in alerts_4plus:
            sc = report["clients"][name].get("strategy_changed_on")
            lines.append(f"  {name} — {fl}: {streak} consecutive weeks")
            if sc:
                lines.append(f"  ⚙ strategy changed {sc} — streak restarted from this date")
            pairs = [f"{d[:10]}: {v}" for d, v in zip(week_dates, week_vals)]
            lines.append("    " + " | ".join(pairs))
            lines.append("")

    warning_only = [e for e in alerts_3plus if e[3] < STREAK_CRIT]
    if warning_only:
        lines += ["━━━ WARNING — 3 Consecutive Weeks ━━━", ""]
        for name, fk, fl, streak, week_vals, week_dates in warning_only:
            sc = report["clients"][name].get("strategy_changed_on")
            lines.append(f"  {name} — {fl}: {streak} consecutive weeks")
            if sc:
                lines.append(f"  ⚙ strategy changed {sc} — streak restarted from this date")
            pairs = [f"{d[:10]}: {v}" for d, v in zip(week_dates, week_vals)]
            lines.append("    " + " | ".join(pairs))
            lines.append("")

    if not alerts_3plus:
        lines += [
            "━━━ All Clear ━━━",
            "",
            "  No clients with 3+ consecutive weeks on any flag.",
            "",
        ]

    # ── LOW VOLUME ────────────────────────────────────────────────────────────
    alerts_lowvol = report.get("alerts_lowvol", [])
    if alerts_lowvol:
        lines += [f"━━━ LOW VOLUME — {LOW_VOL_WARN}+ Consecutive Weeks (NP < {LOW_VOL_MIN}) ━━━", ""]
        # Sorted by streak descending (already sorted in _build_report, but make explicit)
        for name, fk, fl, streak, week_vals, week_dates in sorted(alerts_lowvol, key=lambda x: -x[3]):
            sev = "★★ CRITICAL" if streak >= LOW_VOL_CRIT else "★  WARNING "
            # NP pattern: all 6 weeks from flags_out, not just the streak window
            full_wv = report["clients"][name]["flags"]["low_volume"]["week_values"]
            np_pattern = "  ".join(
                f"[{v}]" if int(v) < LOW_VOL_MIN else f" {v} "
                for v in full_wv
            )
            lines.append(f"  {sev}  {name} — {streak} consecutive weeks")
            lines.append(f"    NP last 6 wks (oldest→newest): {np_pattern}")
            lines.append("")
    else:
        lines += [
            f"━━━ LOW VOLUME (NP < {LOW_VOL_MIN}) ━━━",
            "",
            f"  No clients with {LOW_VOL_WARN}+ consecutive low-volume weeks.",
            "",
        ]

    # ── DATA COVERAGE ─────────────────────────────────────────────────────────
    lines += ["━━━ DATA COVERAGE ━━━", ""]

    frontier = coverage.get("frontier")
    behind   = coverage.get("behind_frontier", [])
    gaps     = coverage.get("gap_clients", [])

    if frontier:
        lines.append(f"  Most recent data entry across all clients: {frontier}")
        lines.append("")

    if behind:
        lines.append(f"  ⚠ Behind the most recently entered week ({frontier}):")
        for name, weeks_behind, last_e in behind:
            wk_word = "week" if weeks_behind == 1 else "weeks"
            lines.append(f"    {name} — {weeks_behind} {wk_word} behind (last entry: {last_e})")
        lines.append("")

    if gaps:
        lines.append("  ⚠ Interior gaps in sequence (missing weeks mid-history):")
        for name, missing in gaps:
            lines.append(f"    {name}: missing {', '.join(missing)}")
        lines.append("")

    if not behind and not gaps:
        lines.append("  ✓ All clients have consistent data — no coverage gaps detected.")
        lines.append("")

    # ── ERRORS ───────────────────────────────────────────────────────────────
    if errors:
        lines += ["━━━ Errors (sheet read failures) ━━━", ""]
        for name, reason in sorted(errors.items()):
            lines.append(f"  {name}: {reason}")
        lines.append("")

    # ── GHL SYNC FAILURES ────────────────────────────────────────────────────
    ghl_errors = report.get("ghl_errors", {})
    if ghl_errors:
        lines += ["━━━ !! GHL SYNC FAILURES — ACTION REQUIRED ━━━", ""]
        lines.append("  The following clients had GHL counter/note failures.")
        lines.append("  Their workflows did NOT fire. Manual follow-up required.")
        lines.append("")
        for name, reason in sorted(ghl_errors.items()):
            lines.append(f"  !! {name}: {reason}")
        lines.append("")

    lines += [
        "─" * 60,
        "This report is sent every Saturday regardless of results.",
        "If you stop receiving it, the scheduled task has stopped.",
    ]

    return subject, "\n".join(lines)


# ── Dry-run printer ───────────────────────────────────────────────────────────

def _print_report(report: dict) -> None:
    SEP = "═" * 110
    coverage = report.get("coverage", {})

    print("\n" + SEP)
    print("  PER-CLIENT BREACH TABLE  (last 6 data weeks, "
          "H=Hook TL/NP<60%, O=Offer Rev/TL<60%, C=Cancel>25%, L=Loyalty≥30%, V=Low Vol NP<7)")
    print(SEP)

    for name, cdata in sorted(report["clients"].items()):
        flags = cdata["flags"]
        ghl_c = cdata.get("ghl_counter")
        dw    = cdata["data_weeks"]
        gap   = cdata.get("coverage_gap")
        le    = cdata.get("last_entry") or "?"
        ig    = cdata.get("interior_gaps", [])

        print(f"\n{'◼ ' + name}")
        ghl_s = str(ghl_c) if ghl_c is not None else "N/A"
        ig_s  = f"  gaps: {', '.join(ig)}" if ig else ""
        print(f"  GHL counter: {ghl_s}  |  Data wks: {dw}  |  Last entry: {le}{ig_s}")
        if gap:
            print(f"  ⚠ {gap}")
        sc = cdata.get("strategy_changed_on")
        if sc:
            print(f"  ⚙ strategy changed {sc} — current streak restarted from this date")

        for fk in ALL_FLAG_KEYS:
            f = flags[fk]
            if f["current"] == 0 and f["max"] == 0:
                continue
            wv_str = "  ".join(
                f"{d[:10]}:{v}" for d, v in zip(f["week_dates"], f["week_values"])
            )
            print(f"  {FLAG_LABELS[fk]:<22} "
                  f"current={f['current']}wk  max={f['max']}wk  [{wv_str}]")

    print("\n\n" + SEP)
    print("  ★  BREACH ALERTS — 3+ CONSECUTIVE WEEKS")
    print(SEP)
    if report["alerts_3plus"]:
        for name, fk, fl, streak, week_vals, week_dates in report["alerts_3plus"]:
            pairs = "  ".join(f"{d[:10]}:{v}" for d, v in zip(week_dates, week_vals))
            alert = "★★ CRITICAL" if streak >= STREAK_CRIT else "★  WARNING "
            print(f"  {alert}  {name} — {fl}: {streak}wk  [{pairs}]")
    else:
        print("  None.")

    print("\n" + SEP)
    print(f"  ▼  LOW VOLUME ALERTS — {LOW_VOL_WARN}+ CONSECUTIVE WEEKS (NP < {LOW_VOL_MIN})")
    print(SEP)
    lv_alerts = report.get("alerts_lowvol", [])
    if lv_alerts:
        for name, fk, fl, streak, week_vals, week_dates in lv_alerts:
            full_wv = report["clients"][name]["flags"]["low_volume"]["week_values"]
            np_pat  = "  ".join(
                f"[{v}]" if int(v) < LOW_VOL_MIN else f" {v} "
                for v in full_wv
            )
            alert = "★★ CRITICAL" if streak >= LOW_VOL_CRIT else "★  WARNING "
            print(f"  {alert}  {name}: {streak}wk  NP pattern: {np_pat}")
    else:
        print("  None.")

    print("\n" + SEP)
    print("  DATA COVERAGE")
    print(SEP)
    frontier = coverage.get("frontier")
    behind   = coverage.get("behind_frontier", [])
    gaps     = coverage.get("gap_clients", [])
    if frontier:
        print(f"  Frontier (most recent entry across all clients): {frontier}")
    if behind:
        print(f"\n  ⚠ Behind frontier:")
        for name, wks, le in behind:
            print(f"    {name}: {wks}wk behind (last: {le})")
    if gaps:
        print(f"\n  ⚠ Interior gaps:")
        for name, missing in gaps:
            print(f"    {name}: {', '.join(missing)}")
    if not behind and not gaps:
        print("  ✓ No coverage gaps.")

    reset_pending = [(n, cd) for n, cd in sorted(report["clients"].items()) if cd.get("reset_needed")]
    print("\n" + SEP)
    print("  ⚙  STRATEGY RESETS PENDING  (dry run — no writes to GHL or col M)")
    print(SEP)
    if reset_pending:
        for rname, rcdata in reset_pending:
            sc_str = rcdata.get("strategy_changed_on", "?")
            print(f"  {rname}: GHL → 0  |  col M ← {sc_str}  |  note queued")
    else:
        print("  None pending.")

    print("\n" + SEP)
    print("  GHL COUNTER SYNC PREVIEW  (dry run — nothing written to GHL)")
    print("  Counter = overall consecutive weeks with ANY breach flag  ·  drops to 0 on recovery")
    print(SEP)
    print(f"  {'Client':<42} {'GHL Now':>8} {'→ Write':>8}  GHL Automation")
    print(f"  {'-'*42} {'-'*8} {'-'*8}  {'-'*32}")
    for name, cdata in sorted(report["clients"].items()):
        ghl_c  = cdata.get("ghl_counter")
        streak = cdata.get("current_streak", 0)
        ghl_s  = str(ghl_c)  if ghl_c  is not None else "N/A"
        wrt_s  = str(streak)
        delta  = (f"  ↓{ghl_c - streak}" if ghl_c is not None and ghl_c > streak else
                  f"  ↑{streak - ghl_c}"  if ghl_c is not None and streak > ghl_c else
                  "  =" if ghl_c is not None else "")
        auto   = _ghl_auto_label(streak)
        print(f"  {name:<42} {ghl_s:>8} {wrt_s:>8}  {auto}{delta}")
        if cdata.get("reset_needed"):
            print(f"  {'':42}  [⚙ STRATEGY RESET will also fire: GHL → 0]")


# ── Inactivity alert helpers ──────────────────────────────────────────────────

def _read_inactivity_log() -> dict:
    try:
        if INACTIVITY_LOG_PATH.exists():
            return json.loads(INACTIVITY_LOG_PATH.read_text(encoding="utf-8"))
        return {}
    except Exception:
        return {}


def _write_inactivity_log(log: dict) -> None:
    try:
        INACTIVITY_LOG_PATH.write_text(
            json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception as e:
        print(f"[breach] inactivity log write error: {e}")


def _throttle_wait_and_record() -> None:
    try:
        elapsed = DRIP_DELAY + 1
        if THROTTLE_PATH.exists():
            raw = json.loads(THROTTLE_PATH.read_text(encoding="utf-8"))
            elapsed = time.time() - float(raw.get("ts", 0))
        if elapsed < DRIP_DELAY:
            wait = DRIP_DELAY - elapsed
            print(f"[breach] throttle: waiting {int(wait)}s...")
            time.sleep(wait)
        THROTTLE_PATH.write_text(json.dumps({"ts": time.time()}), encoding="utf-8")
    except Exception as e:
        print(f"[breach] throttle error: {e}")


def _send_inactivity_alerts(report: dict, clients: list, settings: dict) -> None:
    from ghl_client import get_has_recent_ghl_activity
    threeup_api_key = settings.get("threeup_api_key", "")
    if not threeup_api_key:
        print("[breach] inactivity check: no threeup_api_key — skipping")
        return
    client_map  = {c["name"]: c for c in clients}
    log         = _read_inactivity_log()
    run_date    = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d")
    alerts_sent = 0

    for name, cdata in sorted(report["clients"].items()):
        if cdata.get("data_weeks", 0) < INACTIVITY_MIN_WEEKS:
            continue
        last2 = cdata.get("last2_np_tl", [])
        if len(last2) < 2:
            continue
        client_active = any(np > 0 or tl > 0 for np, tl in last2)
        if client_active:
            if log.get(name, {}).get("currently_inactive"):
                log[name]["currently_inactive"] = False
                print(f"[breach] inactivity: {name} — recovered, flag reset")
            continue
        if log.get(name, {}).get("currently_inactive"):
            print(f"[breach] inactivity: {name} — NP=TL=0 but already alerted this spell, skipping")
            continue
        c = client_map.get(name, {})
        api_key     = c.get("ghl_api_key", "")
        location_id = c.get("location_id", "")
        if not api_key or not location_id:
            print(f"[breach] inactivity: {name} — no api_key/location_id, skipping")
            continue
        has_ghl_activity = get_has_recent_ghl_activity(
            api_key, location_id, window_days=14, client_name=name
        )
        if has_ghl_activity:
            print(
                f"[breach] inactivity: {name} — NP=TL=0 but GHL has activity"
                f" → likely setup issue, not abandonment"
            )
            continue
        cid = c.get("contact_id", "")
        if not cid or cid in _KNOWN_400_CONTACT_IDS:
            print(f"[breach] inactivity: {name} — no contact_id, skipping")
            continue
        note_text = (
            f"\U0001f6d1 INACTIVE — {name}: 0 new patients tagged and 0 trigger link clicks "
            f"for 2 consecutive weeks, with no other GHL activity detected. "
            f"The clinic may have stopped using the system. ({run_date})"
        )
        _throttle_wait_and_record()
        ok = _ghl_add_note(threeup_api_key, cid, note_text)
        if ok:
            if name not in log:
                log[name] = {}
            log[name]["last_alert_date"]    = run_date
            log[name]["currently_inactive"] = True
            alerts_sent += 1
            print(f"[breach] inactivity ⚠ ALERT: {name} → note sent ✅")
        else:
            print(f"[breach] inactivity ⚠ ALERT: {name} → note FAILED")
        time.sleep(0.5)

    _write_inactivity_log(log)
    print(f"[breach] inactivity check done: {alerts_sent} alert(s) sent this run")


# ── Public entry point ────────────────────────────────────────────────────────

def run(clients: list, settings: dict, dry_run: bool = False) -> dict:
    """
    Main entry point. Called by main.py --breach-audit.

    clients:  full client list (churned clients filtered internally)
    settings: dict from config/settings.json
    dry_run:  if True, print full report + email preview; skip sheet write and send
    """
    from notifier_email import send_email

    threeup_api_key = settings.get("threeup_api_key", "")
    if not threeup_api_key:
        print("[breach] WARNING: threeup_api_key missing — GHL counter reads will fail")

    try:
        sheet_ids = json.loads(SHEET_IDS_PATH.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[breach] Cannot load sheet_ids.json: {e}")
        return {}

    non_churned = [c for c in clients if not c.get("churned")]
    print(f"[breach] Starting breach audit — {len(non_churned)} non-churned clients")

    report = _build_report(non_churned, sheet_ids, threeup_api_key)

    n4 = len({e[0] for e in report["alerts_4plus"]})
    n3 = len({e[0] for e in report["alerts_3plus"]}) - n4
    cv = report.get("coverage", {})
    print(f"[breach] {n4} critical (4+wk), {n3} warning (3wk), "
          f"{len(cv.get('behind_frontier', []))} behind frontier, "
          f"{len(cv.get('gap_clients', []))} interior gaps, "
          f"{len(report['errors'])} errors")

    if dry_run:
        subject, body = _compose_email(report)
        _print_report(report)
        print("\n" + "=" * 70)
        print("DRY RUN — sheet NOT written, email NOT sent, GHL NOT written")
        print(f"Subject: {subject}")
        print("=" * 70)
        print(body)
        return report

    write_breach_to_sheet(report)

    # ── GHL counter field write (field only — no note, no workflow trigger) ────
    contact_map = {c["name"]: c.get("contact_id", "") for c in non_churned}
    ghl_errors  = {}
    ghl_ok = ghl_field_fail = 0

    for name, cdata in sorted(report["clients"].items()):
        cid = contact_map.get(name, "")
        if not cid or cid in _KNOWN_400_CONTACT_IDS:
            continue
        streak  = cdata.get("current_streak", 0)
        written = min(streak, 4)

        ok_field = _ghl_write_counter(threeup_api_key, cid, written)
        if not ok_field:
            ghl_field_fail += 1
            ghl_errors[name] = f"GHL field write FAILED (contact: {cid}, value: {written})"
            logger.error(f"[breach] !! FIELD WRITE FAILED — {name} ({cid})")
        else:
            ghl_ok += 1
        time.sleep(0.10)

    report["ghl_errors"] = ghl_errors
    print(f"[breach] GHL counter writes: {ghl_ok} ok  ·  {ghl_field_fail} failed")
    if ghl_errors:
        print(f"[breach] !! {len(ghl_errors)} GHL failure(s) — check email for details")

    # ── Strategy-change GHL counter resets ───────────────────────────────────
    # First run after strategy_changed_on is set: reset GHL to 0 and stamp col M.
    # Idempotent: reset_needed is False when col L == col M (already done this run).
    reset_writes = []
    for name, cdata in sorted(report["clients"].items()):
        if not cdata.get("reset_needed"):
            continue
        cid = contact_map.get(name, "")
        if not cid or cid in _KNOWN_400_CONTACT_IDS:
            logger.warning(f"[breach] strategy reset skipped: {name} — no contact_id")
            continue
        ok_field = _ghl_write_counter(threeup_api_key, cid, 0)
        if ok_field:
            sc_date_str = cdata.get("strategy_changed_on", "")
            row = cdata.get("sc_sheet_row")
            if row:
                reset_writes.append((row, sc_date_str))
            print(f"[breach] Strategy reset: {name} GHL→0")
        else:
            print(f"[breach] !! Strategy reset field failed: {name}")
        time.sleep(0.10)
    _write_strategy_reset_done(reset_writes)

    _send_inactivity_alerts(report, non_churned, settings)

    from system_health import health_block, record_run
    record_run("breach_monitor")
    subject, body = _compose_email(report)
    send_email(subject, health_block(settings) + body)
    print(f"[breach] Done — email sent, sheet updated")
    return report
