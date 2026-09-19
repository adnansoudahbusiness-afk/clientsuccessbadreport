#!/usr/bin/env python3
"""
ghl_backfill.py — Backfill real GHL data into existing placeholder sheet rows.
Finds each row by its date_range label in col A and overwrites D-AD in-place.
Does NOT call append_weekly_row — safe for non-last rows.
"""

import sys
import os
import json
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
AGENT_DIR = ROOT / "agent"
if str(AGENT_DIR) not in sys.path:
    sys.path.insert(0, str(AGENT_DIR))

import pytz
from dotenv import load_dotenv
load_dotenv(ROOT / ".env")

from google.oauth2 import service_account
from googleapiclient.discovery import build

from ghl_client import (
    get_new_patients,
    get_returning_patients_tagged,
    get_trigger_link_clicks,
    get_appointments,
    get_loyalty_points,
    get_active_workflow,
    get_active_snippet,
)

AMMAN_TZ = pytz.timezone("Asia/Amman")
CREDENTIALS_PATH = ROOT / "config" / "credentials.json"
SHEET_IDS_PATH   = ROOT / "config" / "sheet_ids.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ── Helpers ────────────────────────────────────────────────────────────────────

def _localize(*ymdhms) -> datetime:
    return AMMAN_TZ.localize(datetime(*ymdhms))


def _fmt_date_range(start: datetime, end: datetime) -> str:
    day_names = {5: "Sat", 6: "Sun", 0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri"}
    s_day = day_names.get(start.weekday(), "")
    e_day = day_names.get(end.weekday(), "")
    return f"{s_day} {start.day}/{start.month} – {e_day} {end.day}/{end.month}"


def _get_sheets():
    creds = service_account.Credentials.from_service_account_file(
        str(CREDENTIALS_PATH), scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _load_sheet_ids() -> dict:
    with open(SHEET_IDS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _get_sid(client_name: str, sheet_ids: dict) -> str:
    entry = sheet_ids.get(client_name)
    if isinstance(entry, dict):
        return entry.get("spreadsheet_id", "")
    return entry or ""


def _find_row_by_label(sheets, sid: str, label: str) -> int:
    """Return 1-based row number where col A matches label. 0 = not found."""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=sid,
        range="Sheet1!A:A",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    rows = result.get("values", [])
    for i, row in enumerate(rows):
        if row and str(row[0]).strip() == label.strip():
            return i + 1
    return 0


def _write_data_to_row(sheets, sid: str, row_num: int, data_cols: list):
    """Overwrite cols D-AD (27 values) of the given 1-based row."""
    sheets.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"Sheet1!D{row_num}:AD{row_num}",
        valueInputOption="USER_ENTERED",
        body={"values": [data_cols]},
    ).execute()


def _infer_strategy(snippet_name, snippet_value) -> str:
    if not snippet_name and not snippet_value:
        return "Unknown"
    snippet_name  = snippet_name  or ""
    snippet_value = snippet_value or ""
    text = (snippet_name + " " + snippet_value).lower()
    if "2" in text and ("personal" in text or "doctor" in text or "دكتور" in text or "اسم" in text):
        return "Strategy 1"
    if "2" in text and ("warm" in text or "دفء" in text):
        return "Strategy 2"
    if "1" in text and ("thank" in text or "شكر" in text or "دكتور" in text):
        return "Strategy 3"
    if "1" in text and ("satisfaction" in text or "رضا" in text or "direct" in text):
        return "Strategy 4"
    return snippet_name if snippet_name else "Unknown"


def _build_data_columns(d: dict) -> list:
    """Return 27 values for cols D(3)–AD(29), matching append_weekly_row layout."""
    return [
        d.get("new_patients_this_week", 0),           # D   3
        d.get("new_patients_last_week", 0),            # E   4
        d.get("new_patients_growth", 0),               # F   5
        d.get("returning_patients", 0),                # G   6
        d.get("trigger_links", 0),                     # H   7
        "",                                            # I   8  Hook % (computed later)
        "",                                            # J   9  Offer %
        "",                                            # K  10  Hook Flag
        "",                                            # L  11  Offer Flag
        d.get("appointments_booked", 0),               # M  12
        d.get("appointments_confirmed", 0),            # N  13
        d.get("appointments_cancelled", 0),            # O  14
        d.get("cancellation_rate", 0),                 # P  15
        "",                                            # Q  16  Cancel Flag
        d.get("appointments_confirmed_last_week", 0),  # R  17
        d.get("loyalty_this", 0),                      # S  18
        d.get("appointments_growth", 0),               # T  19
        d.get("loyalty_prior", 0),                     # U  20
        "",                                            # V  21  Seasonal Flag
        d.get("workflow_name", ""),                    # W  22
        d.get("snippet_name", ""),                     # X  23
        d.get("current_strategy", ""),                 # Y  24
        "",                                            # Z  25  Recommended Strategy
        "",                                            # AA 26  Diagnosis Notes
        "",                                            # AB 27  Change Type
        "",                                            # AC 28  Form Submitted?
        "",                                            # AD 29  Submission Timestamp
    ]


def pull_week(client: dict,
              this_start: datetime, this_end: datetime,
              last_start: datetime, last_end: datetime) -> dict:
    name        = client["name"]
    api_key     = client["ghl_api_key"]
    location_id = client["location_id"]
    np_tag      = client.get("new_patient_tag") or None
    tl_tag      = client.get("trigger_link_tag") or None

    np_this   = get_new_patients(api_key, location_id, this_start, this_end, name, new_patient_tag=np_tag)
    np_last   = get_new_patients(api_key, location_id, last_start, last_end, name, new_patient_tag=np_tag)
    returning = get_returning_patients_tagged(api_key, location_id, this_start, this_end, name)
    tl        = get_trigger_link_clicks(api_key, location_id, this_start, this_end, name, trigger_link_tag=tl_tag)

    appts_this = get_appointments(api_key, location_id, this_start, this_end, name)
    appts_last = get_appointments(api_key, location_id, last_start, last_end, name)

    loyalty_this = get_loyalty_points(api_key, location_id, this_start, this_end, name)
    loyalty_last = get_loyalty_points(api_key, location_id, last_start, last_end, name)

    workflow_name = get_active_workflow(api_key, location_id, name)
    snippet       = get_active_snippet(api_key, location_id, workflow_name, name) if workflow_name else None
    snippet_name  = snippet["name"]  if snippet else ""
    snippet_value = snippet["value"] if snippet else ""
    strategy      = _infer_strategy(snippet_name, snippet_value)

    appts_this_confirmed = appts_this.get("confirmed", appts_this.get("total", 0))
    appts_last_confirmed = appts_last.get("confirmed", appts_last.get("total", 0))

    np_growth    = round((np_this - np_last) / np_last * 100, 2) if np_last > 0 else 0
    appts_growth = round((appts_this_confirmed - appts_last_confirmed) / appts_last_confirmed * 100, 2) if appts_last_confirmed > 0 else 0

    return {
        "date_range":                      _fmt_date_range(this_start, this_end),
        "new_patients_this_week":           np_this,
        "new_patients_last_week":           np_last,
        "new_patients_growth":              np_growth,
        "returning_patients":               returning,
        "trigger_links":                    tl,
        "appointments_booked":              appts_this.get("booked", 0),
        "appointments_confirmed":           appts_this_confirmed,
        "appointments_cancelled":           appts_this.get("cancelled", 0),
        "cancellation_rate":                appts_this.get("cancellation_rate", 0),
        "appointments_confirmed_last_week": appts_last_confirmed,
        "appointments_growth":              appts_growth,
        "loyalty_this":                     loyalty_this,
        "loyalty_prior":                    loyalty_last,
        "workflow_name":                    workflow_name or "",
        "snippet_name":                     snippet_name,
        "current_strategy":                 strategy,
    }


# ── Client / week schedule ─────────────────────────────────────────────────────
# Each week: (this_start_ymdhms, this_end_ymdhms, last_start_ymdhms, last_end_ymdhms)

CLIENTS_WEEKS = [
    {
        "name": "Dr Basil Obeidat",
        "weeks": [
            ((2026,9,5,0,0,0),  (2026,9,11,23,59,59), (2026,8,29,0,0,0), (2026,9,4,23,59,59)),   # W25
            ((2026,9,12,0,0,0), (2026,9,18,23,59,59), (2026,9,5,0,0,0),  (2026,9,11,23,59,59)),   # W26
        ],
    },
    {
        "name": "Dr Mahmoud Malkawi",
        "weeks": [
            ((2026,9,5,0,0,0),  (2026,9,11,23,59,59), (2026,8,29,0,0,0), (2026,9,4,23,59,59)),
            ((2026,9,12,0,0,0), (2026,9,18,23,59,59), (2026,9,5,0,0,0),  (2026,9,11,23,59,59)),
        ],
    },
    {
        "name": "Dr Hisham Qaisi",
        "weeks": [
            ((2026,8,1,0,0,0),  (2026,8,7,23,59,59),  (2026,7,25,0,0,0), (2026,7,31,23,59,59)),   # W20
            ((2026,8,8,0,0,0),  (2026,8,14,23,59,59), (2026,8,1,0,0,0),  (2026,8,7,23,59,59)),    # W21
            ((2026,8,15,0,0,0), (2026,8,21,23,59,59), (2026,8,8,0,0,0),  (2026,8,14,23,59,59)),   # W22
            ((2026,8,22,0,0,0), (2026,8,28,23,59,59), (2026,8,15,0,0,0), (2026,8,21,23,59,59)),   # W23
            ((2026,8,29,0,0,0), (2026,9,4,23,59,59),  (2026,8,22,0,0,0), (2026,8,28,23,59,59)),   # W24
            ((2026,9,5,0,0,0),  (2026,9,11,23,59,59), (2026,8,29,0,0,0), (2026,9,4,23,59,59)),    # W25
            ((2026,9,12,0,0,0), (2026,9,18,23,59,59), (2026,9,5,0,0,0),  (2026,9,11,23,59,59)),   # W26
        ],
    },
    {
        "name": "Dr Shadi AlHourani",
        "weeks": [
            ((2026,9,5,0,0,0),  (2026,9,11,23,59,59), (2026,8,29,0,0,0), (2026,9,4,23,59,59)),
            ((2026,9,12,0,0,0), (2026,9,18,23,59,59), (2026,9,5,0,0,0),  (2026,9,11,23,59,59)),
        ],
    },
    {
        "name": "Dr Shadi Al Khateeb",
        "weeks": [
            ((2026,7,18,0,0,0), (2026,7,24,23,59,59), (2026,7,11,0,0,0), (2026,7,17,23,59,59)),   # W18
            ((2026,7,25,0,0,0), (2026,7,31,23,59,59), (2026,7,18,0,0,0), (2026,7,24,23,59,59)),   # W19
            ((2026,8,1,0,0,0),  (2026,8,7,23,59,59),  (2026,7,25,0,0,0), (2026,7,31,23,59,59)),   # W20
            ((2026,8,8,0,0,0),  (2026,8,14,23,59,59), (2026,8,1,0,0,0),  (2026,8,7,23,59,59)),    # W21
            ((2026,8,15,0,0,0), (2026,8,21,23,59,59), (2026,8,8,0,0,0),  (2026,8,14,23,59,59)),   # W22
            ((2026,8,22,0,0,0), (2026,8,28,23,59,59), (2026,8,15,0,0,0), (2026,8,21,23,59,59)),   # W23
            ((2026,8,29,0,0,0), (2026,9,4,23,59,59),  (2026,8,22,0,0,0), (2026,8,28,23,59,59)),   # W24
            ((2026,9,5,0,0,0),  (2026,9,11,23,59,59), (2026,8,29,0,0,0), (2026,9,4,23,59,59)),    # W25
            ((2026,9,12,0,0,0), (2026,9,18,23,59,59), (2026,9,5,0,0,0),  (2026,9,11,23,59,59)),   # W26
        ],
    },
]


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    sheets    = _get_sheets()
    sheet_ids = _load_sheet_ids()

    with open(ROOT / "config" / "clients.json", "r", encoding="utf-8") as f:
        client_map = {c["name"]: c for c in json.load(f)}

    results = {}

    for client_def in CLIENTS_WEEKS:
        cname  = client_def["name"]
        client = client_map.get(cname)
        sid    = _get_sid(cname, sheet_ids)

        if not client:
            print(f"\n[ERROR] Client not found in clients.json: {cname}")
            continue
        if not sid:
            print(f"\n[ERROR] No sheet ID for: {cname}")
            continue

        print(f"\n{'='*60}")
        print(f"  {cname}")
        print(f"  Sheet: {sid}")
        print(f"{'='*60}")

        results[cname] = []

        for week_t in client_def["weeks"]:
            ts_t, te_t, ls_t, le_t = week_t
            this_start = _localize(*ts_t)
            this_end   = _localize(*te_t)
            last_start = _localize(*ls_t)
            last_end   = _localize(*le_t)

            label = _fmt_date_range(this_start, this_end)
            print(f"\n  [{label}]")

            row_num = _find_row_by_label(sheets, sid, label)
            time.sleep(0.5)

            if row_num == 0:
                print(f"    [WARN] Row not found — skipping")
                results[cname].append({"week": label, "status": "ROW_NOT_FOUND"})
                continue

            print(f"    Row {row_num} found — pulling GHL...")

            try:
                data = pull_week(client, this_start, this_end, last_start, last_end)
                cols = _build_data_columns(data)
                _write_data_to_row(sheets, sid, row_num, cols)
                time.sleep(1)

                np   = data["new_patients_this_week"]
                tl   = data["trigger_links"]
                appt = data["appointments_confirmed"]
                print(f"    [OK] NP={np}  TL={tl}  Appts={appt}")
                results[cname].append({
                    "week": label, "row": row_num,
                    "NP": np, "TL": tl, "Appts": appt, "status": "OK"
                })
            except Exception as e:
                print(f"    [ERROR] {e}")
                results[cname].append({"week": label, "status": f"ERROR: {e}"})

            time.sleep(2)

        time.sleep(3)

    # ── Summary ──────────────────────────────────────────────────────────────
    print("\n\n" + "="*60)
    print("  BACKFILL SUMMARY")
    print("="*60)
    for cname, weeks in results.items():
        print(f"\n{cname}:")
        for w in weeks:
            if w["status"] == "OK":
                print(f"  {w['week']:<32} NP={w['NP']:>3}  TL={w['TL']:>3}  Appts={w['Appts']:>3}")
            else:
                print(f"  {w['week']:<32} {w['status']}")


if __name__ == "__main__":
    main()
