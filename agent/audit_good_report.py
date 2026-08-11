import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

"""
audit_good_report.py — Full audit of the Good Report system.
Reads GHL contacts and AM dates sheet, cross-references, and reports.
Run from project root or agent/ directory — paths are resolved automatically.
"""

import json
import os
import requests
from datetime import datetime, timedelta, date

from google.oauth2 import service_account
from googleapiclient.discovery import build

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)

CREDENTIALS_PATH = os.path.join(_ROOT, "config", "credentials.json")
SETTINGS_PATH    = os.path.join(_ROOT, "config", "settings.json")
CLIENTS_PATH     = os.path.join(_ROOT, "config", "clients.json")
LOGS_DIR         = os.path.join(_ROOT, "logs")

PYTHON_EXE    = r"C:\Users\User\AppData\Local\Python\pythoncore-3.14-64\python.exe"
AM_SHEET_ID   = "12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA"
GHL_BASE      = "https://services.leadconnectorhq.com"
LAST_SENT_ID  = "M70Sd18pm9ZPg35ecix1"
CAO_REPORT_ID = "Y1lUr9X7ACLNmfsXmsCX"
MONTH_NAMES   = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
_SKIP = {"waiting", "pause", ""}


def _tab_to_month(title: str):
    """Return canonical month name for a tab title, or None if not a month tab.
    Handles both 'June' and 'June 2026' style names."""
    parts = title.strip().split()
    if parts and parts[0] in set(MONTH_NAMES):
        return parts[0]
    return None


# ── Google Sheets ──────────────────────────────────────────────────────────────

def _get_sheets_service():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_PATH, scopes=scopes
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _parse_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    try:
        parts = s.split("/")
        if len(parts) == 2:
            return date(2026, int(parts[1]), int(parts[0]))
    except Exception:
        pass
    return None


def _read_tab(service, title):
    """Parse one AM sheet tab. Returns list of entry dicts."""
    try:
        data = service.spreadsheets().values().get(
            spreadsheetId=AM_SHEET_ID,
            range=f"'{title}'!A:H",
        ).execute()
    except Exception as e:
        print(f"  [WARN] Tab '{title}': {e}")
        return []

    rows   = data.get("values", [])
    result = []
    for i, row in enumerate(rows):
        if i == 0:
            continue  # header
        if len(row) < 6:
            continue
        client_name = str(row[0]).strip()
        sus_raw     = str(row[1]).strip() if len(row) > 1 else ""
        renewal_raw = str(row[4]).strip() if len(row) > 4 else ""
        payment_raw = str(row[5]).strip() if len(row) > 5 else ""
        term        = str(row[6]).strip() if len(row) > 6 else ""

        if not client_name or sus_raw.lower() in _SKIP:
            continue
        if renewal_raw.lower() in _SKIP or payment_raw.lower() in _SKIP:
            continue

        sus     = _parse_date(sus_raw)
        renewal = _parse_date(renewal_raw)
        payment = _parse_date(payment_raw)
        if not sus or not renewal or not payment:
            print(f"  [SKIP] {client_name}: bad dates sus='{sus_raw}' renewal='{renewal_raw}' payment='{payment_raw}'")
            continue

        result.append({
            "client_name":  client_name,
            "first_sus":    sus,
            "renewal_date": renewal,
            "payment_date": payment,
            "term":         term,
            "tab":          title,
        })
    return result


def load_am_entries(service, today):
    """
    Read current month + next month tabs.
    Falls back to walking backward if neither tab exists.
    Deduplicates by client_name (current month wins).
    """
    try:
        spreadsheet    = service.spreadsheets().get(spreadsheetId=AM_SHEET_ID).execute()
        available_titles = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]
        month_to_title   = {_tab_to_month(t): t for t in available_titles if _tab_to_month(t)}
    except Exception as e:
        print(f"[ERROR] Could not open AM sheet: {e}")
        return []

    current = MONTH_NAMES[today.month - 1]
    nxt     = MONTH_NAMES[today.month % 12]   # Dec(12) % 12 = 0 → January
    to_read = [month_to_title[m] for m in [current, nxt] if m in month_to_title]

    if not to_read:
        for offset in range(1, 12):
            candidate = MONTH_NAMES[(today.month - 1 - offset) % 12]
            if candidate in month_to_title:
                to_read = [month_to_title[candidate]]
                break

    seen, entries = set(), []
    for tab in to_read:
        for entry in _read_tab(service, tab):
            key = entry["client_name"].lower()
            if key not in seen:
                seen.add(key)
                entries.append(entry)
    return entries


# ── GHL ────────────────────────────────────────────────────────────────────────

def _fetch_ghl_fields(contact_id, api_key):
    """Return (last_sent_str_or_None, preview_str_or_None)."""
    try:
        r = requests.get(
            f"{GHL_BASE}/contacts/{contact_id}",
            headers={"Authorization": f"Bearer {api_key}", "Version": "2021-07-28"},
            timeout=20,
        )
        if r.status_code != 200:
            return None, f"HTTP {r.status_code}"
        body   = r.json()
        fields = (body.get("contact", {}).get("customFields", [])
                  or body.get("customFields", []))
        last_sent = None
        preview   = None
        for f in fields:
            fid = f.get("id", "")
            v   = str(f.get("value", "")).strip()
            if fid == LAST_SENT_ID:
                last_sent = v or None
            elif fid == CAO_REPORT_ID and v:
                preview = (v[:50] + "…") if len(v) > 50 else v
        return last_sent, preview
    except Exception as e:
        return None, f"ERR: {e}"


# ── Matching ───────────────────────────────────────────────────────────────────

def _match_client(am_name, clients):
    """Identical 4-pass logic as good_report_engine.match_client."""
    am = am_name.strip().lower()
    for c in clients:
        if c.get("am_name", "").strip().lower() == am:   return c
    for c in clients:
        if c["name"].strip().lower() == am:               return c
    for c in clients:
        if am in c["name"].strip().lower():               return c
    for c in clients:
        if c["name"].strip().lower() in am:               return c
    return None


def _find_am_for_client(client, am_entries):
    """Reverse lookup: find the AM sheet entry that corresponds to a client."""
    am_name = client.get("am_name", "").strip().lower()
    name    = client.get("name", "").strip().lower()
    for entry in am_entries:
        cn = entry["client_name"].strip().lower()
        if am_name and cn == am_name:    return entry
    for entry in am_entries:
        cn = entry["client_name"].strip().lower()
        if cn == name:                   return entry
    for entry in am_entries:
        cn = entry["client_name"].strip().lower()
        if am_name and cn in am_name:    return entry
    for entry in am_entries:
        cn = entry["client_name"].strip().lower()
        if name and cn in name:          return entry
    for entry in am_entries:
        cn = entry["client_name"].strip().lower()
        if am_name and am_name in cn:    return entry
    for entry in am_entries:
        cn = entry["client_name"].strip().lower()
        if name and name in cn:          return entry
    return None


# ── Schedule math ──────────────────────────────────────────────────────────────

def _calc_schedule(first_sus, renewal_date, payment_date):
    """
    Returns sorted list of (date, type_str).
    Sustainability: first_sus + every 14 days while <= renewal - 6d.
    Renewal reminder: payment - 6d.
    """
    schedule = []
    d = first_sus
    while d <= renewal_date - timedelta(days=6):
        schedule.append((d, "sustainability"))
        d += timedelta(days=14)
    schedule.append((payment_date - timedelta(days=6), "renewal"))
    return schedule  # naturally sorted forward


def _classify(schedule, last_sent_str, today):
    """
    sent    — past dates where date <= last_sent (covered by prior send history)
    missed  — past dates where date > last_sent, or all if last_sent is None
    upcoming — future dates within 30 days of today
    """
    last_sent = _parse_date(last_sent_str) if last_sent_str else None
    sent, missed, upcoming = [], [], []
    for d, t in schedule:
        if d > today:
            if d <= today + timedelta(days=30):
                upcoming.append((d, t))
        else:
            if last_sent and d <= last_sent:
                sent.append((d, t))
            else:
                missed.append((d, t))
    return sent, missed, upcoming


def _derive_status(am_entry, last_sent, missed):
    if not am_entry:
        return "UNKNOWN"
    if not last_sent:
        return "UNKNOWN"
    return "BEHIND" if missed else "ON TRACK"


# ── Main audit ─────────────────────────────────────────────────────────────────

def run_audit():
    os.makedirs(LOGS_DIR, exist_ok=True)
    today = datetime.now().date()
    lines = []

    def p(text=""):
        print(text)
        lines.append(str(text))

    p("=" * 70)
    p(f"  GOOD REPORT AUDIT — {today.isoformat()}")
    p("=" * 70)
    p()

    with open(SETTINGS_PATH, encoding="utf-8") as f:
        settings = json.load(f)
    with open(CLIENTS_PATH, encoding="utf-8") as f:
        clients = json.load(f)

    clients = [c for c in clients if not c.get("churned")]
    api_key = settings["threeup_api_key"]

    # ── 1. Load AM sheet ───────────────────────────────────────────────────────
    p("[1/3] Loading AM dates sheet…")
    service    = _get_sheets_service()
    am_entries = load_am_entries(service, today)
    p(f"      {len(am_entries)} entries found (current + next month tabs).")
    p()

    # ── 2. Read GHL contacts ───────────────────────────────────────────────────
    p("[2/3] Reading GHL contacts…")
    results = []

    for client in clients:
        name       = client.get("name", "?")
        contact_id = client.get("contact_id", "")

        if not contact_id:
            results.append({
                "name": name, "contact_id": None,
                "last_sent": None, "preview": None,
                "am_entry": None, "schedule": [],
                "sent": [], "missed": [], "upcoming": [],
                "status": "UNKNOWN",
            })
            print(f"      ! {name}  [no contact_id]")
            continue

        last_sent, preview = _fetch_ghl_fields(contact_id, api_key)
        am_entry = _find_am_for_client(client, am_entries)

        if am_entry:
            schedule = _calc_schedule(
                am_entry["first_sus"],
                am_entry["renewal_date"],
                am_entry["payment_date"],
            )
            sent, missed, upcoming = _classify(schedule, last_sent, today)
        else:
            schedule, sent, missed, upcoming = [], [], [], []

        status = _derive_status(am_entry, last_sent, missed)

        results.append({
            "name":       name,
            "contact_id": contact_id,
            "last_sent":  last_sent,
            "preview":    preview,
            "am_entry":   am_entry,
            "schedule":   schedule,
            "sent":       sent,
            "missed":     missed,
            "upcoming":   upcoming,
            "status":     status,
        })
        print(f"      ✓ {name}  last_sent={last_sent or '—'}  "
              f"missed={len(missed)}  upcoming={len(upcoming)}")

    p()

    # ── 3. Cross-reference ─────────────────────────────────────────────────────
    p("[3/3] Cross-referencing…")
    p()

    no_am_entry   = [r for r in results if not r["am_entry"]]
    no_contact_id = [r for r in results if not r["contact_id"]]
    unmatched_am  = [
        e["client_name"]
        for e in am_entries
        if not _match_client(e["client_name"], clients)
    ]

    # ══ SECTION 1 ══════════════════════════════════════════════════════════════
    p("═" * 70)
    p("  SECTION 1 — CLIENT STATUS SUMMARY")
    p("═" * 70)
    p()
    p(f"  {'CLIENT':<40} {'LAST SENT':<14} {'STATUS'}")
    p(f"  {'-'*40} {'-'*14} {'-'*12}")
    for r in sorted(results, key=lambda x: x["name"]):
        ls = r["last_sent"] or "—"
        p(f"  {r['name']:<40} {ls:<14} {r['status']}")
        if r["preview"]:
            p(f"  {'':>43}└ {r['preview']}")
    p()

    # ══ SECTION 2 ══════════════════════════════════════════════════════════════
    p("═" * 70)
    p("  SECTION 2 — MISSED MESSAGES")
    p("═" * 70)
    p()
    any_missed = False
    for r in sorted(results, key=lambda x: x["name"]):
        if r["missed"]:
            any_missed = True
            p(f"  {r['name']}")
            for d, t in r["missed"]:
                ago = (today - d).days
                p(f"    • {d}  [{t}]  ({ago} days ago)")
    if not any_missed:
        p("  No missed messages detected.")
    p()

    # ══ SECTION 3 ══════════════════════════════════════════════════════════════
    p("═" * 70)
    p("  SECTION 3 — UPCOMING SCHEDULE (next 30 days)")
    p("═" * 70)
    p()
    all_upcoming = []
    for r in results:
        for d, t in r["upcoming"]:
            all_upcoming.append((d, r["name"], t, (d - today).days))
    all_upcoming.sort()
    if all_upcoming:
        p(f"  {'DATE':<14} {'CLIENT':<40} {'TYPE':<16} {'DAYS'}")
        p(f"  {'-'*14} {'-'*40} {'-'*16} {'-'*4}")
        for d, name, t, dfn in all_upcoming:
            p(f"  {str(d):<14} {name:<40} {t:<16} +{dfn}")
    else:
        p("  No upcoming messages in next 30 days.")
    p()

    # ══ SECTION 4 ══════════════════════════════════════════════════════════════
    p("═" * 70)
    p("  SECTION 4 — SYSTEM HEALTH")
    p("═" * 70)
    p()
    on_track = [r for r in results if r["status"] == "ON TRACK"]
    behind   = [r for r in results if r["status"] == "BEHIND"]
    unknown  = [r for r in results if r["status"] == "UNKNOWN"]

    p(f"  Total clients in clients.json  : {len(clients)}")
    p(f"  Total entries in AM sheet      : {len(am_entries)}")
    p()
    p(f"  ON TRACK  ({len(on_track)})")
    for r in on_track:
        p(f"    ✓ {r['name']}  (last sent {r['last_sent']})")
    p()
    p(f"  BEHIND    ({len(behind)})")
    for r in behind:
        p(f"    ✗ {r['name']}  — {len(r['missed'])} missed  (last sent {r['last_sent'] or '—'})")
    p()
    p(f"  UNKNOWN   ({len(unknown)})")
    for r in unknown:
        reason = "no AM entry" if not r["am_entry"] else "never sent"
        p(f"    ? {r['name']}  ({reason})")
    p()
    p(f"  No AM entry ({len(no_am_entry)})")
    for r in no_am_entry:
        p(f"    - {r['name']}")
    p()
    p(f"  No GHL contact_id ({len(no_contact_id)})")
    for r in no_contact_id:
        p(f"    - {r['name']}")
    p()
    p(f"  In AM sheet, not matched in clients.json ({len(unmatched_am)})")
    for n in unmatched_am:
        p(f"    - {n}")
    p()

    # ══ SECTION 5 ══════════════════════════════════════════════════════════════
    p("═" * 70)
    p("  SECTION 5 — FORCE SEND COMMANDS")
    p("═" * 70)
    p()
    force_list = [r for r in results if r["missed"]]
    if force_list:
        for r in sorted(force_list, key=lambda x: x["name"]):
            p(f"  # {r['name']}  ({len(r['missed'])} missed message(s))")
            p(f'  {PYTHON_EXE} agent\\main.py --good-report-force "{r["name"]}"')
            p()
    else:
        p("  No force sends required — all clients are on track.")
    p()
    p("=" * 70)

    # ── Save to file ───────────────────────────────────────────────────────────
    out_path = os.path.join(LOGS_DIR, f"audit_{today.isoformat()}.txt")
    with open(out_path, "w", encoding="utf-8") as fout:
        fout.write("\n".join(lines))
    print(f"\n[SAVED] {out_path}")


if __name__ == "__main__":
    run_audit()
