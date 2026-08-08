#!/usr/bin/env python3
"""
generate_month_tab.py — Generate or regenerate a monthly AM communication tab.

Derivation rule (single source of truth = Last Payment in master col J):
  - Last Payment in the PAST  (LP <= today): walk LP forward by term_months
    until next_payment > today.  term_start = previous step.
  - Last Payment in the FUTURE (LP > today): LP IS next_payment;
    term_start = LP - term_months.

Computed values written back to master after each run:
  - Col D (Cadence Anchor): first sustainability of the current term
    (term_start + 14 for auto cadence; fixed day for fixed: pattern)

Humans hand-edit ONLY: Status, Term Months, Term Label,
Sus/Renewal Pattern, Last Payment, Notes.

Safe to re-run: if the month tab already exists, only data rows (row 3+) are
cleared and rewritten.  Title and header rows are preserved.
Writes ONLY to the named month tab and to master col D — no other tabs touched.

Usage:
    python generate_month_tab.py "October 2026"
    python generate_month_tab.py "October 2026" --dry-run
"""

import sys
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
os.chdir(os.path.dirname(_HERE))  # repo root

from datetime import date, timedelta
from calendar import monthrange
from dateutil.relativedelta import relativedelta

from good_report_engine import _get_sheets_service, AM_DATES_SHEET_ID, _parse_date

MASTER_TAB = "Clients Master"

MONTH_TAB_TITLE_ROW = "{label} — Account Manager Communication Dates"
MONTH_TAB_HEADER    = ["Client Name", "Sustainability", "Renewal", "Payment Date", "Term"]

# Column indices in master (0-based)
COL_NAME     = 0
COL_DOCTOR   = 1
COL_STATUS   = 2
COL_ANCHOR   = 3   # Cadence Anchor — COMPUTED, written back by this script
COL_TERM_MO  = 4
COL_TERM_LBL = 5
COL_SUS_PAT  = 6
COL_REN_PAT  = 7
COL_LAST_SUS = 8   # Last Sus — AGENT-WRITTEN by good_report_engine after sends
COL_LAST_PAY = 9
COL_NOTES    = 10


# ── Master reader ─────────────────────────────────────────────────────────────

def _read_master(service) -> list:
    """Return list of client dicts from Clients Master; include 1-based row index."""
    data = service.spreadsheets().values().get(
        spreadsheetId=AM_DATES_SHEET_ID,
        range=f"'{MASTER_TAB}'!A:K",
    ).execute()
    rows = data.get("values", [])
    clients = []
    for i, row in enumerate(rows):
        if i < 2:
            continue  # skip title + header
        if not row or not row[COL_NAME].strip():
            continue

        def col(j, default=""):
            return row[j].strip() if len(row) > j and row[j].strip() else default

        tm_raw = col(COL_TERM_MO)
        try:
            term_months = int(tm_raw) if tm_raw.isdigit() else 0
        except (ValueError, AttributeError):
            term_months = 0

        clients.append({
            "sheet_row":      i + 1,          # 1-based row in spreadsheet
            "am_name":        col(COL_NAME),
            "status":         col(COL_STATUS).lower() or "active",
            "term_months":    term_months,
            "term_label":     col(COL_TERM_LBL),
            "sus_pattern":    col(COL_SUS_PAT) or "auto",
            "renewal_pattern":col(COL_REN_PAT) or "auto",
            "last_payment":   _parse_date(col(COL_LAST_PAY)) if col(COL_LAST_PAY) else None,
        })
    return clients


# ── Derivation helpers ────────────────────────────────────────────────────────

def _derive_term(last_payment, term_months, today=None):
    """
    Return (term_start, next_payment).

    FUTURE (LP > today): LP is next_payment; term_start = LP - N months.
    PAST   (LP <= today): advance LP by N months until next_payment > today.
    """
    if today is None:
        today = date.today()
    if last_payment > today:
        return last_payment - relativedelta(months=term_months), last_payment

    term_start   = last_payment
    next_payment = last_payment + relativedelta(months=term_months)
    while next_payment <= today:
        term_start   = next_payment
        next_payment = next_payment + relativedelta(months=term_months)
    return term_start, next_payment


def _first_sus(term_start, sus_pattern):
    """First sustainability date of a term (= computed cadence anchor)."""
    if sus_pattern.startswith("fixed:"):
        sday = int(sus_pattern.split(":")[1])
        try:
            s = date(term_start.year, term_start.month, sday)
        except ValueError:
            s = date(term_start.year, term_start.month, 28)
        # If fixed day falls before term start, advance to next month
        if s < term_start:
            s = s + relativedelta(months=1)
            try:
                s = date(s.year, s.month, sday)
            except ValueError:
                s = date(s.year, s.month, 28)
        return s
    return term_start + timedelta(days=14)


def _compute_renewal(next_payment, renewal_pattern):
    if renewal_pattern.startswith("fixed:"):
        rday = int(renewal_pattern.split(":")[1])
        return date(next_payment.year, next_payment.month, rday)
    return next_payment - timedelta(days=6)


# ── Row computation ───────────────────────────────────────────────────────────

def _compute_rows(clients, target_year, target_month):
    """
    Return (tab_rows, skipped_list, anchor_updates).

    tab_rows:       list of [name, sus_str, renewal_str, payment_str, term_label]
    anchor_updates: list of (sheet_row, anchor_date_str) for writing back to master
    """
    t_start = date(target_year, target_month, 1)
    t_end   = date(target_year, target_month, monthrange(target_year, target_month)[1])
    today   = date.today()

    tab_rows       = []
    skipped        = []
    anchor_updates = []

    for c in clients:
        if c["status"] == "churned":
            continue
        lp = c["last_payment"]
        tm = c["term_months"]
        if not lp or tm == 0:
            skipped.append(f"{c['am_name']} (no term data)")
            continue

        # Current term (anchors the cadence-anchor write-back)
        term_start_curr, next_pay_curr = _derive_term(lp, tm, today)
        anchor = _first_sus(term_start_curr, c["sus_pattern"])
        anchor_updates.append((c["sheet_row"], anchor.strftime("%d/%m/%Y")))

        # Walk forward until a term window overlaps the target month
        # (monthly clients whose current term ends before the target month
        #  need one or more advances to find their next-cycle dates)
        term_start = term_start_curr
        next_pay   = next_pay_curr
        found = False
        for _ in range(24):
            if term_start > t_end:
                break                       # walked past target month
            if next_pay >= t_start:
                found = True
                break                       # this term overlaps target month
            term_start = next_pay           # advance one term
            next_pay   = next_pay + relativedelta(months=tm)
        if not found:
            continue

        renewal  = _compute_renewal(next_pay, c["renewal_pattern"])
        first_s  = _first_sus(term_start, c["sus_pattern"])

        tab_rows.append([
            c["am_name"],
            first_s.strftime("%d/%m/%Y"),
            renewal.strftime("%d/%m/%Y"),
            next_pay.strftime("%d/%m/%Y"),
            c["term_label"],
        ])

    # Sort by sustainability date (earliest term start first)
    tab_rows.sort(key=lambda r: r[1])
    return tab_rows, skipped, anchor_updates


# ── Sheet writer ──────────────────────────────────────────────────────────────

def _write_month_tab(service, month_label, tab_rows):
    """
    Create the month tab if needed, then write data rows from row 3.
    If tab already exists: only rows 3+ are cleared — title and header preserved.
    Does NOT touch any other tab.
    """
    meta     = service.spreadsheets().get(spreadsheetId=AM_DATES_SHEET_ID).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}

    if month_label not in existing:
        print(f"  Creating new tab: '{month_label}'")
        service.spreadsheets().batchUpdate(
            spreadsheetId=AM_DATES_SHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": month_label}}}]},
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=AM_DATES_SHEET_ID,
            range=f"'{month_label}'!A1:E2",
            valueInputOption="RAW",
            body={"values": [
                [MONTH_TAB_TITLE_ROW.format(label=month_label)],
                MONTH_TAB_HEADER,
            ]},
        ).execute()
    else:
        print(f"  Tab '{month_label}' exists — clearing data rows (3+) only")
        service.spreadsheets().values().clear(
            spreadsheetId=AM_DATES_SHEET_ID,
            range=f"'{month_label}'!A3:E1000",
        ).execute()

    if not tab_rows:
        print("  Warning: no data rows to write (all clients skipped or churned)")
        return

    service.spreadsheets().values().update(
        spreadsheetId=AM_DATES_SHEET_ID,
        range=f"'{month_label}'!A3",
        valueInputOption="RAW",
        body={"values": tab_rows},
    ).execute()
    print(f"  Wrote {len(tab_rows)} data rows to '{month_label}'")


def _write_back_anchors(service, anchor_updates):
    """Write computed cadence anchors back to master col D (one batchUpdate)."""
    if not anchor_updates:
        return
    data = [
        {
            "range": f"'{MASTER_TAB}'!D{row}",
            "values": [[anchor_str]],
        }
        for row, anchor_str in anchor_updates
    ]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=AM_DATES_SHEET_ID,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()
    print(f"  Updated {len(anchor_updates)} anchor cells in '{MASTER_TAB}' col D")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    dry_run = "--dry-run" in sys.argv
    args    = [a for a in sys.argv[1:] if not a.startswith("--")]

    if not args:
        print("Usage: python generate_month_tab.py \"October 2026\" [--dry-run]")
        sys.exit(1)

    month_label = args[0].strip()

    from datetime import datetime
    try:
        dt = datetime.strptime(month_label, "%B %Y")
    except ValueError:
        try:
            dt = datetime.strptime(month_label, "%b %Y")
        except ValueError:
            print(f"[X] Cannot parse '{month_label}' — use format 'October 2026'")
            sys.exit(1)

    target_year  = dt.year
    target_month = dt.month

    mode = " [DRY RUN — no writes]" if dry_run else ""
    print(f"\ngenerating: {month_label} ({target_year}-{target_month:02d}){mode}")
    print(f"source:     '{MASTER_TAB}' tab in AM sheet")
    print(f"today:      {date.today()}")

    service = _get_sheets_service()
    clients = _read_master(service)
    active  = [c for c in clients if c["status"] != "churned"]
    print(f"master:     {len(clients)} total clients ({len(active)} active)")

    tab_rows, skipped, anchor_updates = _compute_rows(clients, target_year, target_month)

    if skipped:
        print(f"  Skipped:  {', '.join(skipped)}")

    # Print summary table
    print(f"\n{'─'*76}")
    print(f"{'Client':<28} {'Sustainability':14} {'Renewal':12} {'Payment':12} {'Term'}")
    print(f"{'─'*76}")
    for row in tab_rows:
        print(f"{row[0]:<28} {row[1]:14} {row[2]:12} {row[3]:12} {row[4]}")
    print(f"{'─'*76}")
    print(f"Total: {len(tab_rows)} rows for '{month_label}'")

    if dry_run:
        print("\n[DRY RUN] No writes performed.")
        return

    _write_month_tab(service, month_label, tab_rows)
    _write_back_anchors(service, anchor_updates)
    print(f"\nDone — '{month_label}' tab updated, anchors written to master.")


if __name__ == "__main__":
    main()
