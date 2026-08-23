import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

"""
reconcile_counters.py

Counter rule (cap-at-4):
  written = min(true_current_streak, 4)

  true_current_streak = consecutive weeks from END of last-6 data
                        where ANY breach flag was True (0 if none / recovered).

Every counter write is paired with a GHL note so the 'note added'
workflow fires and reads the already-updated counter field.

Commands:
  python agent/main.py --reconcile-counters          dry-run table
  python agent/main.py --apply-counters              write field + add note
  python agent/main.py --catchup-notes               dry-run note list (catch-up)
  python agent/main.py --catchup-notes --apply       post notes (no field re-write)
"""

import json
import time
from pathlib import Path

_HERE = Path(__file__).parent.resolve()
_ROOT = _HERE.parent

if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

import breach_monitor as bm

CLIENTS_PATH   = _ROOT / "config" / "clients.json"
SETTINGS_PATH  = _ROOT / "config" / "settings.json"
SHEET_IDS_PATH = _ROOT / "config" / "sheet_ids.json"


def _collect(clients: list, settings: dict) -> tuple:
    """
    Read all client sheets + GHL counters, compute written values.
    Returns (results, errors).
    """
    threeup_api_key = settings.get("threeup_api_key", "")
    sheet_ids       = json.loads(SHEET_IDS_PATH.read_text(encoding="utf-8"))
    non_churned     = [c for c in clients if not c.get("churned")]

    svc     = bm._get_sheets_service(readonly=True)
    results = []
    errors  = []

    for c in non_churned:
        name       = c.get("name", "")
        sid        = sheet_ids.get(name)
        contact_id = c.get("contact_id", "")

        if not sid:
            errors.append((name, "no sheet_id"))
            continue

        rows = bm._read_client_rows(svc, sid)
        if rows is None:
            errors.append((name, "Sheets API error"))
            continue

        computed    = [bm._compute_flags(r) for r in rows]
        data_rows   = [r for r in computed if r["has_data"]]
        last6       = data_rows[-6:]
        true_streak = bm._overall_current_streak(last6)
        written     = min(true_streak, 4)

        ghl_note    = ""
        ghl_current = None
        if not contact_id:
            ghl_note = "no contact_id"
        elif contact_id in bm._KNOWN_400_CONTACT_IDS:
            ghl_note = "known bad contact"
        elif threeup_api_key:
            ghl_current = bm._ghl_get_counter(threeup_api_key, contact_id)
            if ghl_current is None:
                ghl_note = "GHL read error"

        can_write = bool(contact_id) and contact_id not in bm._KNOWN_400_CONTACT_IDS

        results.append({
            "name":        name,
            "contact_id":  contact_id,
            "frozen":      ghl_current,
            "ghl_note":    ghl_note,
            "true_streak": true_streak,
            "written":     written,
            "can_write":   can_write,
        })
        time.sleep(0.10)

    return results, errors


def run(clients: list, settings: dict) -> None:
    """Dry-run: show table, write nothing."""
    print(f"[reconcile] collecting data...")
    results, errors = _collect(clients, settings)

    W = 116
    print("\n" + "═" * W)
    print("  GHL COUNTER — FINAL DRY RUN  (cap-at-4, no writes)")
    print("  written = min(true_streak, 4)  ·  0 = recovery  ·  each write paired with a note")
    print("═" * W)
    print(f"  {'Client':<42} {'Frozen':>7} {'Streak':>7} {'Write':>6}  Automation + note")
    print(f"  {'─'*42} {'─'*7} {'─'*7} {'─'*6}  {'─'*50}")

    at_4=[]; at_3=[]; at_2=[]; at_1=[]; at_0=[]; cant=[]
    for r in sorted(results, key=lambda x: x["name"]):
        frozen_s = str(r["frozen"]) if r["frozen"] is not None else "N/A"
        note_txt = bm._breach_note_body(r["written"], r["true_streak"])
        auto     = bm._ghl_auto_label(r["written"])
        delta    = ""
        if r["frozen"] is not None:
            d = r["written"] - r["frozen"]
            delta = f" ({'+' if d >= 0 else ''}{d})"
        extra = f"  ← {r['ghl_note']}" if r["ghl_note"] else ""
        print(f"  {r['name']:<42} {frozen_s:>7} {r['true_streak']:>7} {r['written']:>6}  {auto}{delta}{extra}")
        print(f"  {'':42}  note: \"{note_txt}\"")

        if not r["can_write"]:        cant.append(r["name"])
        elif r["written"] == 4:       at_4.append(r["name"])
        elif r["written"] == 3:       at_3.append(r["name"])
        elif r["written"] == 2:       at_2.append(r["name"])
        elif r["written"] == 1:       at_1.append(r["name"])
        else:                         at_0.append(r["name"])

    print()
    if at_4:
        print(f"  ★★ → 4 change-strategy ({len(at_4)}): " + ", ".join(at_4))
    if at_3:
        print(f"  ★  → 3 note created    ({len(at_3)}): " + ", ".join(at_3))
    if at_2:
        print(f"  ·  → 2 noted           ({len(at_2)}): " + ", ".join(at_2))
    if at_1:
        print(f"  ·  → 1 first bad week  ({len(at_1)}): " + ", ".join(at_1))
    if at_0:
        print(f"  ✓  → 0 recovery        ({len(at_0)}): " + ", ".join(at_0))
    if cant:
        print(f"  ⚠  skipped             ({len(cant)}): " + ", ".join(cant))
    print(f"\n  {len(results)-len(cant)} writable  ·  nothing written")
    if errors:
        print("\n  ⚠ Sheet errors: " + "; ".join(f"{n}: {r}" for n,r in errors))
    print("═" * W)


def run_apply(clients: list, settings: dict) -> None:
    """Apply: write field + post note for each client."""
    threeup_api_key = settings.get("threeup_api_key", "")
    print(f"[apply] collecting data...")
    results, errors = _collect(clients, settings)

    writable = [r for r in results if r["can_write"]]
    print(f"[apply] writing {len(writable)} GHL counters...")
    for r in results:
        if not r["can_write"]:
            r["write_result"] = "SKIPPED"
            continue
        ok_field = bm._ghl_write_counter(threeup_api_key, r["contact_id"], r["written"])
        if ok_field:
            r["write_result"] = "OK (field)"
        else:
            r["write_result"] = "FAIL (field)"
        print(f"  {r['name']:<44} → {r['written']}  {r['write_result']}")
        time.sleep(0.12)

    W = 116
    print("\n" + "═" * W)
    print("  GHL COUNTER APPLY — RESULTS")
    print("═" * W)
    print(f"  {'Client':<42} {'Frozen':>7} {'Streak':>7} {'Written':>8}  Status")
    print(f"  {'─'*42} {'─'*7} {'─'*7} {'─'*8}  {'─'*36}")
    for r in sorted(results, key=lambda x: x["name"]):
        frozen_s = str(r["frozen"]) if r["frozen"] is not None else "N/A"
        print(f"  {r['name']:<42} {frozen_s:>7} {r['true_streak']:>7} {r['written']:>8}  {r['write_result']}")
    ok   = sum(1 for r in results if "OK" in r.get("write_result",""))
    fail = sum(1 for r in results if "FAIL" in r.get("write_result",""))
    skip = sum(1 for r in results if r.get("write_result") == "SKIPPED")
    print(f"\n  {ok} OK  ·  {fail} FAILED  ·  {skip} skipped")
    if errors:
        print("  ⚠ Sheet errors: " + "; ".join(f"{n}: {r}" for n,r in errors))
    print("═" * W)


def run_catchup_notes(clients: list, settings: dict, dry_run: bool = True) -> None:
    """
    One-time catch-up: add notes for contacts whose field was already
    written today WITHOUT a note.  Does NOT re-write the field.
    """
    threeup_api_key = settings.get("threeup_api_key", "")
    tag = "DRY RUN — " if dry_run else ""
    print(f"[catchup] {tag}collecting data...")
    results, errors = _collect(clients, settings)

    writable = [r for r in results if r["can_write"]]
    W = 110
    print("\n" + "═" * W)
    print(f"  NOTE CATCH-UP — {tag}posting notes only (field already correct)")
    print(f"  {len(writable)} contacts  ·  {'nothing posted yet' if dry_run else 'LIVE POSTING'}")
    print("═" * W)

    for r in sorted(writable, key=lambda x: x["name"]):
        print(f"  {r['name']:<44}  counter={r['written']}")

    if r_cant := [r for r in results if not r["can_write"]]:
        for r in r_cant:
            print(f"  {r['name']:<44}  SKIPPED  ← {r['ghl_note']}")

    print()
    print(f"  {len(writable)} contact(s) listed — GHL notes are no longer posted by this tool.")
    if errors:
        print("  ⚠ Sheet errors: " + "; ".join(f"{n}: {e}" for n,e in errors))
    print("═" * W)


def main():
    try:
        clients  = json.loads(CLIENTS_PATH.read_text(encoding="utf-8"))
        settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"ERROR loading config: {e}")
        sys.exit(1)

    if "--apply" in sys.argv:
        run_apply(clients, settings)
    else:
        run(clients, settings)


if __name__ == "__main__":
    main()
