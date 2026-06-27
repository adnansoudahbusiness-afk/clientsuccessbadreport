import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

"""
main.py — Orchestrator entry point for ThreeUp CAO Agent.
Run: python agent/main.py
"""

import json
import logging
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import pytz
from dotenv import load_dotenv

# ── Path setup ────────────────────────────────────────────────────────────────
AGENT_DIR = Path(__file__).parent.resolve()
BASE_DIR  = AGENT_DIR.parent
LOGS_DIR  = BASE_DIR / "logs"

if str(AGENT_DIR) not in sys.path:
    sys.path.insert(0, str(AGENT_DIR))

# ── Env & logging ─────────────────────────────────────────────────────────────
load_dotenv(BASE_DIR / ".env")

LOGS_DIR.mkdir(parents=True, exist_ok=True)
AMMAN_TZ  = pytz.timezone("Asia/Amman")
today_str = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d")
log_path  = LOGS_DIR / f"run_{today_str}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[
        logging.FileHandler(str(log_path), encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("main")

# ── Module imports ─────────────────────────────────────────────────────────────
import good_report_engine
from notifier_email import send_email
from ghl_client import (
    get_new_patients,
    get_returning_patients_tagged,
    get_trigger_link_clicks,
    get_appointments,
    get_loyalty_points,
    get_active_workflow,
    get_active_snippet,
    get_conversations_sample,
)
from sheets_manager import append_weekly_row, clear_sheet_data, clear_form_submitted, force_update_all_headers, check_date_range_exists, backfill_total_reviews, write_audit_cell
from watcher import Watcher
from form_submitter import scrape_form_fields

# ── Settings ──────────────────────────────────────────────────────────────────
_settings_path = os.path.join(os.path.dirname(__file__), '..', 'config', 'settings.json')
with open(_settings_path, encoding='utf-8') as _sf:
    _settings = json.load(_sf)

# ── Constants ─────────────────────────────────────────────────────────────────
PLACEHOLDER_API_KEY     = "PASTE_YOUR_SUBACCOUNT_API_KEY_HERE"
PLACEHOLDER_LOCATION_ID = "PASTE_YOUR_LOCATION_ID_HERE"

# ── Hardcoded historical weeks (Asia/Amman) ───────────────────────────────────
# Each entry: (year, month, day, hour, minute, second)
ALL_WEEKS = [
    {
        "label":      "Week 1",
        "this_start": (2026,  3, 21,  0,  0,  0),
        "this_end":   (2026,  3, 26, 23, 59, 59),
        "last_start": (2026,  3, 14,  0,  0,  0),
        "last_end":   (2026,  3, 19, 23, 59, 59),
    },
    {
        "label":      "Week 2",
        "this_start": (2026,  3, 28,  0,  0,  0),
        "this_end":   (2026,  4,  2, 23, 59, 59),
        "last_start": (2026,  3, 21,  0,  0,  0),
        "last_end":   (2026,  3, 26, 23, 59, 59),
    },
    {
        "label":      "Week 3",
        "this_start": (2026,  4,  4,  0,  0,  0),
        "this_end":   (2026,  4,  9, 23, 59, 59),
        "last_start": (2026,  3, 28,  0,  0,  0),
        "last_end":   (2026,  4,  2, 23, 59, 59),
    },
    {
        "label":      "Week 4",
        "this_start": (2026,  4, 11,  0,  0,  0),
        "this_end":   (2026,  4, 16, 23, 59, 59),
        "last_start": (2026,  4,  4,  0,  0,  0),
        "last_end":   (2026,  4,  9, 23, 59, 59),
    },
    {
        "label":      "Week 5",
        "this_start": (2026,  4, 18,  0,  0,  0),
        "this_end":   (2026,  4, 24, 23, 59, 59),
        "last_start": (2026,  4, 11,  0,  0,  0),
        "last_end":   (2026,  4, 16, 23, 59, 59),
    },
    {
        "label":      "Week 6",
        "this_start": (2026,  4, 25,  0,  0,  0),
        "this_end":   (2026,  4, 30, 23, 59, 59),
        "last_start": (2026,  4, 18,  0,  0,  0),
        "last_end":   (2026,  4, 23, 23, 59, 59),
    },
    {
        "label":      "Week 7",
        "this_start": (2026,  5,  2,  0,  0,  0),
        "this_end":   (2026,  5,  8, 23, 59, 59),
        "last_start": (2026,  4, 25,  0,  0,  0),
        "last_end":   (2026,  5,  1, 23, 59, 59),
    },
    {
        "label":      "Week 8",
        "this_start": (2026,  5,  9,  0,  0,  0),
        "this_end":   (2026,  5, 15, 23, 59, 59),
        "last_start": (2026,  5,  2,  0,  0,  0),
        "last_end":   (2026,  5,  8, 23, 59, 59),
    },
    {
        "label":      "Week 9",
        "this_start": (2026,  5, 16,  0,  0,  0),
        "this_end":   (2026,  5, 22, 23, 59, 59),
        "last_start": (2026,  5,  9,  0,  0,  0),
        "last_end":   (2026,  5, 15, 23, 59, 59),
    },
    {
        "label":      "Week 10",
        "this_start": (2026,  5, 23,  0,  0,  0),
        "this_end":   (2026,  5, 29, 23, 59, 59),
        "last_start": (2026,  5, 16,  0,  0,  0),
        "last_end":   (2026,  5, 22, 23, 59, 59),
    },
    {
        "label":      "Week 11",
        "this_start": (2026,  5, 30,  0,  0,  0),
        "this_end":   (2026,  6,  5, 23, 59, 59),
        "last_start": (2026,  5, 23,  0,  0,  0),
        "last_end":   (2026,  5, 29, 23, 59, 59),
    },
    {
        "label":      "Week 12",
        "this_start": (2026,  6,  6,  0,  0,  0),
        "this_end":   (2026,  6, 12, 23, 59, 59),
        "last_start": (2026,  5, 30,  0,  0,  0),
        "last_end":   (2026,  6,  5, 23, 59, 59),
    },
    {
        "label":      "Week 13",
        "this_start": (2026,  6, 13,  0,  0,  0),
        "this_end":   (2026,  6, 19, 23, 59, 59),
        "last_start": (2026,  6,  6,  0,  0,  0),
        "last_end":   (2026,  6, 12, 23, 59, 59),
    },
    {
        "label":      "Week 14",
        "this_start": (2026,  6, 20,  0,  0,  0),
        "this_end":   (2026,  6, 26, 23, 59, 59),
        "last_start": (2026,  6, 13,  0,  0,  0),
        "last_end":   (2026,  6, 19, 23, 59, 59),
    },
]


def _localize_week(week_def: dict) -> tuple:
    """Convert ALL_WEEKS entry to 4 timezone-aware datetimes."""
    tz = pytz.timezone("Asia/Amman")
    return (
        tz.localize(datetime(*week_def["this_start"])),
        tz.localize(datetime(*week_def["this_end"])),
        tz.localize(datetime(*week_def["last_start"])),
        tz.localize(datetime(*week_def["last_end"])),
    )


# ── Date helpers ──────────────────────────────────────────────────────────────

def get_week_bounds(reference_date: datetime):
    """
    Return (this_start, this_end, last_start, last_end) in Amman TZ.
    Reporting week: Saturday 00:00:00 → Friday 23:59:59.
    """
    tz = pytz.timezone("Asia/Amman")
    if reference_date.tzinfo is None:
        reference_date = tz.localize(reference_date)
    else:
        reference_date = reference_date.astimezone(tz)

    today = reference_date.replace(hour=0, minute=0, second=0, microsecond=0)
    wd = today.weekday()  # Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6

    if wd == 4:    # Friday: week ends today
        this_end   = today.replace(hour=23, minute=59, second=59, microsecond=0)
        this_start = today - timedelta(days=6)          # last Saturday
        last_start = this_start - timedelta(days=7)
        last_end   = (this_start - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)
    elif wd == 5:  # Saturday: week just started
        this_start = today
        this_end   = (today + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)
        last_start = today - timedelta(days=7)
        last_end   = (today - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)
    else:          # Sun–Thu: most recent Saturday
        days_back  = (wd - 5) % 7
        this_start = today - timedelta(days=days_back)
        this_end   = (this_start + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)
        last_start = this_start - timedelta(days=7)
        last_end   = (this_start - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)

    return this_start, this_end, last_start, last_end


def _fmt_date_range(start: datetime, end: datetime) -> str:
    day_names = {5: "Sat", 6: "Sun", 0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri"}
    s_day = day_names.get(start.weekday(), "")
    e_day = day_names.get(end.weekday(), "")
    return f"{s_day} {start.day}/{start.month} – {e_day} {end.day}/{end.month}"


def _check_seasonal_flag(start: datetime) -> str:
    """Flag Ramadan overlap heuristically (Feb/Mar months)."""
    config_path = BASE_DIR / "config" / "thresholds.json"
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            thresholds = json.load(f)
        for period in thresholds.get("seasonal_periods", []):
            if period.get("name") == "Ramadan" and start.month in (2, 3):
                return "Check-Ramadan"
        return "FALSE"
    except Exception:
        return "FALSE"


def _infer_strategy_from_snippet(snippet_name: str, snippet_value: str) -> str:
    """Infer strategy number from snippet name/value patterns."""
    if snippet_name is None and snippet_value is None:
        return "Unknown"
    if snippet_name is None:
        snippet_name = ""
    if snippet_value is None:
        snippet_value = ""
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


# ── Banner ────────────────────────────────────────────────────────────────────

def print_banner(now: datetime):
    print("=" * 60)
    print("  ThreeUp CAO Agent")
    print(f"  {now.strftime('%A %d/%m/%Y %H:%M')} (Amman)")
    print("=" * 60)


def _is_placeholder_client(client: dict) -> bool:
    return (
        client.get("ghl_api_key", "")  == PLACEHOLDER_API_KEY or
        client.get("location_id", "") == PLACEHOLDER_LOCATION_ID
    )


# ── Data pull ─────────────────────────────────────────────────────────────────

def pull_client_data(client: dict,
                     this_start: datetime, this_end: datetime,
                     last_start: datetime, last_end: datetime) -> tuple:
    """
    Pull all GHL metrics for one client.
    Returns (data_dict, messages_list).
    """
    name        = client["name"]
    api_key     = client["ghl_api_key"]
    location_id = client["location_id"]
    _np_tag  = client.get("new_patient_tag") or None
    _tl_tag  = client.get("trigger_link_tag") or None
    new_patients_this  = get_new_patients(api_key, location_id, this_start, this_end, name, new_patient_tag=_np_tag)
    new_patients_last  = get_new_patients(api_key, location_id, last_start, last_end, name, new_patient_tag=_np_tag)
    returning_tagged   = get_returning_patients_tagged(api_key, location_id, this_start, this_end, name)
    trigger_clicks     = get_trigger_link_clicks(api_key, location_id, this_start, this_end, name, trigger_link_tag=_tl_tag)

    appts_this = get_appointments(api_key, location_id, this_start, this_end, name)
    logger.info(f"[{name}] Last week range: {last_start} → {last_end}")
    logger.info(f"[{name}] Last week timestamps: {int(last_start.timestamp()*1000)} → {int(last_end.timestamp()*1000)}")
    appts_last = get_appointments(api_key, location_id, last_start, last_end, name)

    loyalty_this = get_loyalty_points(api_key, location_id, this_start, this_end, name)
    loyalty_last = get_loyalty_points(api_key, location_id, last_start, last_end, name)

    workflow_name = get_active_workflow(api_key, location_id, name)

    snippet = None
    if workflow_name:
        snippet = get_active_snippet(api_key, location_id, workflow_name, name)

    snippet_name     = snippet["name"]  if snippet else ""
    snippet_value    = snippet["value"] if snippet else ""
    current_strategy = _infer_strategy_from_snippet(snippet_name, snippet_value)

    messages_list = get_conversations_sample(api_key, location_id, name, limit=30)

    appts_this_confirmed = appts_this.get("confirmed", appts_this.get("total", 0))
    appts_last_confirmed = appts_last.get("confirmed", appts_last.get("total", 0))

    new_patients_growth = (
        round((new_patients_this - new_patients_last) / new_patients_last * 100, 2)
        if new_patients_last > 0 else 0
    )
    appointments_growth = (
        round((appts_this_confirmed - appts_last_confirmed) / appts_last_confirmed * 100, 2)
        if appts_last_confirmed > 0 else 0
    )
    loyalty_change = (
        round((loyalty_this - loyalty_last) / loyalty_last * 100, 2)
        if loyalty_last > 0 else 0
    )

    data_dict = {
        "date_range":                       _fmt_date_range(this_start, this_end),
        "new_patients_this_week":            new_patients_this,
        "new_patients_last_week":            new_patients_last,
        "new_patients_growth":               new_patients_growth,
        "returning_patients":                returning_tagged,
        "trigger_links":                     trigger_clicks,
        "appointments_booked":               appts_this.get("booked", 0),
        "appointments_confirmed":            appts_this_confirmed,
        "appointments_cancelled":            appts_this.get("cancelled", 0),
        "cancellation_rate":                 appts_this.get("cancellation_rate", 0),
        "appointments_confirmed_last_week":  appts_last_confirmed,
        "appointments_growth":               appointments_growth,
        "loyalty_this":                      loyalty_this,
        "loyalty_prior":                     loyalty_last,
        "loyalty_change":                    loyalty_change,
        "workflow_name":                     workflow_name,
        "snippet_name":                      snippet_name,
        "current_strategy":                  current_strategy,
    }

    return data_dict, messages_list


# ── Shared helpers ────────────────────────────────────────────────────────────

def _load_clients() -> tuple:
    """Load, validate and return (all_clients, valid_clients). Churned clients are excluded from both."""
    clients_path = BASE_DIR / "config" / "clients.json"
    if not clients_path.exists():
        print(f"[X] config/clients.json not found at {clients_path}")
        sys.exit(1)

    with open(clients_path, "r", encoding="utf-8") as f:
        all_raw = json.load(f)

    if not all_raw:
        print("[X] clients.json is empty.")
        sys.exit(1)

    clients = []
    for c in all_raw:
        if c.get("churned"):
            name = c.get("name", "<unnamed>")
            print(f"[!] Skipping '{name}' — churned.")
            logger.warning(f"[main] Skipping '{name}': churned")
        else:
            clients.append(c)

    valid_clients = []
    for client in clients:
        name = client.get("name", "<unnamed>")
        err  = _validate_client(client)
        if err:
            print(f"[!] Skipping '{name}' — {err}")
            logger.warning(f"[main] Skipping '{name}': {err}")
        elif _is_placeholder_client(client):
            print(f"[!] Skipping '{name}' — placeholder credentials.")
            logger.warning(f"[main] Skipping '{name}': placeholder credentials")
        else:
            valid_clients.append(client)

    if not valid_clients:
        print("[X] No valid clients to process.")
        sys.exit(1)

    return clients, valid_clients


def _pull_and_append_week(
    valid_clients: list,
    this_start: datetime, this_end: datetime,
    last_start: datetime, last_end: datetime,
    ghl_data_cache: dict,
    skip_dup: bool = True,
) -> tuple:
    """
    Pull data for all valid_clients for one week and append rows.
    skip_dup=True  → skip clients whose date_range row already exists (safe re-run).
    skip_dup=False → always append (use for targeted --week reruns).
    Returns (succeeded_list, failed_list).
    """
    date_range_label = _fmt_date_range(this_start, this_end)
    succeeded: list  = []
    failed:    list  = []
    total            = len(valid_clients)

    for i, client in enumerate(valid_clients, 1):
        name = client["name"]

        if skip_dup:
            try:
                if check_date_range_exists(name, date_range_label):
                    print(f"[{i}/{total}] {name} — row for '{date_range_label}' already exists, skipping.")
                    logger.info(f"[main] [{name}] Duplicate row skipped: {date_range_label}")
                    succeeded.append(name)  # treat as success for watcher cache
                    continue
            except Exception as dup_err:
                logger.warning(f"[main] [{name}] check_date_range_exists failed: {dup_err}")

        print(f"[{i}/{total}] Pulling data for {name}...")
        try:
            data_dict, messages_list = pull_client_data(
                client, this_start, this_end, last_start, last_end
            )
            append_weekly_row(name, data_dict)
            ghl_data_cache[name] = {"messages_list": messages_list}

            ts = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M:%S")
            print(f"[OK] {name} — sheet updated")
            logger.info(f"[{ts}] [{name}] [main] Data pulled and written successfully")
            succeeded.append(name)

        except Exception as e:
            ts = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M:%S")
            print(f"[X] {name} — FAILED: {e}")
            logger.error(f"[{ts}] [{name}] [main] Pull failed: {e}", exc_info=True)
            failed.append(f"{name} ({e})")

        time.sleep(2)

    print("\n" + "=" * 44)
    print(f"  Week: {date_range_label}")
    print(f"  Succeeded: {len(succeeded)}  |  Failed: {len(failed)}")
    if failed:
        print(f"  Failed: {'; '.join(failed)}")
    print("=" * 44)

    return succeeded, failed


def _start_and_watch(clients: list, ghl_data_cache: dict):
    """Print action instructions, start watcher, and block until done."""
    print("\n" + "=" * 44)
    print("   ALL DATA PULLED — ACTION REQUIRED")
    print("=" * 44)
    print("For each client Google Sheet:")
    print("  1. Fill column B (Google Reviews — yellow cell)")
    print("  2. Set column C to YES when ready")
    print("Watcher running every 2 minutes. Forms auto-submit on confirmation.")
    print("=" * 44 + "\n")

    watcher = Watcher(clients=clients, ghl_data_cache=ghl_data_cache)
    watcher.start()

    ts = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M:%S")
    logger.info(f"[{ts}] [main] Watcher thread started")

    try:
        while True:
            time.sleep(60)
            if not watcher.is_alive():
                print(f"\n[{datetime.now(AMMAN_TZ).strftime('%Y-%m-%d %H:%M:%S')}] Watcher finished. Exiting.")
                break
    except KeyboardInterrupt:
        print(f"\n[{datetime.now(AMMAN_TZ).strftime('%Y-%m-%d %H:%M:%S')}] Interrupted — shutting down.")
        watcher.stop()


# ── Main ──────────────────────────────────────────────────────────────────────

def _validate_client(client: dict) -> str:
    """Return error string if client entry is invalid, else empty string."""
    name = client.get("name", "")
    if not name or not isinstance(name, str):
        return "missing or empty 'name'"
    key = client.get("ghl_api_key", "")
    if not key or not str(key).startswith("pit-"):
        return "ghl_api_key missing or does not start with 'pit-'"
    if not client.get("location_id", ""):
        return "missing 'location_id'"
    if not client.get("phone", ""):
        return "missing 'phone'"
    return ""


def _rotate_logs(logs_dir: Path, keep: int = 10):
    """Delete oldest log files, keeping only the most recent `keep` files."""
    log_files = sorted(logs_dir.glob("run_*.log"), key=lambda p: p.stat().st_mtime)
    for old in log_files[:-keep]:
        try:
            old.unlink()
            logger.info(f"[log rotation] Deleted old log: {old.name}")
        except Exception:
            pass


def main():
    """Dynamic mode — auto-calculates current week (Sat–Fri) and runs once."""
    now = datetime.now(AMMAN_TZ)
    print_banner(now)
    _rotate_logs(LOGS_DIR, keep=10)

    # Dynamic week bounds: last Saturday 00:00 → last Friday 23:59:59
    this_start, this_end, last_start, last_end = get_week_bounds(now)
    print(f"  This week : {_fmt_date_range(this_start, this_end)}")
    print(f"  Last week : {_fmt_date_range(last_start, last_end)}")
    print()

    clients, valid_clients = _load_clients()
    ghl_data_cache = {}
    _pull_and_append_week(valid_clients, this_start, this_end, last_start, last_end,
                          ghl_data_cache, skip_dup=True)
    send_email(
        "ThreeUp CAO — Weekly Run Complete ✅",
        f"Weekly data pull finished at {datetime.now().strftime('%d/%m/%Y %H:%M')}.\n"
        f"{len(valid_clients)} clients processed.\n"
        f"Good report engine running next.\n"
        f"Waiting for Laith to confirm reviews on dashboard."
    )
    amman_hour = datetime.now(pytz.timezone("Asia/Amman")).hour
    if 9 <= amman_hour <= 17:
        good_report_engine.run(valid_clients, _settings)
    else:
        print(f"[good_report] Skipping — outside business hours ({amman_hour}:00 Amman time)")
    _start_and_watch(clients, ghl_data_cache)


def all_weeks_mode():
    """Run all 4 historical weeks sequentially, then start watcher once."""
    now = datetime.now(AMMAN_TZ)
    print_banner(now)
    _rotate_logs(LOGS_DIR, keep=10)

    clients, valid_clients = _load_clients()
    ghl_data_cache = {}

    for week_def in ALL_WEEKS:
        this_start, this_end, last_start, last_end = _localize_week(week_def)
        print(f"\n{'=' * 56}")
        print(f"  {week_def['label']}: {_fmt_date_range(this_start, this_end)}")
        print(f"  Last week: {_fmt_date_range(last_start, last_end)}")
        print(f"{'=' * 56}")
        _pull_and_append_week(valid_clients, this_start, this_end, last_start, last_end,
                              ghl_data_cache, skip_dup=True)

    print(f"\n{'=' * 56}")
    print(f"  All {len(ALL_WEEKS)} weeks pulled — starting watcher")
    print(f"{'=' * 56}")
    _start_and_watch(clients, ghl_data_cache)


def week_n_mode(n: int):
    """Run a single hardcoded week by number (1–4), then start watcher."""
    if n < 1 or n > len(ALL_WEEKS):
        print(f"[X] --week must be 1–{len(ALL_WEEKS)}")
        sys.exit(1)

    now = datetime.now(AMMAN_TZ)
    print_banner(now)
    _rotate_logs(LOGS_DIR, keep=10)

    clients, valid_clients = _load_clients()
    week_def = ALL_WEEKS[n - 1]
    this_start, this_end, last_start, last_end = _localize_week(week_def)

    print(f"  {week_def['label']}: {_fmt_date_range(this_start, this_end)}")
    print(f"  Last week : {_fmt_date_range(last_start, last_end)}")
    print()

    ghl_data_cache = {}
    # skip_dup=False so --week reruns always re-pull even if row exists
    _pull_and_append_week(valid_clients, this_start, this_end, last_start, last_end,
                          ghl_data_cache, skip_dup=False)
    _start_and_watch(clients, ghl_data_cache)


# ── clear_all_sheets ──────────────────────────────────────────────────────────

def clear_all_sheets():
    """Load all clients and wipe every sheet's data rows, leaving headers."""
    clients_path = BASE_DIR / "config" / "clients.json"
    if not clients_path.exists():
        print(f"[X] config/clients.json not found at {clients_path}")
        sys.exit(1)

    with open(clients_path, "r", encoding="utf-8") as f:
        clients = json.load(f)

    if not clients:
        print("[X] clients.json is empty.")
        sys.exit(1)

    clients = [c for c in clients if not c.get("churned")]
    print(f"\nClearing {len(clients)} sheet(s)...\n")
    for client in clients:
        name = client.get("name", "<unnamed>")
        try:
            clear_sheet_data(name)
            print(f"  {name} — Sheet cleared")
        except Exception as e:
            print(f"  {name} — clear_sheet_data FAILED: {e}")
            logger.error(f"[main] clear_all_sheets failed for '{name}': {e}")
        try:
            clear_form_submitted(name)
            print(f"  {name} — AC/AD cleared")
        except Exception as e:
            print(f"  {name} — clear_form_submitted FAILED: {e}")
            logger.error(f"[main] clear_form_submitted failed for '{name}': {e}")

    print("\nDone.")


# ── watch_only ────────────────────────────────────────────────────────────────

def watch_only():
    """Skip data pull — start the watcher immediately on all clients."""
    clients_path = BASE_DIR / "config" / "clients.json"
    if not clients_path.exists():
        print(f"[X] config/clients.json not found at {clients_path}")
        sys.exit(1)

    with open(clients_path, "r", encoding="utf-8") as f:
        clients = json.load(f)

    if not clients:
        print("[X] clients.json is empty.")
        sys.exit(1)

    clients = [c for c in clients if not c.get("churned")]
    print(f"Watcher started — monitoring {len(clients)} clients")
    ts = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M:%S")
    logger.info(f"[{ts}] [main] --watch-only mode — skipping data pull")

    watcher = Watcher(clients=clients, ghl_data_cache={})
    watcher.start()

    try:
        while True:
            time.sleep(60)
            if not watcher.is_alive():
                print(f"\n[{datetime.now(AMMAN_TZ).strftime('%Y-%m-%d %H:%M:%S')}] Watcher finished. Exiting.")
                break
    except KeyboardInterrupt:
        print(f"\n[{datetime.now(AMMAN_TZ).strftime('%Y-%m-%d %H:%M:%S')}] Interrupted by user — shutting down.")
        watcher.stop()


# ── audit_mode ────────────────────────────────────────────────────────────────

def audit_mode(fix: bool = False, week_index: int = None):
    """
    --audit [--week N] : Compare sheet values against fresh GHL data.
    --audit --fix      : Same, plus overwrite sheet cells where discrepancies exist.

    week_index: 1-based week number (matches ALL_WEEKS order).
                Defaults to ALL_WEEKS[-2] (most recent week with pulled data).
    Checks: New Patients, Trigger Links, Confirmed Appts, Cancel Rate.
    Also validates formula correctness: hook_pct, offer_pct.
    """
    from sheets_manager import get_all_sheet_rows

    clients_path = BASE_DIR / "config" / "clients.json"
    if not clients_path.exists():
        print(f"[X] config/clients.json not found at {clients_path}")
        sys.exit(1)
    with open(clients_path, "r", encoding="utf-8") as f:
        clients = json.load(f)

    # Resolve which week to audit
    if week_index is not None:
        if week_index < 1 or week_index > len(ALL_WEEKS):
            print(f"[X] --week must be 1–{len(ALL_WEEKS)} (got {week_index})")
            sys.exit(1)
        week_def = ALL_WEEKS[week_index - 1]
    else:
        week_def = ALL_WEEKS[-1]   # default: most recent week

    is_historical = (week_def is not ALL_WEEKS[-1])
    this_start, this_end, _, _ = _localize_week(week_def)
    print(f"\n{'=' * 60}")
    print(f"  AUDIT — {week_def['label']}: {_fmt_date_range(this_start, this_end)}")
    if fix:
        print("  MODE: --fix enabled — discrepancies will be corrected")
    print(f"{'=' * 60}\n")
    if is_historical:
        print(
            f"  \u26A0\uFE0F  WARNING: Auditing historical week {week_def['label']}.\n"
            f"  GHL data has changed since this week ran.\n"
            f"  Discrepancies may be false positives.\n"
            f"  Only audit the current week for accurate results.\n"
        )

    clean_clients      = []
    discrepant_clients = []

    # Column indices (0-based) matching sheet layout
    COL_NEW_PATIENTS  = 3   # D
    COL_TRIGGER_LINKS = 7   # H
    COL_HOOK_PCT      = 8   # I
    COL_OFFER_PCT     = 9   # J
    COL_CONF_APPTS    = 13  # N
    COL_CANCEL_RATE   = 15  # P

    for client in clients:
        name        = client.get("name", "<unnamed>")
        api_key     = client.get("ghl_api_key", "")
        location_id = client.get("location_id", "")

        if client.get("churned"):
            print(f"  ⏭  {name} — skipped (churned)")
            continue

        # Skip placeholder clients
        if api_key in ("PASTE_YOUR_SUBACCOUNT_API_KEY_HERE", "") or \
           location_id in ("PASTE_YOUR_LOCATION_ID_HERE", ""):
            print(f"  ⏭  {name} — skipped (not configured)")
            continue

        print(f"  Auditing {name}...")

        # ── Pull fresh GHL values ──────────────────────────────────────────────
        try:
            ghl_new_patients = get_new_patients(
                api_key, location_id, this_start, this_end, name,
                new_patient_tag=client.get("new_patient_tag") or None)
            ghl_trigger_links = get_trigger_link_clicks(
                api_key, location_id, this_start, this_end, name,
                trigger_link_tag=client.get("trigger_link_tag") or None)
            appts = get_appointments(
                api_key, location_id, this_start, this_end, name)
            ghl_confirmed  = appts.get("confirmed", 0)
            ghl_booked     = appts.get("booked", 0)
            ghl_cancelled  = appts.get("cancelled", 0)
            ghl_cancel_rate = appts.get("cancellation_rate", 0.0)
        except Exception as e:
            print(f"  ❌ {name} — GHL fetch failed: {e}\n")
            discrepant_clients.append(name)
            continue

        # ── Read most recent sheet row ─────────────────────────────────────────
        try:
            all_rows = get_all_sheet_rows(name)
            if not all_rows:
                print(f"  ⚠️  {name} — no sheet rows found\n")
                discrepant_clients.append(name)
                continue
            last_row = all_rows[-1]
            row_num  = len(all_rows) + 1   # +1 for header row → 1-based sheet row

            def _sheet_num(col):
                try:
                    v = last_row[col] if len(last_row) > col else ""
                    return float(v) if v not in ("", None) else 0.0
                except (ValueError, TypeError):
                    return 0.0

            sheet_new_patients  = _sheet_num(COL_NEW_PATIENTS)
            sheet_trigger_links = _sheet_num(COL_TRIGGER_LINKS)
            sheet_confirmed     = _sheet_num(COL_CONF_APPTS)
            sheet_cancel_rate   = _sheet_num(COL_CANCEL_RATE)
            sheet_hook_pct      = _sheet_num(COL_HOOK_PCT)
            sheet_offer_pct     = _sheet_num(COL_OFFER_PCT)
        except Exception as e:
            print(f"  ❌ {name} — sheet read failed: {e}\n")
            discrepant_clients.append(name)
            continue

        # ── Recalculate expected formula values ────────────────────────────────
        calc_hook_pct  = round(sheet_trigger_links / sheet_new_patients, 4) \
                         if sheet_new_patients > 0 else 0.0
        calc_offer_pct = round(sheet_new_patients and
                               _sheet_num(COL_NEW_PATIENTS) or 0, 4)   # placeholder; use reviews below
        # For offer_pct: reviews = col 1 (B), trigger = col 7 (H)
        sheet_reviews  = _sheet_num(1)
        calc_offer_pct = round(sheet_reviews / sheet_trigger_links, 4) \
                         if sheet_trigger_links > 0 else 0.0
        calc_cancel_rt = round(ghl_cancelled / ghl_booked, 4) \
                         if ghl_booked > 0 else 0.0

        # ── Compare and report ─────────────────────────────────────────────────
        checks = [
            # (label, col_idx, sheet_val, ghl_val, is_float_compare)
            ("New Patients",     COL_NEW_PATIENTS,  sheet_new_patients,  float(ghl_new_patients),  False),
            ("Trigger Links",    COL_TRIGGER_LINKS, sheet_trigger_links, float(ghl_trigger_links),  False),
            ("Confirmed Appts",  COL_CONF_APPTS,    sheet_confirmed,     float(ghl_confirmed),      False),
            ("Cancellation Rate",COL_CANCEL_RATE,   sheet_cancel_rate,   ghl_cancel_rate,           True),
        ]

        mismatches = []
        lines      = []

        for label, col_idx, s_val, g_val, is_float in checks:
            if is_float:
                match = abs(s_val - g_val) <= 0.005
            else:
                match = int(s_val) == int(g_val)

            if match:
                lines.append(f"    {label}: sheet={s_val}, GHL={g_val} ✅")
            else:
                lines.append(f"    {label}: sheet={s_val}, GHL={g_val} ❌ MISMATCH")
                mismatches.append((label, col_idx, s_val, g_val))

        # ── Formula validation ─────────────────────────────────────────────────
        hook_formula_ok  = abs(sheet_hook_pct  - calc_hook_pct)  <= 0.01
        offer_formula_ok = abs(sheet_offer_pct - calc_offer_pct) <= 0.01

        if hook_formula_ok:
            lines.append(f"    Hook %:  sheet={sheet_hook_pct}, calculated={calc_hook_pct} ✅")
        else:
            lines.append(
                f"    Hook %:  sheet={sheet_hook_pct}, calculated={calc_hook_pct} ❌ MISMATCH"
                f" — trigger_links={int(sheet_trigger_links)}, new_patients={int(sheet_new_patients)}"
                f", expected={calc_hook_pct}"
            )
            mismatches.append(("Hook %", COL_HOOK_PCT, sheet_hook_pct, calc_hook_pct))

        if offer_formula_ok:
            lines.append(f"    Offer %: sheet={sheet_offer_pct}, calculated={calc_offer_pct} ✅")
        else:
            lines.append(
                f"    Offer %: sheet={sheet_offer_pct}, calculated={calc_offer_pct} ❌ MISMATCH"
                f" — reviews={int(sheet_reviews)}, trigger_links={int(sheet_trigger_links)}"
                f", expected={calc_offer_pct}"
            )
            mismatches.append(("Offer %", COL_OFFER_PCT, sheet_offer_pct, calc_offer_pct))

        # ── Print block ────────────────────────────────────────────────────────
        if mismatches:
            print(f"\n  ⚠️  {name} — DISCREPANCIES FOUND:")
            discrepant_clients.append(name)
        else:
            print(f"\n  ✅ {name} — all values match")
            clean_clients.append(name)

        for line in lines:
            print(line)

        # ── Auto-fix ───────────────────────────────────────────────────────────
        if fix and mismatches:
            print(f"    [fix] Applying corrections to row {row_num}...")
            for label, col_idx, old_val, new_val in mismatches:
                # Only fix GHL-sourced fields (not formula recalculations from sheet data)
                ghl_fields = {"New Patients", "Trigger Links", "Confirmed Appts", "Cancellation Rate"}
                if label not in ghl_fields:
                    # Formula fields: recalculate and write
                    pass
                success = write_audit_cell(name, row_num, col_idx, new_val)
                status  = "✅" if success else "❌ FAILED"
                print(f"    [fix] {name} — updated {label} from {old_val} to {new_val} {status}")
                logger.info(
                    f"[audit][fix] {name} — updated {label} "
                    f"col={col_idx} row={row_num}: {old_val} → {new_val}"
                )

        print()

    # ── Final summary ──────────────────────────────────────────────────────────
    print("=" * 60)
    print(
        f"Audit complete — {len(clean_clients)} client(s) clean, "
        f"{len(discrepant_clients)} client(s) have discrepancies"
    )
    if discrepant_clients:
        print(f"Discrepancy clients: {discrepant_clients}")
    if fix and discrepant_clients:
        print("Corrections have been written to the sheet.")
    print("=" * 60)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if "--eid-message" in sys.argv:
        import eid_message
        _all_clients, _valid_clients = _load_clients()
        eid_message.run_eid(_valid_clients, _settings)
        sys.exit(0)

    if "--payment-reminder" in sys.argv:
        import eid_message
        _all_clients, _valid_clients = _load_clients()
        eid_message.run_payment_reminder(_valid_clients, _settings)
        sys.exit(0)

    if "--good-report" in sys.argv:
        _all_clients, _valid_clients = _load_clients()
        good_report_engine.run(_valid_clients, _settings)
        sys.exit(0)

    if "--good-report-force" in sys.argv:
        _force_idx = sys.argv.index("--good-report-force")
        _force_name = sys.argv[_force_idx + 1] if _force_idx + 1 < len(sys.argv) else ""
        if not _force_name:
            print("[X] Usage: --good-report-force \"Client Name\"")
            sys.exit(1)
        _all_clients, _valid_clients = _load_clients()
        good_report_engine.run_force(_valid_clients, _settings, _force_name)
        sys.exit(0)

    if "--schedule" in sys.argv:
        import json as _json
        _all_clients, _valid_clients = _load_clients()
        _schedule = good_report_engine.get_full_schedule(_valid_clients, _settings)
        print(_json.dumps(_schedule, default=str, ensure_ascii=False, indent=2))
        good_report_engine.write_schedule_to_sheet(_valid_clients, _settings)
        good_report_engine.write_history_to_sheet()
        good_report_engine.write_future_events_to_sheet(_valid_clients)
        sys.exit(0)

    if "--breach-audit" in sys.argv:
        import breach_monitor
        _all_clients, _ = _load_clients()
        _dry_run = "--dry-run" in sys.argv
        breach_monitor.run(_all_clients, _settings, dry_run=_dry_run)
        sys.exit(0)

    if "--reconcile-counters" in sys.argv:
        import reconcile_counters
        _all_clients, _ = _load_clients()
        reconcile_counters.run(_all_clients, _settings)
        sys.exit(0)

    if "--apply-counters" in sys.argv:
        import reconcile_counters
        _all_clients, _ = _load_clients()
        reconcile_counters.run_apply(_all_clients, _settings)
        sys.exit(0)

    if "--catchup-notes" in sys.argv:
        import reconcile_counters
        _all_clients, _ = _load_clients()
        _dry = "--apply" not in sys.argv
        reconcile_counters.run_catchup_notes(_all_clients, _settings, dry_run=_dry)
        sys.exit(0)

    if "--validate-tabs" in sys.argv:
        good_report_engine.validate_tabs()
        sys.exit(0)

    if "--dry-audit" in sys.argv:
        _all_clients, _valid_clients = _load_clients()
        good_report_engine.dry_audit(_all_clients)
        sys.exit(0)

    if "--clear" in sys.argv:
        clear_all_sheets()
        sys.exit(0)

    if "--watch-only" in sys.argv:
        watch_only()
        sys.exit(0)

    if "--update-headers" in sys.argv:
        print("\nUpdating headers for all client sheets...\n")
        force_update_all_headers()
        print("\nDone.")
        sys.exit(0)

    if "--status" in sys.argv:
        from sheets_manager import get_all_sheet_rows
        print("\nClient confirmation status across all sheets:\n")
        clients, _ = _load_clients()
        IDX_CONFIRM = 2   # col C
        IDX_AC      = 28  # col AC

        def _cell(row, idx):
            return str(row[idx]).strip().upper() if len(row) > idx else ""

        n_done, n_skip, n_failed, n_none = 0, 0, 0, 0
        for client in clients:
            name = client["name"]
            try:
                all_rows = get_all_sheet_rows(name)
                if not all_rows:
                    print(f"  \u2796 {name} \u2014 no data rows")
                    n_none += 1
                    continue

                yes_rows    = [r for r in all_rows if _cell(r, IDX_CONFIRM) == "YES"]
                if not yes_rows:
                    print(f"  \u2796 {name} \u2014 no confirmations")
                    n_none += 1
                    continue

                skip_rows   = [r for r in yes_rows if _cell(r, IDX_AC) == "SKIP"]
                failed_rows = [r for r in yes_rows if _cell(r, IDX_AC) == "FAILED"]
                done_rows   = [r for r in yes_rows if _cell(r, IDX_AC) == "TRUE"]

                if skip_rows:
                    print(f"  \u26A0\uFE0F  {name} \u2014 {len(skip_rows)} row(s) SKIP (bad contact_id)")
                    n_skip += 1
                elif failed_rows:
                    print(f"  \u23F3 {name} \u2014 {len(failed_rows)} row(s) FAILED (pending retry)")
                    n_failed += 1
                elif len(done_rows) == len(yes_rows):
                    print(f"  \u2705 {name} \u2014 all rows done ({len(done_rows)} confirmed)")
                    n_done += 1
                else:
                    pending = len(yes_rows) - len(done_rows)
                    print(f"  \u23F3 {name} \u2014 {pending} row(s) pending confirmation")
                    n_failed += 1
            except Exception as e:
                print(f"  \u2753 {name} \u2014 error: {e}")
                n_none += 1

        print(
            f"\nTotals: {n_done} fully done, {n_skip} skipped, "
            f"{n_failed} pending, {n_none} not confirmed yet"
        )
        sys.exit(0)

    if "--backfill-reviews" in sys.argv:
        print("\nBackfilling Total Google Reviews (col AE) for all clients...\n")
        clients, _ = _load_clients()
        total_written = 0
        for client in clients:
            name = client["name"]
            n = backfill_total_reviews(name)
            status = f"{n} row(s) written" if n > 0 else "nothing to fill"
            print(f"  {name} — {status}")
            total_written += n
        print(f"\nDone. {total_written} total row(s) backfilled.")
        sys.exit(0)

    if "--reset-counters" in sys.argv:
        import json as _json
        import requests as _requests
        _settings_path = BASE_DIR / "config" / "settings.json"
        try:
            with open(_settings_path, "r", encoding="utf-8") as _sf:
                _settings = _json.load(_sf)
        except Exception as _e:
            print(f"[X] Cannot load settings.json: {_e}")
            sys.exit(1)
        _api_key   = _settings.get("threeup_api_key", "")
        _field_id  = _settings.get("counter_field_id", "nNx5vev4O2dBgLbqYNSh")
        if not _api_key:
            print("[X] threeup_api_key missing from settings.json")
            sys.exit(1)
        _all_clients, _ = _load_clients()
        _ok = 0
        _headers = {
            "Authorization": f"Bearer {_api_key}",
            "Version": "2021-07-28",
            "Content-Type": "application/json",
        }
        print(f"\nResetting counters for {len(_all_clients)} client(s)...\n")
        for _c in _all_clients:
            _name = _c.get("name", "<unnamed>")
            _cid  = _c.get("contact_id", "")
            if not _cid:
                print(f"  [reset] {_name} — SKIPPED (no contact_id)")
                continue
            try:
                _r = _requests.put(
                    f"https://services.leadconnectorhq.com/contacts/{_cid}",
                    json={"customFields": [{"id": _field_id, "field_value": 0}]},
                    headers=_headers,
                    timeout=30,
                )
                if _r.status_code in (200, 201):
                    print(f"  [reset] {_name} counter reset to 0 \u2705")
                    _ok += 1
                else:
                    print(f"  [reset] {_name} FAILED: {_r.status_code} — {_r.text[:120]}")
            except Exception as _re:
                print(f"  [reset] {_name} ERROR: {_re}")
        print(f"\nReset complete — {_ok}/{len(_all_clients)} counters cleared")
        sys.exit(0)

    if "--run-governor" in sys.argv:
        import omni_governor as _og
        from sheets_manager import get_all_sheet_rows as _gasr
        _all_clients, _ = _load_clients()
        print(f"\nRunning omni_governor for {len(_all_clients)} client(s)...\n")
        _strikes = {0: 0, 1: 0, 2: 0, "3+": 0}

        def _calc_strikes(rows):
            count = 0
            for row in reversed(rows):
                h = str(row[10]).strip().upper() if len(row) > 10 else "FALSE"
                o = str(row[11]).strip().upper() if len(row) > 11 else "FALSE"
                c = str(row[16]).strip().upper() if len(row) > 16 else "FALSE"
                if h == "TRUE" or o == "TRUE" or c == "TRUE":
                    count += 1
                else:
                    break
            return count

        for _c in _all_clients:
            _name = _c.get("name", "<unnamed>")
            try:
                _rows = _gasr(_name)
                _n    = _calc_strikes(_rows)
                _og.run(_c, _rows)
                print(f"  [governor] {_name} \u2014 Strike {_n}")
                if _n == 0:
                    _strikes[0] += 1
                elif _n == 1:
                    _strikes[1] += 1
                elif _n == 2:
                    _strikes[2] += 1
                else:
                    _strikes["3+"] += 1
            except Exception as _ge:
                print(f"  [governor] {_name} ERROR: {_ge}")
        print(
            f"\nGovernor complete \u2014 {_strikes[0]} clients at Strike 0, "
            f"{_strikes[1]} at Strike 1, {_strikes[2]} at Strike 2, "
            f"{_strikes['3+']} at Strike 3+"
        )
        sys.exit(0)

    if "--fix-week5-label" in sys.argv:
        from sheets_manager import _get_services, _get_spreadsheet_id, _col_letter
        _all_clients, _ = _load_clients()
        _old_label = "Sat 18/4 \u2013 Fri 24/4"
        _new_label = "Sat 18/4 \u2013 Thu 23/4"
        print(f"\nFixing Week 5 label: '{_old_label}' \u2192 '{_new_label}'\n")
        _fixed = 0
        for _c in _all_clients:
            _name = _c.get("name", "<unnamed>")
            try:
                _sheets, _ = _get_services()
                _sid = _get_spreadsheet_id(_name)
                if not _sid:
                    print(f"  [fix] {_name} — SKIPPED (no sheet_id)")
                    continue
                _result = _sheets.spreadsheets().values().get(
                    spreadsheetId=_sid,
                    range="Sheet1!A:A",
                    valueRenderOption="UNFORMATTED_VALUE",
                ).execute()
                _col_a = _result.get("values", [])
                _row_num = None
                for _ri, _rv in enumerate(_col_a):
                    if _rv and str(_rv[0]).strip() == _old_label:
                        _row_num = _ri + 1  # 1-based
                        break
                if _row_num is None:
                    print(f"  [fix] {_name} — label not found, skipping")
                    continue
                _sheets.spreadsheets().values().update(
                    spreadsheetId=_sid,
                    range=f"Sheet1!A{_row_num}",
                    valueInputOption="RAW",
                    body={"values": [[_new_label]]},
                ).execute()
                print(f"  [fix] {_name} \u2014 Week 5 label corrected (row {_row_num})")
                _fixed += 1
            except Exception as _fe:
                print(f"  [fix] {_name} ERROR: {_fe}")
        print(f"\nDone \u2014 {_fixed}/{len(_all_clients)} client(s) updated.")
        sys.exit(0)

    if "--fix-week7-label" in sys.argv:
        from sheets_manager import _get_services, _get_spreadsheet_id
        _all_clients, _ = _load_clients()
        _old_label = "Sat 2/5 \u2013 Thu 7/5"
        _new_label = "Sat 2/5 \u2013 Fri 8/5"
        print(f"\nFixing Week 7 label: '{_old_label}' \u2192 '{_new_label}'\n")
        _fixed = 0
        for _c in _all_clients:
            _name = _c.get("name", "<unnamed>")
            try:
                _sheets, _ = _get_services()
                _sid = _get_spreadsheet_id(_name)
                if not _sid:
                    print(f"  [fix] {_name} \u2014 SKIPPED (no sheet_id)")
                    continue
                _result = _sheets.spreadsheets().values().get(
                    spreadsheetId=_sid,
                    range="Sheet1!A:A",
                    valueRenderOption="UNFORMATTED_VALUE",
                ).execute()
                _col_a = _result.get("values", [])
                _row_num = None
                for _ri, _rv in enumerate(_col_a):
                    if _rv and str(_rv[0]).strip() == _old_label:
                        _row_num = _ri + 1  # 1-based
                        break
                if _row_num is None:
                    print(f"  [fix] {_name} \u2014 label not found, skipping")
                    continue
                _sheets.spreadsheets().values().update(
                    spreadsheetId=_sid,
                    range=f"Sheet1!A{_row_num}",
                    valueInputOption="RAW",
                    body={"values": [[_new_label]]},
                ).execute()
                print(f"  [fix] {_name} \u2014 Week 7 label corrected (row {_row_num})")
                _fixed += 1
            except Exception as _fe:
                print(f"  [fix] {_name} ERROR: {_fe}")
        print(f"\nDone \u2014 {_fixed}/{len(_all_clients)} client(s) updated.")
        sys.exit(0)

    if "--clear-last-sent" in sys.argv:
        import requests as _requests
        _arg_idx = sys.argv.index("--clear-last-sent")
        try:
            _target_name = sys.argv[_arg_idx + 1]
        except IndexError:
            print("[X] --clear-last-sent requires a client name argument")
            sys.exit(1)
        _all_clients, _ = _load_clients()
        _match = next(
            (c for c in _all_clients
             if c.get("name") == _target_name or c.get("am_name") == _target_name),
            None
        )
        if not _match:
            print(f"[X] Client not found: {_target_name!r}")
            print(f"    Available: {[c['name'] for c in _all_clients]}")
            sys.exit(1)
        _cls_api_key    = _settings.get("threeup_api_key", "")
        _cls_contact_id = _match.get("contact_id", "")
        _cls_name       = _match.get("name", "")
        if not _cls_api_key:
            print("[X] threeup_api_key missing from settings.json")
            sys.exit(1)
        if not _cls_contact_id:
            print(f"[X] contact_id missing for {_cls_name!r}")
            sys.exit(1)
        _cls_url = f"https://services.leadconnectorhq.com/contacts/{_cls_contact_id}"
        _cls_headers = {
            "Authorization": f"Bearer {_cls_api_key}",
            "Version": "2021-07-28",
            "Content-Type": "application/json",
        }
        _cls_payload = {"customFields": [{"id": "M70Sd18pm9ZPg35ecix1", "field_value": ""}]}
        _cls_r = _requests.put(_cls_url, headers=_cls_headers, json=_cls_payload, timeout=30)
        if _cls_r.status_code == 200:
            print(f"Cleared good_report_last_sent for {_cls_name}")
        else:
            print(f"[X] GHL returned {_cls_r.status_code}: {_cls_r.text[:200]}")
            sys.exit(1)
        sys.exit(0)

    if "--audit" in sys.argv:
        _audit_week = None
        for _arg in sys.argv[1:]:
            if _arg.startswith("--week"):
                try:
                    _audit_week = int(_arg.split("=", 1)[1]) if "=" in _arg \
                        else int(sys.argv[sys.argv.index(_arg) + 1])
                except (ValueError, IndexError):
                    print(f"[X] --audit --week requires a number (e.g. --week 4)")
                    sys.exit(1)
                break
        audit_mode(fix="--fix" in sys.argv, week_index=_audit_week)
        sys.exit(0)

    if "--all-weeks" in sys.argv:
        all_weeks_mode()
        sys.exit(0)

    if "--daily-catchup" in sys.argv:
        import daily_catchup as _daily_catchup
        _dc_clients, _ = _load_clients()
        _daily_catchup.run(_dc_clients, _settings)
        sys.exit(0)

    # --week N  (e.g. --week 1  or  --week=2)
    week_num = None
    for _arg in sys.argv[1:]:
        if _arg.startswith("--week"):
            try:
                if "=" in _arg:
                    week_num = int(_arg.split("=", 1)[1])
                else:
                    _idx = sys.argv.index(_arg)
                    week_num = int(sys.argv[_idx + 1])
            except (ValueError, IndexError):
                print(f"[X] Usage: --week 1|2|3|4|5|6  (got: {_arg})")
                sys.exit(1)
            break

    try:
        scrape_form_fields()
    except Exception as e:
        ts = datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M:%S")
        print(f"[!] Form field scrape failed: {e} — continuing.")
        logger.error(f"[{ts}] [main] scrape_form_fields failed: {e}")

    if "--client" in sys.argv:
        _ci = sys.argv.index("--client")
        _client_name = sys.argv[_ci + 1] if _ci + 1 < len(sys.argv) else ""
        if not _client_name:
            print("[X] Usage: --client \"Client Name\" [--week N]")
            sys.exit(1)
        _all_clients, _ = _load_clients()
        _target = next(
            (c for c in _all_clients if c.get("name", "").strip().lower() == _client_name.strip().lower()),
            None
        )
        if _target is None:
            print(f"[X] Client '{_client_name}' not found in clients.json")
            print(f"    Available: {[c['name'] for c in _all_clients]}")
            sys.exit(1)
        now = datetime.now(AMMAN_TZ)
        print_banner(now)
        if week_num is not None:
            if week_num < 1 or week_num > len(ALL_WEEKS):
                print(f"[X] --week must be 1–{len(ALL_WEEKS)} (got {week_num})")
                sys.exit(1)
            _week_def = ALL_WEEKS[week_num - 1]
            this_start, this_end, last_start, last_end = _localize_week(_week_def)
            print(f"  Client    : {_target['name']}")
            print(f"  Week      : {_week_def['label']} — {_fmt_date_range(this_start, this_end)}")
            print(f"  Last week : {_fmt_date_range(last_start, last_end)}")
            _skip_dup = False  # targeted historical rerun — always re-pull
        else:
            this_start, this_end, last_start, last_end = get_week_bounds(now)
            print(f"  Client    : {_target['name']}")
            print(f"  This week : {_fmt_date_range(this_start, this_end)}")
            print(f"  Last week : {_fmt_date_range(last_start, last_end)}")
            _skip_dup = True   # current week — skip if row already exists
        print()
        _ghl_cache = {}
        _pull_and_append_week(
            [_target], this_start, this_end, last_start, last_end,
            _ghl_cache, skip_dup=_skip_dup
        )
        print(f"[OK] Single-client run complete for '{_target['name']}'")
        sys.exit(0)

    if week_num is not None:
        week_n_mode(week_num)
        sys.exit(0)

    main()  # dynamic mode: auto-calculates current Sat–Fri week
