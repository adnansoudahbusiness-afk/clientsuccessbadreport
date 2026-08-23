import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

"""
watcher.py — Background watcher thread for ThreeUp CAO Agent.
Polls every 120 seconds. On CONFIRM = YES: calculates flags, records to sheet, marks as done.
Exits after all clients confirmed or 48-hour timeout.
"""

import json
import logging
import os
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import pytz

_settings = json.load(open(
    os.path.join(os.path.dirname(__file__), '..', 'config', 'settings.json'),
    encoding='utf-8'))
_good_report_run_dates: set = set()

_agent_dir = os.path.dirname(os.path.abspath(__file__))
if _agent_dir not in sys.path:
    sys.path.insert(0, _agent_dir)

import good_report_engine
from notifier_email import send_email

AMMAN_TZ          = pytz.timezone("Asia/Amman")
POLL_INTERVAL_SEC = 120
TIMEOUT_HOURS     = 48

logger = logging.getLogger("watcher")


def _ts() -> str:
    return datetime.now(AMMAN_TZ).strftime("%Y-%m-%d %H:%M:%S")


class Watcher(threading.Thread):
    """
    Daemon thread. Watches Google Sheets and triggers diagnosis + form submission
    when users set CONFIRM = YES.
    """

    def __init__(self, clients: list, ghl_data_cache: dict = None):
        """
        clients         — list of client dicts from clients.json
        ghl_data_cache  — dict keyed by client name with {messages_list: [...]}
        """
        super().__init__(daemon=True)
        self.clients        = clients
        self.ghl_data_cache = ghl_data_cache or {}
        self.confirmed_set  = set()
        self.start_time     = None
        self._stop_event    = threading.Event()

    def stop(self):
        self._stop_event.set()

    def run(self):
        # Ensure agent dir is on path for imports
        agent_dir = Path(__file__).resolve().parent
        if str(agent_dir) not in sys.path:
            sys.path.insert(0, str(agent_dir))

        from sheets_manager import (
            get_all_confirmable_rows,
            mark_form_submitted,
            mark_row_skip,
            calculate_and_flag,
            get_all_sheet_rows,
            count_consecutive_low_appointments,
        )
        from form_submitter import send_hovara_referral
        import omni_governor

        self.start_time     = time.time()
        timeout_seconds     = TIMEOUT_HOURS * 3600
        valid_clients       = [c for c in self.clients if self._is_configured(c)]

        print(f"[{_ts()}] Watcher started — polling every {POLL_INTERVAL_SEC}s "
              f"(48h timeout, {len(valid_clients)} client(s))")
        logger.info(f"[{_ts()}] [watcher] Started — {len(valid_clients)} valid clients")

        while not self._stop_event.is_set():
            # 48h timeout guard
            if time.time() - self.start_time > timeout_seconds:
                print(f"[{_ts()}] Watcher 48h timeout reached — stopping.")
                logger.info(f"[{_ts()}] [watcher] 48h timeout — exiting")
                break

            # All done?
            if len(self.confirmed_set) >= len(valid_clients):
                print(f"[{_ts()}] All clients confirmed and processed — watcher complete.")
                logger.info(f"[{_ts()}] [watcher] All clients done — exiting")
                break

            poll_total_rows     = 0
            poll_active_clients = 0
            flagged_count       = 0

            for client in valid_clients:
                name = client["name"]

                if name in self.confirmed_set:
                    continue

                try:
                    # Get every row that has CONFIRM=YES and hasn't been submitted yet
                    confirmable_rows = get_all_confirmable_rows(name)
                    if not confirmable_rows:
                        continue

                    print(f"\n[{_ts()}] [{name}] {len(confirmable_rows)} confirmable row(s) found — processing...")
                    logger.info(f"[{_ts()}] [{name}] [watcher] {len(confirmable_rows)} confirmable row(s)")

                    poll_total_rows     += len(confirmable_rows)
                    poll_active_clients += 1

                    all_rows_ok     = True
                    _client_flagged = False

                    for conf_row in confirmable_rows:
                        row_index = conf_row["row_index"]
                        row_data  = conf_row["data"]

                        flag_results = calculate_and_flag(name, row_num=row_index, row_data=row_data)
                        logger.info(f"[{_ts()}] [{name}] [watcher] Row {row_index} flags: {flag_results}")

                        flag_map = {
                            "Hook":         flag_results.get("hook", False),
                            "Offer":        flag_results.get("offer", False),
                            "Cancellation": flag_results.get("cancellation", False),
                            "Loyalty":      flag_results.get("loyalty", False),
                        }

                        active_flags = [k for k, v in flag_map.items() if v]
                        if active_flags:
                            _client_flagged = True
                            logger.info(f"[{_ts()}] [{name}] [watcher] Row {row_index}: flags {active_flags} written to sheet")
                            print(f"[{_ts()}] [{name}] Row {row_index} — flags: {', '.join(active_flags)}")
                        else:
                            logger.info(f"[{_ts()}] [{name}] [watcher] Row {row_index}: no flags")

                        mark_form_submitted(name, all_succeeded=True, row_num=row_index)

                    if _client_flagged:
                        flagged_count += 1

                    # Step 5: After all rows processed — run omni_governor once per client
                    # Always runs regardless of submission success/failure
                    try:
                        all_rows = get_all_sheet_rows(name)
                        omni_governor.run(client, all_rows)
                    except Exception as og_err:
                        logger.warning(f"[{_ts()}] [{name}] [omni_governor] Error: {og_err}")

                    # Step 6: Hovara referral if streak >= 3 consecutive low-appointment weeks
                    try:
                        streak = count_consecutive_low_appointments(name)
                        if streak >= 3:
                            _threeup_api_key = _settings.get("threeup_api_key", "")
                            send_hovara_referral(
                                client_name=name,
                                contact_id=client.get("contact_id", ""),
                                streak=streak,
                                threeup_api_key=_threeup_api_key,
                                hovara_field_id=None,
                            )
                    except Exception as hov_err:
                        logger.warning(f"[{_ts()}] [{name}] [hovara] Error: {hov_err}")

                    if all_rows_ok:
                        self.confirmed_set.add(name)

                except Exception as e:
                    logger.error(f"[{_ts()}] [{name}] [watcher] Unexpected error: {e}", exc_info=True)
                    print(f"[{_ts()}] [{name}] ERROR in watcher: {e}")
                finally:
                    time.sleep(1)   # brief pause between client checks

            # Poll summary
            if poll_total_rows > 0:
                print(f"[{_ts()}] [watcher] Poll complete — {poll_total_rows} confirmable row(s) found across {poll_active_clients} client(s)")
                logger.info(f"[{_ts()}] [watcher] Poll complete — {poll_total_rows} row(s), {poll_active_clients} client(s)")
            else:
                print(f"[{_ts()}] [watcher] Polling... (no confirmations yet)")

            # Per-client status summary (uses raw list rows: col C=2=CONFIRM, col AC=28)
            print(f"[{_ts()}] [watcher] Client summary:")
            IDX_CONFIRM = 2   # col C
            IDX_AC      = 28  # col AC
            for client in valid_clients:
                name = client["name"]
                try:
                    all_rows = get_all_sheet_rows(name)  # list of raw value lists, header skipped
                    if not all_rows:
                        print(f"  \u2796 {name} \u2014 no data")
                        continue

                    def _cell(row, idx):
                        return str(row[idx]).strip().upper() if len(row) > idx else ""

                    yes_rows    = [r for r in all_rows if _cell(r, IDX_CONFIRM) == "YES"]
                    if not yes_rows:
                        print(f"  \u2796 {name} \u2014 no confirmations")
                        continue

                    skip_rows   = [r for r in yes_rows if _cell(r, IDX_AC) == "SKIP"]
                    failed_rows = [r for r in yes_rows if _cell(r, IDX_AC) == "FAILED"]
                    done_rows   = [r for r in yes_rows if _cell(r, IDX_AC) == "TRUE"]

                    if skip_rows:
                        print(f"  \u26A0\uFE0F  {name} \u2014 {len(skip_rows)} row(s) SKIP (bad contact_id)")
                    elif failed_rows:
                        print(f"  \u23F3 {name} \u2014 {len(failed_rows)} row(s) pending (FAILED)")
                    elif len(done_rows) == len(yes_rows):
                        print(f"  \u2705 {name} \u2014 all rows done")
                    else:
                        pending = len(yes_rows) - len(done_rows)
                        print(f"  \u23F3 {name} \u2014 {pending} row(s) pending confirmation")
                except Exception as summary_err:
                    print(f"  \u2753 {name} \u2014 error reading status: {summary_err}")

            # Email alert if any clients were flagged this poll cycle
            if flagged_count > 0:
                send_email(
                    f"ThreeUp CAO — {flagged_count} Client(s) Flagged 🚨",
                    f"Watcher processed confirmations at {datetime.now().strftime('%d/%m/%Y %H:%M')}.\n"
                    f"Clients with active flags: {flagged_count}\n"
                    f"Flags recorded to sheet. Check dashboard for details.\n"
                    f"clientsuccess.threeupworld.com"
                )

            # Daily good_report run — once per calendar day, only between 09:00–11:59 Amman
            _now_amman = datetime.now(AMMAN_TZ)
            _today     = _now_amman.date()
            _hour      = _now_amman.hour
            if _today not in _good_report_run_dates and 9 <= _hour <= 11:
                _good_report_run_dates.add(_today)
                gr_thread = threading.Thread(
                    target=good_report_engine.run,
                    args=(self.clients, _settings),
                    daemon=True,
                )
                gr_thread.start()

            # Sleep in 1s increments so we can stop cleanly
            for _ in range(POLL_INTERVAL_SEC):
                if self._stop_event.is_set():
                    break
                time.sleep(1)

        logger.info(f"[{_ts()}] [watcher] Thread exiting")

    @staticmethod
    def _is_configured(client: dict) -> bool:
        """Return False if client still has placeholder values."""
        placeholders = {"PASTE_YOUR_SUBACCOUNT_API_KEY_HERE", "PASTE_YOUR_LOCATION_ID_HERE"}
        return (
            client.get("ghl_api_key", "")  not in placeholders and
            client.get("location_id", "") not in placeholders
        )
