import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

"""
good_report_engine.py — Generates and sends periodic Good Reports to clients via GHL.
Runs on a schedule: sustainability every 14 days, renewal reminder 6 days before payment.
"""

import json
import logging
import os
import time
import requests
import pytz
from datetime import datetime, timedelta, date

from google.oauth2 import service_account
from googleapiclient.discovery import build
from notifier_email import send_email

logger = logging.getLogger("good_report_engine")

AM_DATES_SHEET_ID  = "12KEc1_CIkAHpfA74y660zsWSGnkbcSoltcSosuk4smA"
DRIP_DELAY_SECONDS = 1800

_HERE            = os.path.dirname(os.path.abspath(__file__))
_ROOT            = os.path.dirname(_HERE)
CREDENTIALS_PATH = os.path.join(_ROOT, "config", "credentials.json")
_QUEUE_PATH      = os.path.join(_ROOT, "logs",   "send_queue.json")
_HISTORY_PATH    = os.path.join(_ROOT, "logs",   "send_history.json")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

GHL_BASE = "https://services.leadconnectorhq.com"

# Confirmed field IDs (GHL returns key=None on contact GET; must use IDs for reads)
_CAO_GOOD_REPORT_FIELD_ID  = "Y1lUr9X7ACLNmfsXmsCX"
_GOOD_REPORT_LAST_SENT_ID  = "M70Sd18pm9ZPg35ecix1"


# ── History helper ────────────────────────────────────────────────────────────

def _append_history(entry: dict) -> None:
    try:
        history = []
        if os.path.exists(_HISTORY_PATH):
            with open(_HISTORY_PATH, "r", encoding="utf-8") as f:
                history = json.load(f)
    except Exception:
        history = []
    history.append(entry)
    try:
        os.makedirs(os.path.dirname(_HISTORY_PATH), exist_ok=True)
        with open(_HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2, ensure_ascii=False, default=str)
    except Exception as e:
        print(f"[history] Failed to write: {e}")


# ── Google Sheets helper ──────────────────────────────────────────────────────

def _get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_PATH, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# ── FUNCTION 1: load_am_dates ─────────────────────────────────────────────────

_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
_MONTH_ORDER = {name: i for i, name in enumerate(_MONTH_NAMES)}


def _tab_to_month(title: str):
    """Return canonical month name for a tab title, or None if not a month tab.
    Handles both bare 'June' and year-suffixed 'June 2026' tab names."""
    parts = title.strip().split()
    if parts and parts[0] in _MONTH_ORDER:
        return parts[0]
    return None


def _parse_date(s: str):
    from datetime import datetime, date as _date
    for fmt in ["%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"]:
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except Exception:
            continue
    # DD/MM without year — assume 2026
    try:
        parts = s.strip().split("/")
        if len(parts) == 2:
            return _date(2026, int(parts[1]), int(parts[0]))
    except Exception:
        pass
    return None


def _read_tab(service, title: str) -> list:
    """Parse one sheet tab and return a list of entry dicts."""
    data = service.spreadsheets().values().get(
        spreadsheetId=AM_DATES_SHEET_ID,
        range=f"'{title}'!A:F",
    ).execute()

    rows = data.get("values", [])
    print(f"[good_report] Tab '{title}': {len(rows)} rows")

    _SKIP = {"waiting", "pause", ""}
    results = []
    for i, row in enumerate(rows):
        if i == 0:
            continue  # header
        if len(row) < 4:
            continue

        client_name = str(row[0]).strip()
        sus_raw     = str(row[1]).strip() if len(row) > 1 else ""
        renewal_raw = str(row[2]).strip() if len(row) > 2 else ""
        payment_raw = str(row[3]).strip() if len(row) > 3 else ""
        term        = str(row[4]).strip() if len(row) > 4 else ""

        if not client_name or sus_raw.lower() in _SKIP:
            continue
        if renewal_raw.lower() in _SKIP or payment_raw.lower() in _SKIP:
            continue

        sus_date     = _parse_date(sus_raw)
        renewal_date = _parse_date(renewal_raw)
        payment_date = _parse_date(payment_raw)

        if not sus_date or not renewal_date or not payment_date:
            print(f"  -> Skipped {client_name}: bad dates sus='{sus_raw}' renewal='{renewal_raw}' payment='{payment_raw}'")
            continue

        results.append({
            "client_name":               client_name,
            "first_sustainability_date": sus_date,
            "renewal_date":              renewal_date,
            "payment_date":              payment_date,
            "term":                      term,
        })
        print(f"  -> {client_name}: sus={sus_date} renewal={renewal_date} payment={payment_date}")

    return results


def load_all_am_tabs() -> dict:
    """Read every available month tab. Returns {canonical_month: [entries]}.
    Accepts tab titles like 'June' or 'June 2026' — both map to key 'June'."""
    try:
        service    = _get_sheets_service()
        spreadsheet = service.spreadsheets().get(spreadsheetId=AM_DATES_SHEET_ID).execute()
        available  = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]
        result = {}
        for title in available:
            month = _tab_to_month(title)
            if month:
                entries = _read_tab(service, title)  # reads by exact tab title
                if entries:
                    result[month] = entries  # stored by canonical month name
        return result
    except Exception as e:
        print(f"[good_report] load_all_am_tabs failed: {e}")
        import traceback; traceback.print_exc()
        return {}


def find_active_term(am_name: str, all_tabs: dict, today):
    """Return the term entry for am_name whose window (first_sus..payment_date) contains today.
    Falls back to most recent past term, then earliest future term."""
    am_lower   = am_name.strip().lower()
    candidates = []
    for tab_name in sorted(all_tabs.keys(), key=lambda t: _MONTH_ORDER.get(t, 99)):
        for entry in all_tabs[tab_name]:
            if entry["client_name"].strip().lower() == am_lower:
                candidates.append(entry)
                break
    if not candidates:
        return None
    active = [e for e in candidates
              if e["first_sustainability_date"] <= today <= e["payment_date"]]
    if active:
        return active[-1]
    past = [e for e in candidates if e["payment_date"] < today]
    if past:
        return past[-1]
    return candidates[0]


def _walk_cadence_events(anchor, all_tabs, am_name, today, lookback_days=7, horizon_days=90):
    """Walk cadence events forward from anchor, resetting at each renewal boundary.

    Rule: for each candidate sus date D (anchor, anchor+14, ...):
      - If renewal_reminder <= D: emit renewal_reminder instead, reset anchor to
        payment_date, advance to payment_date+14, load next term.
      - Else: emit D as sustainability, advance +14.

    Only events in [today-lookback_days, today+horizon_days] are returned.
    The walk still processes dates outside the window so renewal resets fire correctly.
    Stops when the sheet has no next-term data.

    Returns [(date, event_type), ...] sorted ascending."""
    start_window = today - timedelta(days=lookback_days)
    end_window   = today + timedelta(days=horizon_days)
    events = []

    # Start from the term active at the anchor date (handles multi-term walks correctly)
    current_term = find_active_term(am_name, all_tabs, anchor)
    if not current_term:
        current_term = find_active_term(am_name, all_tabs, today)
    if not current_term:
        return events

    payment_date     = current_term["payment_date"]
    renewal_reminder = payment_date - timedelta(days=6)
    renewal_emitted  = False

    d = anchor
    while d <= end_window:
        if not renewal_emitted and renewal_reminder <= d:
            # This sustainability slot is replaced by the renewal reminder
            if start_window <= renewal_reminder <= end_window:
                events.append((renewal_reminder, "renewal"))
            renewal_emitted = True
            d = payment_date + timedelta(days=14)
            # Find next term: chronological scan for smallest payment_date > current.
            # Cannot use find_active_term(payment+1) here because the next term's
            # first_sus may be after payment+1, failing the active-window check.
            am_lower  = am_name.strip().lower()
            next_term = None
            for _tab in sorted(all_tabs.keys(), key=lambda t: _MONTH_ORDER.get(t, 99)):
                for _e in all_tabs[_tab]:
                    if _e["client_name"].strip().lower() == am_lower:
                        if _e["payment_date"] > payment_date:
                            if next_term is None or _e["payment_date"] < next_term["payment_date"]:
                                next_term = _e
            if next_term:
                payment_date     = next_term["payment_date"]
                renewal_reminder = payment_date - timedelta(days=6)
                renewal_emitted  = False
            else:
                break  # no next-term data in sheet yet
        else:
            if start_window <= d <= end_window:
                events.append((d, "sustainability"))
            d += timedelta(days=14)

    # The cadence walk can step from d < renewal_reminder directly to d > end_window
    # (when anchor+14n never lands on renewal_reminder and horizon_days=0).
    # Catch that case here so a renewal due today is never silently dropped.
    if not renewal_emitted and start_window <= renewal_reminder <= end_window:
        events.append((renewal_reminder, "renewal"))

    return sorted(events, key=lambda x: x[0])


def load_am_dates() -> list:
    """Load per-client active term entries using per-term selection across all month tabs."""
    try:
        today    = datetime.now(pytz.timezone("Asia/Amman")).date()
        all_tabs = load_all_am_tabs()
        if not all_tabs:
            print("[good_report] No AM tabs loaded")
            return []

        # Collect all unique client names across every tab (preserve first-seen order)
        all_names = []
        seen_lower = set()
        for tab_name in sorted(all_tabs.keys(), key=lambda t: _MONTH_ORDER.get(t, 99)):
            for e in all_tabs[tab_name]:
                nm = e["client_name"]
                if nm.strip().lower() not in seen_lower:
                    all_names.append(nm)
                    seen_lower.add(nm.strip().lower())

        results = []
        for name in all_names:
            entry = find_active_term(name, all_tabs, today)
            if entry:
                results.append(entry)
                print(f"  -> {name}: sus={entry['first_sustainability_date']} "
                      f"renewal={entry['renewal_date']} payment={entry['payment_date']}")

        print(f"[good_report] Total loaded: {len(results)} entries")
        return results

    except Exception as e:
        print(f"[good_report] FATAL: {e}")
        import traceback
        traceback.print_exc()
        return []


# ── FUNCTION 2: get_due_messages ──────────────────────────────────────────────

def get_due_messages(clients: list, all_tabs: dict, today) -> list:
    """Return messages due today (or missed in last 7 days) using cadence_anchor walk."""
    due_list       = []
    lookback_start = today - timedelta(days=7)

    for client in clients:
        if client.get("churned"):
            continue

        am_name    = client.get("am_name", client["name"])
        anchor_str = client.get("cadence_anchor")

        if not anchor_str:
            # No sustainability cadence — still fire renewal reminder if in window.
            term = find_active_term(am_name, all_tabs, today)
            if term:
                renewal_reminder = term["payment_date"] - timedelta(days=6)
                if lookback_start <= renewal_reminder <= today:
                    due_list.append({
                        "client_name":      am_name,
                        "message_type":     "renewal",
                        "due_date":         renewal_reminder,
                        "payment_date_str": term["payment_date"].strftime("%d/%m/%Y"),
                        "renewal_date_str": term["renewal_date"].strftime("%d/%m/%Y"),
                        "is_missed":        renewal_reminder < today,
                    })
            continue

        try:
            anchor = datetime.strptime(str(anchor_str), "%Y-%m-%d").date()
        except ValueError:
            continue

        term = find_active_term(am_name, all_tabs, today)
        if not term:
            continue
        payment_date_str = term["payment_date"].strftime("%d/%m/%Y")
        renewal_date_str = term["renewal_date"].strftime("%d/%m/%Y")

        events = _walk_cadence_events(anchor, all_tabs, am_name, today,
                                      lookback_days=7, horizon_days=0)
        for event_date, event_type in events:
            is_missed = event_date < today
            due_list.append({
                "client_name":      am_name,
                "message_type":     event_type,
                "due_date":         event_date,
                "payment_date_str": payment_date_str,
                "renewal_date_str": renewal_date_str,
                "is_missed":        is_missed,
            })

    print(f"[debug] Before dedup: {[(e['client_name'], e['due_date'], e['is_missed']) for e in due_list]}")

    seen = {}
    for entry in due_list:
        seen[f"{entry['client_name']}_{entry['due_date']}"] = entry
    return sorted(seen.values(), key=lambda x: (x["is_missed"], x["due_date"]))


# ── FUNCTION 3: match_client ──────────────────────────────────────────────────

def match_client(am_name: str, clients: list):
    """Find a client dict from clients list that matches the AM sheet name."""
    am = am_name.strip().lower()

    # Pass 0: exact am_name alias match (preferred — avoids all fuzzy collisions)
    for c in clients:
        if c.get("am_name", "").strip().lower() == am:
            return c

    # Pass 1: exact name match
    for c in clients:
        if c["name"].strip().lower() == am:
            return c

    # Pass 2: am_name contained in client name
    for c in clients:
        if am in c["name"].strip().lower():
            return c

    # Pass 3: client name contained in am_name
    for c in clients:
        if c["name"].strip().lower() in am:
            return c

    print(f"[good_report] No match found for: '{am_name}'")
    return None


# ── FUNCTION 4: get_client_performance ───────────────────────────────────────

def _is_confirmed_row(row: list) -> bool:
    if len(row) <= 16:
        return False
    hook   = str(row[10]).strip()
    offer  = str(row[11]).strip()
    cancel = str(row[16]).strip()
    return any(v != "" for v in [hook, offer, cancel])


def get_client_performance(client_name: str, all_rows: list, message_type: str):
    """Extract performance metrics from sheet rows for the given message type."""
    if not all_rows:
        return None

    if message_type == "sustainability":
        confirmed_rows = [r for r in all_rows if _is_confirmed_row(r)]
        if not confirmed_rows:
            return None
        rows = confirmed_rows[-2:]
    else:
        rows = [r for r in all_rows if _is_confirmed_row(r)]
        if not rows:
            return None

    most_recent = rows[-1]

    def _safe(row, idx, default=0.0):
        try:
            if len(row) <= idx:
                return default
            v = row[idx]
            if v in ("", None):
                return default
            return float(str(v).replace("%", "").strip())
        except (ValueError, TypeError):
            return default

    new_reviews      = sum(_safe(r, 1)  for r in rows)
    total_reviews    = _safe(most_recent, 30)
    new_patients     = sum(_safe(r, 3)  for r in rows)
    new_patients_pct = _safe(most_recent, 5)
    cancellation_rate = _safe(most_recent, 15)
    website_patients = sum(_safe(r, 31) for r in rows)

    hook_flagged   = (str(most_recent[10]).strip().upper() == "TRUE"
                      if len(most_recent) > 10 else False)
    offer_flagged  = (str(most_recent[11]).strip().upper() == "TRUE"
                      if len(most_recent) > 11 else False)
    cancel_flagged = (str(most_recent[16]).strip().upper() == "TRUE"
                      if len(most_recent) > 16 else False)

    flags_count = sum([hook_flagged, offer_flagged, cancel_flagged])
    all_clean   = flags_count == 0

    return {
        "new_reviews":       new_reviews,
        "total_reviews":     total_reviews,
        "new_patients":      new_patients,
        "new_patients_pct":  new_patients_pct,
        "cancellation_rate": cancellation_rate,
        "website_patients":  website_patients,
        "hook_flagged":      hook_flagged,
        "offer_flagged":     offer_flagged,
        "cancel_flagged":    cancel_flagged,
        "flags_count":       flags_count,
        "all_clean":         all_clean,
    }


# ── FUNCTION 5: generate_message ─────────────────────────────────────────────

def generate_message(
    doctor_name: str,
    performance: dict,
    message_type: str,
    satisfaction_link: str = "",
    renewal_date_str: str = "",
    payment_date_str: str = "",
) -> str:
    """Build the Arabic good-report message based on performance tone."""
    nr = int(performance["new_reviews"])
    tr = int(performance["total_reviews"])
    np_ = int(performance["new_patients"])
    cr  = performance["cancellation_rate"]
    wp  = int(performance["website_patients"])
    fc  = performance["flags_count"]

    new_reviews_str  = str(nr) if nr > 0 else "—"
    total_reviews_str = str(tr) if tr > 0 else "—"
    new_patients_str = str(np_) if np_ > 0 else "—"
    attendance     = round((1 - cr) * 100, 1)
    attendance_str = f"{attendance}٪"
    website_line     = (f"\n🌐 مرضى الموقع: {wp} مريض جديد عبر الموقع."
                        if wp > 0 else "")
    sat_line = f"\n{satisfaction_link}" if satisfaction_link else ""

    flag_items = []
    if performance["hook_flagged"]:
        flag_items.append("• تحسين رسالة الاستقطاب لزيادة التفاعل.")
    if performance["offer_flagged"]:
        flag_items.append("• تطوير آلية طلب التقييمات.")
    if performance["cancel_flagged"]:
        flag_items.append("• تحسين تأكيد المواعيد لتقليل الإلغاء.")
    flag_section = "\n".join(flag_items)

    positive_items = []
    if nr > 0:
        positive_items.append(f"• +{nr} تقييم جوجل جديد ⭐")
    if np_ > 0:
        positive_items.append(f"• {np_} مريض جديد 👥")
    if wp > 0:
        positive_items.append(f"• {wp} مريض من الموقع 🌐")
    positive_section = ("\n".join(positive_items)
                        if positive_items
                        else "• الفريق يعمل على تحسين الأداء.")

    prefix = "📊 ملخص أداء الفترة الكاملة:\n\n" if message_type == "renewal" else ""
    renewal_reminder = (
        f"\n\n📋 تذكير: موعد التجديد قادم قريباً ({payment_date_str}). "
        f"رح نتواصل معك لإتمام التجديد وضمان استمرارية النمو. 🤝"
        if message_type == "renewal" and payment_date_str else ""
    )

    if tr > 0 and nr > 0:
        reviews_line = f"⭐ تقييمات جوجل: +{nr} تقييم جديد، ليصل الإجمالي إلى {tr} تقييم.\n"
    elif nr > 0:
        reviews_line = f"⭐ تقييمات جوجل: +{nr} تقييم جديد.\n"
    else:
        reviews_line = ""
    metrics = (
        f"{reviews_line}"
        f"👥 المرضى الجدد: {new_patients_str} مريض جديد هالفترة.\n"
        f"📅 معدل الحضور: {attendance_str} من المواعيد تمت بنجاح.{website_line}"
    )

    if fc == 0:
        message = (
            f"{prefix}مرحبا {doctor_name}، إن شاء الله أمورك تمام 😊\n\n"
            f"حبينا نبعثلك ونتطمن، وكمان نشاركك أبرز ما حققناه في آخر أسبوعين:\n\n"
            f"{metrics}\n\n"
            f"✅ كل المؤشرات ممتازة — استمر على هذا المستوى!\n\n"
            f"هذه العملية تهدف لضمان استمرارية نمو العيادة — والالتزام بالاستمرارية هو ما سيصنع الفارق الحقيقي 💪\n\n"
            f"يهمنا رأيك — كيف تقيم رضاك عن النتائج من 1 إلى 5؟{sat_line}{renewal_reminder}"
        )
    elif fc <= 2:
        message = (
            f"{prefix}مرحبا {doctor_name}، إن شاء الله أمورك تمام 😊\n\n"
            f"حبينا نبعثلك ونتطمن، وكمان نشاركك أبرز ما حققناه في آخر أسبوعين:\n\n"
            f"{metrics}\n\n"
            f"✅ الإيجابيات:\n{positive_section}\n\n"
            f"🔧 نشتغل عليه:\n{flag_section}\n\n"
            f"هذه العملية تهدف لضمان استمرارية نمو العيادة — والالتزام بالاستمرارية هو ما سيصنع الفارق الحقيقي 💪\n\n"
            f"يهمنا رأيك — كيف تقيم رضاك عن النتائج من 1 إلى 5؟{sat_line}{renewal_reminder}"
        )
    else:
        message = (
            f"{prefix}مرحبا {doctor_name}،\n\n"
            f"نشتغل بجد هالفترة على تحسين أداء عيادتك. عندنا خطة واضحة وإجراءات محددة:\n\n"
            f"🔧 ما نشتغل عليه الآن:\n{flag_section}\n\n"
            f"التزامنا معك مستمر، ورح نشاركك التحديثات أولاً بأول 💪\n\n"
            f"يهمنا رأيك دائماً — كيف تقيم رضاك عن تعاوننا؟{sat_line}{renewal_reminder}"
        )

    return message


# ── FUNCTION 6: check_already_sent_for_date ──────────────────────────────────

def _get_last_sent_value(contact_id: str, threeup_api_key: str) -> str:
    """Return the raw value of good_report_last_sent field, or '' on any error."""
    try:
        headers = {
            "Authorization": f"Bearer {threeup_api_key}",
            "Version": "2021-07-28",
        }
        r = requests.get(f"{GHL_BASE}/contacts/{contact_id}", headers=headers, timeout=30)
        if r.status_code != 200:
            return ""
        body = r.json()
        custom_fields = (
            body.get("contact", {}).get("customFields", [])
            or body.get("customFields", [])
        )
        for field in custom_fields:
            if field.get("id") == _GOOD_REPORT_LAST_SENT_ID:
                return str(field.get("value", "")).strip()
    except Exception:
        pass
    return ""


def check_already_sent_for_date(contact_id: str, api_key: str, due_date_str: str) -> bool:
    """
    Returns True if good_report_last_sent equals
    the specific due_date_str (YYYY-MM-DD).
    """
    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Version": "2021-07-28"
        }
        r = requests.get(
            f"https://services.leadconnectorhq.com/contacts/{contact_id}",
            headers=headers, timeout=30
        )
        if r.status_code != 200:
            return False
        fields = r.json().get("contact", {}).get("customFields", [])
        for f in fields:
            if f.get("id") == _GOOD_REPORT_LAST_SENT_ID:
                return str(f.get("value", "")).strip() == due_date_str
        return False
    except:
        return False


def _has_been_sent_for_date(contact_id: str, api_key: str, due_date_str: str) -> bool:
    """Returns True only if last_sent stamp matches this specific due_date."""
    try:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Version": "2021-07-28"
        }
        r = requests.get(
            f"https://services.leadconnectorhq.com/contacts/{contact_id}",
            headers=headers, timeout=30
        )
        if r.status_code != 200:
            return False
        fields = r.json().get("contact", {}).get("customFields", [])
        for f in fields:
            if f.get("id") == _GOOD_REPORT_LAST_SENT_ID:
                stored = str(f.get("value", "")).strip()
                return stored == due_date_str
        return False
    except:
        return False


# ── FUNCTION 7: write_to_ghl ──────────────────────────────────────────────────

def write_to_ghl(client: dict, message: str, settings: dict, message_type: str, due_date_str: str) -> bool:
    """Write good report message to GHL contact via custom field + tag."""
    threeup_api_key = settings.get("threeup_api_key", "")
    contact_id      = client.get("contact_id", "")
    name            = client.get("name", "")

    if not contact_id:
        logger.error(f"[good_report] {name} — contact_id is missing or empty, skipping")
        return False

    if not threeup_api_key:
        logger.info(f"[good_report] {name} — missing api_key")
        return False

    headers = {
        "Authorization": f"Bearer {threeup_api_key}",
        "Version": "2021-07-28",
        "Content-Type": "application/json",
    }
    contact_url = f"{GHL_BASE}/contacts/{contact_id}"

    # Write message to cao_good_report field by confirmed ID
    r = requests.put(
        contact_url, headers=headers,
        json={"customFields": [{"id": _CAO_GOOD_REPORT_FIELD_ID, "field_value": message}]},
        timeout=30,
    )
    if r.status_code not in (200, 201):
        logger.error(f"[good_report] {name} — field write failed: {r.status_code} {r.text[:200]}")
        return False

    # Add tag to trigger GHL workflow
    r2 = requests.post(
        f"{contact_url}/tags", headers=headers,
        json={"tags": ["good-report-ready"]},
        timeout=30,
    )
    print(f"[good_report] Tag POST status: {r2.status_code}")
    print(f"[good_report] Tag POST response: {r2.text[:300]}")
    if r2.status_code not in (200, 201):
        logger.error(f"[good_report] {name} — tag add failed: {r2.status_code} {r2.text[:200]}")
        return False

    # Stamp the due_date into good_report_last_sent by confirmed ID
    requests.put(
        contact_url, headers=headers,
        json={"customFields": [{"id": _GOOD_REPORT_LAST_SENT_ID, "field_value": due_date_str}]},
        timeout=30,
    )

    logger.info(f"[good_report] {name} ({message_type}) ✅")
    print(f"[good_report] {name} ({message_type}) ✅")
    return True


# ── FUNCTION 8a: queue_send ──────────────────────────────────────────────────

def queue_send(client_name: str, reason: str) -> None:
    """Append client_name to the deferred send queue (no-op if already present)."""
    try:
        queue = []
        if os.path.exists(_QUEUE_PATH):
            with open(_QUEUE_PATH, "r", encoding="utf-8") as f:
                queue = json.load(f)
    except Exception:
        queue = []

    if any(e["client_name"] == client_name for e in queue):
        print(f"[queue] {client_name} already queued — skipping duplicate")
        return

    queue.append({
        "client_name": client_name,
        "queued_at":   datetime.now(pytz.timezone("Asia/Amman")).isoformat(),
        "reason":      reason,
    })

    os.makedirs(os.path.dirname(_QUEUE_PATH), exist_ok=True)
    with open(_QUEUE_PATH, "w", encoding="utf-8") as f:
        json.dump(queue, f, indent=2, ensure_ascii=False)

    print(f"[queue] Added {client_name} (reason={reason})")


# ── FUNCTION 8b: process_queue ────────────────────────────────────────────────

def process_queue(clients: list, settings: dict) -> tuple:
    """Drain the deferred send queue within 9am–5pm Amman time.
    Uses the same cadence logic as run_force.
    Returns (sent_names: list[str], remaining_names: list[str])."""
    from sheets_manager import get_all_sheet_rows

    sent   = []
    remain = []

    if not os.path.exists(_QUEUE_PATH):
        return sent, remain

    try:
        with open(_QUEUE_PATH, "r", encoding="utf-8") as f:
            queue = json.load(f)
    except Exception as e:
        print(f"[queue] Failed to load queue: {e}")
        return sent, remain

    if not queue:
        return sent, remain

    amman_hour = datetime.now(pytz.timezone("Asia/Amman")).hour
    if not (9 <= amman_hour < 17):
        print(f"[queue] Outside business hours ({amman_hour}:xx) — {len(queue)} item(s) deferred")
        return sent, [e["client_name"] for e in queue]

    print(f"[queue] Processing {len(queue)} queued send(s)...")
    satisfaction_link = settings.get("satisfaction_link", "")
    today             = datetime.now(pytz.timezone("Asia/Amman")).date()

    try:
        all_tabs = load_all_am_tabs()
    except Exception as e:
        print(f"[queue] Failed to load AM tabs: {e}")
        return sent, [e["client_name"] for e in queue]

    failed_entries = []

    for i, entry in enumerate(queue):
        # Re-check hours before each send — drip may push past 5pm
        if not (9 <= datetime.now(pytz.timezone("Asia/Amman")).hour < 17):
            print(f"[queue] Exiting business hours — deferring remaining {len(queue) - i} item(s)")
            failed_entries.extend(queue[i:])
            remain.extend(e["client_name"] for e in queue[i:])
            break

        client_name = entry["client_name"]

        # Match client in clients.json
        client      = None
        force_lower = client_name.strip().lower()
        for c in clients:
            c_lower = c["name"].strip().lower()
            if force_lower == c_lower or force_lower in c_lower or c_lower in force_lower:
                client = c
                break

        if not client:
            print(f"[queue] {client_name} — no client record, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue

        if client.get("churned"):
            print(f"[queue] {client['name']} — churned, removing from queue")
            continue

        name        = client["name"]
        doctor_name = client.get("doctor_name", name)
        am_name     = client.get("am_name", name)

        anchor_str = client.get("cadence_anchor")
        if not anchor_str:
            print(f"[queue] {name} — no cadence_anchor, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue
        try:
            anchor = datetime.strptime(str(anchor_str), "%Y-%m-%d").date()
        except ValueError:
            print(f"[queue] {name} — bad cadence_anchor '{anchor_str}', keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue

        term = find_active_term(am_name, all_tabs, today)
        if not term:
            print(f"[queue] {name} — no active term in AM sheet, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue

        events      = _walk_cadence_events(anchor, all_tabs, am_name, today, lookback_days=7, horizon_days=0)
        cands       = [(d, t) for d, t in events if d <= today]
        if not cands:
            print(f"[queue] {name} — no due event, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue
        chosen, msg_type = max(cands, key=lambda x: x[0])
        due_date_str     = chosen.strftime("%Y-%m-%d")
        payment_date_str = term["payment_date"].strftime("%d/%m/%Y")
        renewal_date_str = term["renewal_date"].strftime("%d/%m/%Y")

        try:
            all_rows = get_all_sheet_rows(name)
        except Exception as e:
            print(f"[queue] {name} — sheet error: {e}, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue

        if not all_rows:
            print(f"[queue] {name} — no sheet data, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue

        performance = get_client_performance(name, all_rows, msg_type)
        if not performance:
            print(f"[queue] {name} — no performance data, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)
            continue

        message = generate_message(
            doctor_name, performance, msg_type,
            satisfaction_link, renewal_date_str, payment_date_str,
        )

        success = write_to_ghl(client, message, settings, msg_type, due_date_str)
        if success:
            sent.append(name)
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    today.strftime("%Y-%m-%d"),
                "client":         name,
                "am_name":        am_name,
                "contact_id":     client.get("contact_id", ""),
                "type":           msg_type,
                "status":         "sent_on_time" if due_date_str == today.strftime("%Y-%m-%d") else "sent_late",
                "triggered_by":   "queue_drain",
                "reason":         None,
            })
        else:
            print(f"[queue] {name} — GHL write failed, keeping in queue")
            failed_entries.append(entry)
            remain.append(client_name)

        if i < len(queue) - 1:
            print(f"[queue] Waiting 30 minutes before next queued send...")
            time.sleep(DRIP_DELAY_SECONDS)

    try:
        os.makedirs(os.path.dirname(_QUEUE_PATH), exist_ok=True)
        with open(_QUEUE_PATH, "w", encoding="utf-8") as f:
            json.dump(failed_entries, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[queue] Failed to save updated queue: {e}")

    print(f"[queue] Done — {len(sent)} sent, {len(remain)} remaining in queue")
    return sent, remain


# ── FUNCTION 8c: run ──────────────────────────────────────────────────────────

def run(clients: list, settings: dict) -> None:
    """Main entry point — find due messages, generate, and send."""
    from sheets_manager import get_all_sheet_rows

    _amman_now     = datetime.now(pytz.timezone("Asia/Amman"))
    amman_hour     = _amman_now.hour
    today          = _amman_now.date()
    today_str      = today.strftime("%Y-%m-%d")
    lookback_start = today - timedelta(days=7)

    print(f"[good_report] Running for {today_str}...")

    # ── Business-hours guard ──────────────────────────────────────────────────
    if amman_hour < 9 or amman_hour >= 19:
        all_tabs = load_all_am_tabs()
        due = get_due_messages(clients, all_tabs, today) if all_tabs else []
        for entry in due:
            queue_send(entry["client_name"], f"outside-hours ({amman_hour}:xx)")
            _rec = match_client(entry["client_name"], clients)
            _append_history({
                "scheduled_date": entry["due_date"].strftime("%Y-%m-%d"),
                "actual_date":    None,
                "client":         _rec["name"] if _rec else entry["client_name"],
                "am_name":        entry["client_name"],
                "contact_id":     _rec.get("contact_id", "") if _rec else "",
                "type":           entry["message_type"],
                "status":         "queued",
                "triggered_by":   "outside-hours",
                "reason":         f"outside-hours ({amman_hour}:xx)",
            })
        print(f"[run] Outside 9am-7pm Amman ({amman_hour}:xx) — {len(due)} due item(s) queued.")
        return

    # Drain overnight queue before computing today's due messages
    queue_sent, queue_remaining = process_queue(clients, settings)

    all_tabs = load_all_am_tabs()
    if not all_tabs:
        print("[good_report] No AM tabs loaded — check sheet access")
        return

    due = get_due_messages(clients, all_tabs, today)

    if not due:
        print("[good_report] No good reports due today")

    print(f"[good_report] {len(due)} message(s) due today")

    threeup_api_key   = settings.get("threeup_api_key", "")
    satisfaction_link = settings.get("satisfaction_link", "")
    sent_count  = 0
    skip_count  = 0
    error_count = 0
    sent_list   = []
    skip_list   = []
    error_list  = []
    sent_names  = set()
    processed_today = set()

    for i, entry in enumerate(due):
        am_name          = entry["client_name"]
        message_type     = entry["message_type"]
        due_date_str     = entry["due_date"].strftime("%Y-%m-%d")
        payment_date_str = entry.get("payment_date_str", "")
        renewal_date_str = entry.get("renewal_date_str", "")
        is_missed        = entry.get("is_missed", False)

        # In-run deduplication guard
        if am_name in processed_today:
            print(f"[good_report] {am_name} — DUPLICATE BLOCKED by processed_today")
            skip_count += 1
            skip_list.append(f"{am_name} (duplicate in run)")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         am_name,
                "am_name":        am_name,
                "contact_id":     "",
                "type":           message_type,
                "status":         "skipped_duplicate",
                "triggered_by":   "scheduled",
                "reason":         "duplicate in run",
            })
            continue
        processed_today.add(am_name)

        client = match_client(am_name, clients)
        if not client:
            print(f"[good_report] {am_name} — no client record, skipping")
            skip_count += 1
            skip_list.append(f"{am_name} (no client record)")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         am_name,
                "am_name":        am_name,
                "contact_id":     "",
                "type":           message_type,
                "status":         "missed",
                "triggered_by":   "scheduled",
                "reason":         "no_client_record",
            })
            continue

        name        = client["name"]
        contact_id  = client.get("contact_id", "")
        doctor_name = client.get("doctor_name", name)

        try:
            all_rows = get_all_sheet_rows(name)
        except Exception as e:
            print(f"[good_report] {name} — sheet error: {e}")
            error_count += 1
            error_list.append(f"{name} (sheet error: {e})")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         name,
                "am_name":        am_name,
                "contact_id":     contact_id,
                "type":           message_type,
                "status":         "missed",
                "triggered_by":   "scheduled",
                "reason":         f"sheet_error: {e}",
            })
            continue

        if not all_rows:
            print(f"[good_report] {name} — no data yet, skipping")
            skip_count += 1
            skip_list.append(f"{name} (no sheet data)")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         name,
                "am_name":        am_name,
                "contact_id":     contact_id,
                "type":           message_type,
                "status":         "missed",
                "triggered_by":   "scheduled",
                "reason":         "no_sheet_data",
            })
            continue

        # Both missed and today's entries: skip only if sent for this specific due_date
        if _has_been_sent_for_date(contact_id, threeup_api_key, due_date_str):
            print(f"[good_report] {name} — already sent for {due_date_str}, skipping")
            skip_count += 1
            skip_list.append(f"{name} (already sent for {due_date_str})")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         name,
                "am_name":        am_name,
                "contact_id":     contact_id,
                "type":           message_type,
                "status":         "skipped_duplicate",
                "triggered_by":   "scheduled",
                "reason":         f"already_sent_for_{due_date_str}",
            })
            continue

        performance = get_client_performance(name, all_rows, message_type)
        if not performance:
            print(f"[good_report] {name} — no performance data, skipping")
            skip_count += 1
            skip_list.append(f"{name} (no performance data)")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         name,
                "am_name":        am_name,
                "contact_id":     contact_id,
                "type":           message_type,
                "status":         "missed",
                "triggered_by":   "scheduled",
                "reason":         "no_performance_data",
            })
            continue

        message = generate_message(
            doctor_name, performance, message_type,
            satisfaction_link, renewal_date_str, payment_date_str,
        )

        success = write_to_ghl(client, message, settings, message_type, due_date_str)
        if success:
            sent_count += 1
            sent_list.append(f"{name} ({message_type}{'  [missed]' if is_missed else ''})")
            sent_names.add(name)
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    today_str,
                "client":         name,
                "am_name":        am_name,
                "contact_id":     contact_id,
                "type":           message_type,
                "status":         "sent_on_time" if due_date_str == today_str else "sent_late",
                "triggered_by":   "scheduled",
                "reason":         None,
            })
        else:
            error_count += 1
            error_list.append(f"{name} (GHL write failed)")
            _append_history({
                "scheduled_date": due_date_str,
                "actual_date":    None,
                "client":         name,
                "am_name":        am_name,
                "contact_id":     contact_id,
                "type":           message_type,
                "status":         "missed",
                "triggered_by":   "scheduled",
                "reason":         "ghl_write_failed",
            })

        if i < len(due) - 1:
            print(f"[good_report] Waiting 30 minutes before next message...")
            time.sleep(DRIP_DELAY_SECONDS)

    print(f"[good_report] Done — {sent_count} sent, {skip_count} skipped, {error_count} errors")

    # Build (client_name, scheduled_date) pairs already confirmed sent in history.
    # Avoids false-alarming clients that sent on a prior day within the lookback window.
    _hist_sent = set()
    try:
        if os.path.exists(_HISTORY_PATH):
            with open(_HISTORY_PATH, "r", encoding="utf-8") as _hf:
                for _he in json.load(_hf):
                    if _he.get("status") in ("sent_on_time", "sent_late") \
                            and _he.get("client") and _he.get("scheduled_date"):
                        _hist_sent.add((_he["client"], _he["scheduled_date"]))
    except Exception:
        pass

    all_sent_names = sent_names | set(queue_sent)
    missed_list    = []
    for _client in clients:
        if _client.get("churned"):
            continue
        _cname = _client["name"]
        if _cname in all_sent_names:
            continue
        _anchor_str = _client.get("cadence_anchor")
        if not _anchor_str:
            continue
        try:
            _anchor  = datetime.strptime(str(_anchor_str), "%Y-%m-%d").date()
            _am_name = _client.get("am_name", _client["name"])
            _events  = _walk_cadence_events(_anchor, all_tabs, _am_name, today,
                                            lookback_days=7, horizon_days=0)
            for _dt, _tp in _events:
                if lookback_start <= _dt < today:
                    if (_cname, _dt.strftime("%Y-%m-%d")) in _hist_sent:
                        break  # already sent for this due date — not missed
                    missed_list.append(
                        f"  • {_cname} — {_tp} {_dt} ({(today - _dt).days}d ago)"
                    )
                    break
        except Exception:
            continue

    sent_block   = "\n".join(f"  • {s}" for s in sent_list)       or "  None"
    queued_block = "\n".join(f"  • {n}" for n in queue_remaining)  or "  None"
    missed_block = "\n".join(missed_list)                           or "  None"
    error_block  = "\n".join(f"  • {s}" for s in error_list)       or "  None"

    from system_health import health_block, record_run
    record_run("good_report")
    send_email(
        f"ThreeUp Daily Report — {today_str} — {sent_count} sent",
        health_block(settings) +
        f"ThreeUp Daily Report — {today_str}\n\n"
        f"SENT TODAY ({sent_count}):\n{sent_block}\n\n"
        f"QUEUED FOR NEXT 9AM ({len(queue_remaining)}):\n{queued_block}\n\n"
        f"MISSED — due in last 7 days, not sent ({len(missed_list)}):\n{missed_block}\n\n"
        f"ERRORS ({error_count}):\n{error_block}\n"
    )

    write_schedule_to_sheet(clients, settings)
    write_history_to_sheet()
    write_future_events_to_sheet(clients)


# ── FUNCTION 9: run_force ─────────────────────────────────────────────────────

def run_force(clients: list, settings: dict, client_name: str, defer: bool = False) -> None:
    """Force-send a good report for a single named client, bypassing all sent-checks.
    If defer=True or outside 9am–5pm Amman, queues instead of sending immediately."""
    from sheets_manager import get_all_sheet_rows

    amman_now = datetime.now(pytz.timezone("Asia/Amman"))
    today     = amman_now.date()
    today_str = today.strftime("%Y-%m-%d")
    satisfaction_link = settings.get("satisfaction_link", "")

    if defer or not (9 <= amman_now.hour < 17):
        queue_send(client_name, "force")
        print(f"[good_report_force] Outside hours ({amman_now.hour}:xx) — queued for 9am.")
        return

    print(f"[good_report_force] Forcing report for: '{client_name}' on {today_str}")

    all_tabs = load_all_am_tabs()
    if not all_tabs:
        print("[good_report_force] No AM tabs loaded")
        return

    # Find client in clients.json
    client      = None
    force_lower = client_name.strip().lower()
    for c in clients:
        c_lower = c["name"].strip().lower()
        if force_lower == c_lower or force_lower in c_lower or c_lower in force_lower:
            client = c
            break
    if not client:
        client = match_client(client_name, clients)
    if not client:
        print(f"[good_report_force] No client record for '{client_name}'")
        return
    if client.get("churned"):
        print(f"[good_report_force] {client['name']} is churned — skipping")
        return

    anchor_str = client.get("cadence_anchor")
    if not anchor_str:
        print(f"[good_report_force] No cadence_anchor for '{client['name']}' — cannot compute schedule")
        return

    name        = client["name"]
    doctor_name = client.get("doctor_name", name)
    am_name     = client.get("am_name", name)
    anchor      = datetime.strptime(str(anchor_str), "%Y-%m-%d").date()

    # Get active term for display strings (renewal/payment not used for cadence math)
    term = find_active_term(am_name, all_tabs, today)
    if not term:
        print(f"[good_report_force] No active term in AM sheet for '{name}'")
        return
    payment_date_str = term["payment_date"].strftime("%d/%m/%Y")
    renewal_date_str = term["renewal_date"].strftime("%d/%m/%Y")

    # Walk cadence to find the most recent event due today or in the last 7 days
    events     = _walk_cadence_events(anchor, all_tabs, am_name, today, lookback_days=7, horizon_days=0)
    candidates = [(d, t) for d, t in events if d <= today]
    if not candidates:
        print(f"[good_report_force] No due event found for '{name}' (anchor={anchor})")
        return
    chosen, msg_type = max(candidates, key=lambda x: x[0])
    due_date_str     = chosen.strftime("%Y-%m-%d")

    print(f"[good_report_force] Client: {name} | doctor: {doctor_name}")
    print(f"[good_report_force] anchor={anchor} term_payment={term['payment_date']} chosen={chosen} type={msg_type}")

    try:
        all_rows = get_all_sheet_rows(name)
    except Exception as e:
        print(f"[good_report_force] Sheet error: {e}")
        return

    if not all_rows:
        print(f"[good_report_force] No sheet data for {name}")
        return

    performance = get_client_performance(name, all_rows, msg_type)
    if not performance:
        print(f"[good_report_force] No performance data for {name}")
        return

    message = generate_message(
        doctor_name, performance, msg_type,
        satisfaction_link, renewal_date_str, payment_date_str,
    )

    print(f"[good_report_force] Generated message ({len(message)} chars):")
    print(message)
    print()

    success = write_to_ghl(client, message, settings, msg_type, due_date_str)
    if success:
        _append_history({
            "scheduled_date": due_date_str,
            "actual_date":    today_str,
            "client":         name,
            "am_name":        am_name,
            "contact_id":     client.get("contact_id", ""),
            "type":           msg_type,
            "status":         "forced",
            "triggered_by":   "force",
            "reason":         None,
        })


# ── FUNCTION 10: get_full_schedule ────────────────────────────────────────────

def get_full_schedule(clients: list, settings: dict) -> list:
    """Return schedule data for all non-churned clients for dashboard display."""
    threeup_api_key = settings.get("threeup_api_key", "")
    today    = datetime.now(pytz.timezone("Asia/Amman")).date()
    all_tabs = load_all_am_tabs()

    rows = []
    for client in clients:
        if client.get("churned"):
            continue

        name        = client["name"]
        doctor_name = client.get("doctor_name", "")
        contact_id  = client.get("contact_id", "")
        am_name     = client.get("am_name", name)

        last_sent = ""
        strikes   = 0
        if contact_id and threeup_api_key:
            try:
                headers = {
                    "Authorization": f"Bearer {threeup_api_key}",
                    "Version": "2021-07-28",
                }
                r = requests.get(f"{GHL_BASE}/contacts/{contact_id}", headers=headers, timeout=30)
                if r.status_code == 200:
                    cf = r.json().get("contact", {}).get("customFields", [])
                    for field in cf:
                        if field.get("id") == _GOOD_REPORT_LAST_SENT_ID:
                            last_sent = str(field.get("value", "")).strip()
                        if field.get("id") == "nNx5vev4O2dBgLbqYNSh":
                            try:
                                strikes = int(field.get("value", 0) or 0)
                            except Exception:
                                strikes = 0
            except Exception:
                pass

        next_send  = None
        msg_type   = ""
        anchor_str = client.get("cadence_anchor")
        if anchor_str:
            try:
                anchor    = datetime.strptime(str(anchor_str), "%Y-%m-%d").date()
                all_dates = _walk_cadence_events(anchor, all_tabs, am_name, today,
                                                 lookback_days=0, horizon_days=90)
                future = [(d, t) for d, t in all_dates if d >= today]
                if future:
                    next_send, msg_type = min(future, key=lambda x: x[0])
                else:
                    past = [(d, t) for d, t in all_dates if d < today]
                    if past:
                        next_send, msg_type = max(past, key=lambda x: x[0])
            except Exception:
                pass

        days = (next_send - today).days if next_send else None

        if days is None:
            status = "NO SCHEDULE"
        elif days == 0:
            status = "DUE TODAY"
        elif days > 0:
            status = "UPCOMING"
        else:
            status = "OVERDUE"

        rows.append({
            "client_name": name,
            "doctor_name": doctor_name,
            "strikes":     strikes,
            "last_sent":   last_sent,
            "next_send":   next_send.strftime("%Y-%m-%d") if next_send else "",
            "type":        msg_type,
            "days":        days,
            "due_today":   days == 0,
            "status":      status,
        })

    return sorted(rows, key=lambda x: (x["days"] is None, x["days"] if x["days"] is not None else 9999))


# ── FUNCTION 11: write_schedule_to_sheet ─────────────────────────────────────

def write_schedule_to_sheet(clients: list, settings: dict) -> None:
    """Write schedule data to 'CAO Schedule' tab in AM sheet for dashboard consumption."""
    try:
        schedule = get_full_schedule(clients, settings)
        if not schedule:
            print("[good_report] No schedule data to write")
            return

        print(f"[good_report] write_schedule_to_sheet: {len(schedule)} entries, sheet={AM_DATES_SHEET_ID}")
        service = _get_sheets_service()
        today   = datetime.now(pytz.timezone("Asia/Amman")).date()

        # Create tab if it doesn't exist
        print("[good_report] Fetching spreadsheet metadata...")
        meta            = service.spreadsheets().get(spreadsheetId=AM_DATES_SHEET_ID).execute()
        existing_titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
        print(f"[good_report] Existing tabs: {existing_titles}")
        if "CAO Schedule" not in existing_titles:
            print("[good_report] 'CAO Schedule' tab missing — creating...")
            service.spreadsheets().batchUpdate(
                spreadsheetId=AM_DATES_SHEET_ID,
                body={"requests": [{"addSheet": {"properties": {"title": "CAO Schedule"}}}]},
            ).execute()
            print("[good_report] Created 'CAO Schedule' tab")
        else:
            print("[good_report] 'CAO Schedule' tab already exists")

        header = ["Client", "Doctor", "Strikes", "Last Sent", "Next Send", "Type", "Days", "Due Today", "Updated"]
        rows   = [header]
        for s in schedule:
            rows.append([
                s["client_name"],
                s.get("doctor_name", ""),
                s["strikes"],
                s["last_sent"],
                s["next_send"],
                s["type"],
                str(s["days"]) if s["days"] is not None else "",
                "TRUE" if s.get("due_today") else "FALSE",
                today.strftime("%Y-%m-%d"),
            ])

        print("[good_report] Clearing 'CAO Schedule'!A:I ...")
        service.spreadsheets().values().clear(
            spreadsheetId=AM_DATES_SHEET_ID,
            range="'CAO Schedule'!A:I",
        ).execute()
        print("[good_report] Clear done. Writing values...")

        service.spreadsheets().values().update(
            spreadsheetId=AM_DATES_SHEET_ID,
            range="'CAO Schedule'!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()

        print(f"[good_report] CAO Schedule written: {len(schedule)} rows")
    except Exception as e:
        print(f"[good_report] write_schedule_to_sheet failed: {e}")
        import traceback
        traceback.print_exc()


# ── FUNCTION 12: write_history_to_sheet ──────────────────────────────────────

def write_history_to_sheet() -> None:
    """Write send_history.json to 'CAO History' tab in AM sheet for dashboard consumption."""
    try:
        if not os.path.exists(_HISTORY_PATH):
            print("[good_report] write_history_to_sheet: no history file")
            return
        with open(_HISTORY_PATH, "r", encoding="utf-8") as f:
            history = json.load(f)
        if not history:
            print("[good_report] write_history_to_sheet: history empty")
            return

        service = _get_sheets_service()
        updated = datetime.now(pytz.timezone("Asia/Amman")).strftime("%Y-%m-%d %H:%M")

        meta   = service.spreadsheets().get(spreadsheetId=AM_DATES_SHEET_ID).execute()
        titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if "CAO History" not in titles:
            service.spreadsheets().batchUpdate(
                spreadsheetId=AM_DATES_SHEET_ID,
                body={"requests": [{"addSheet": {"properties": {"title": "CAO History"}}}]},
            ).execute()

        header = ["Scheduled", "Actual", "Client", "AM Name", "Type",
                  "Status", "Triggered By", "Reason", "Updated"]
        rows = [header]
        for h in history:
            rows.append([
                h.get("scheduled_date") or "",
                h.get("actual_date")    or "",
                h.get("client",       ""),
                h.get("am_name",      ""),
                h.get("type",         ""),
                h.get("status",       ""),
                h.get("triggered_by", ""),
                h.get("reason")        or "",
                updated,
            ])

        service.spreadsheets().values().clear(
            spreadsheetId=AM_DATES_SHEET_ID, range="'CAO History'!A:I"
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=AM_DATES_SHEET_ID,
            range="'CAO History'!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        print(f"[good_report] CAO History written: {len(history)} rows")
    except Exception as e:
        print(f"[good_report] write_history_to_sheet failed: {e}")
        import traceback; traceback.print_exc()


# ── FUNCTION 13: write_future_events_to_sheet ─────────────────────────────────

def write_future_events_to_sheet(clients: list) -> None:
    """Write all forward cadence events per client to 'CAO Future Events' tab.
    One row per event. Walks up to 730 days; stops naturally when sheet has no more term data.
    Anchor-less clients emit only their upcoming renewal reminder from the sheet."""
    try:
        today    = datetime.now(pytz.timezone("Asia/Amman")).date()
        all_tabs = load_all_am_tabs()
        if not all_tabs:
            print("[good_report] write_future_events_to_sheet: no AM tabs loaded")
            return

        service = _get_sheets_service()
        updated = datetime.now(pytz.timezone("Asia/Amman")).strftime("%Y-%m-%d %H:%M")

        meta   = service.spreadsheets().get(spreadsheetId=AM_DATES_SHEET_ID).execute()
        titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
        if "CAO Future Events" not in titles:
            service.spreadsheets().batchUpdate(
                spreadsheetId=AM_DATES_SHEET_ID,
                body={"requests": [{"addSheet": {"properties": {"title": "CAO Future Events"}}}]},
            ).execute()

        header = ["Client", "Doctor", "AM Name", "Event Date", "Type", "Updated"]
        rows = [header]

        for client in clients:
            if client.get("churned"):
                continue

            name        = client["name"]
            doctor_name = client.get("doctor_name", "")
            am_name     = client.get("am_name", name)
            anchor_str  = client.get("cadence_anchor")

            if not anchor_str:
                # No cadence anchor — emit only future renewal reminder from sheet
                term = find_active_term(am_name, all_tabs, today)
                if term:
                    renewal_reminder = term["payment_date"] - timedelta(days=6)
                    if renewal_reminder >= today:
                        rows.append([name, doctor_name, am_name,
                                     renewal_reminder.strftime("%Y-%m-%d"), "renewal", updated])
                continue

            try:
                anchor = datetime.strptime(str(anchor_str), "%Y-%m-%d").date()
            except ValueError:
                continue

            # Walk up to 730 days — naturally terminates when no next-term exists in sheet
            events = _walk_cadence_events(anchor, all_tabs, am_name, today,
                                          lookback_days=0, horizon_days=730)
            for event_date, event_type in events:
                if event_date >= today:
                    rows.append([name, doctor_name, am_name,
                                 event_date.strftime("%Y-%m-%d"), event_type, updated])

        service.spreadsheets().values().clear(
            spreadsheetId=AM_DATES_SHEET_ID, range="'CAO Future Events'!A:F"
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=AM_DATES_SHEET_ID,
            range="'CAO Future Events'!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        print(f"[good_report] CAO Future Events written: {len(rows)-1} events")
    except Exception as e:
        print(f"[good_report] write_future_events_to_sheet failed: {e}")
        import traceback; traceback.print_exc()


# ── FUNCTION 14: dry_audit ───────────────────────────────────────────────────

def dry_audit(clients: list) -> None:
    """Read-only audit: show each client's anchor, computed event sequence (60 days),
    and on-track status. No GHL calls, no writes."""
    today    = datetime.now(pytz.timezone("Asia/Amman")).date()
    all_tabs = load_all_am_tabs()

    print(f"\n{'='*90}")
    print(f"DRY AUDIT — cadence walk (anchor-based) — today={today}")
    print(f"{'='*90}\n")

    for client in sorted(clients, key=lambda c: c["name"]):
        if client.get("churned"):
            continue
        name       = client["name"]
        am_name    = client.get("am_name", name)
        anchor_str = client.get("cadence_anchor")

        if not anchor_str:
            print(f"{'─'*60}")
            print(f"  {name}  [NO ANCHOR — skipped]")
            continue

        anchor = datetime.strptime(str(anchor_str), "%Y-%m-%d").date()

        # Active term from sheet (renewal/payment reference only)
        term = find_active_term(am_name, all_tabs, today)
        term_str = (f"payment={term['payment_date']}  renewal={term['renewal_date']}"
                    if term else "NO TERM IN SHEET")

        # Walk: past 14d + next 60d so we see recent + upcoming
        events = _walk_cadence_events(anchor, all_tabs, am_name, today,
                                      lookback_days=14, horizon_days=60)

        # Determine on-track: is there an event today, or was the most recent missed?
        past_events   = [(d, t) for d, t in events if d < today]
        today_events  = [(d, t) for d, t in events if d == today]
        future_events = [(d, t) for d, t in events if d > today]

        if today_events:
            status = "DUE TODAY"
        elif past_events:
            last_d, last_t = max(past_events, key=lambda x: x[0])
            days_ago = (today - last_d).days
            status = f"OVERDUE ({days_ago}d ago: {last_t} {last_d})"
        elif future_events:
            next_d, next_t = min(future_events, key=lambda x: x[0])
            days_ahead = (next_d - today).days
            status = f"ok — next in {days_ahead}d"
        else:
            status = "NO FUTURE EVENTS — June tab needs update"

        print(f"{'─'*60}")
        print(f"  {name}")
        print(f"  anchor={anchor}  am_name={am_name}")
        print(f"  sheet:  {term_str}")
        if events:
            seq = "  ".join(
                f"[{'TODAY' if d==today else ('PAST' if d<today else f'+{(d-today).days}d')}] {d} ({t[:3].upper()})"
                for d, t in events
            )
            print(f"  events: {seq}")
        else:
            print(f"  events: (none in window — walk stopped at last known term)")
        print(f"  status: {status}")

    print(f"\n{'='*90}\n")


# ── FUNCTION 13: validate_tabs ────────────────────────────────────────────────

def validate_tabs() -> None:
    """Read all month tabs and flag data quality issues for AM review."""
    today        = datetime.now(pytz.timezone("Asia/Amman")).date()
    current_year = today.year

    try:
        service     = _get_sheets_service()
        spreadsheet = service.spreadsheets().get(spreadsheetId=AM_DATES_SHEET_ID).execute()
        available   = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]
        month_tabs  = sorted(
            [t for t in available if _tab_to_month(t)],
            key=lambda t: _MONTH_ORDER[_tab_to_month(t)],
        )

        if not month_tabs:
            print("[validate] No month tabs found")
            return

        print(f"[validate] Scanning tabs: {month_tabs}\n")

        issues       = []
        prev_entries = {}  # client_name_lower -> entry from previous tab

        for tab_name in month_tabs:
            raw_data = service.spreadsheets().values().get(
                spreadsheetId=AM_DATES_SHEET_ID,
                range=f"'{tab_name}'!A:F",
            ).execute()
            raw_rows   = raw_data.get("values", [])
            tab_entries = {}

            for i, row in enumerate(raw_rows):
                if i == 0:
                    continue
                if not row:
                    continue
                name = str(row[0]).strip()
                if not name:
                    continue

                sus_raw     = str(row[1]).strip() if len(row) > 1 else ""
                renewal_raw = str(row[2]).strip() if len(row) > 2 else ""
                payment_raw = str(row[3]).strip() if len(row) > 3 else ""

                if sus_raw.lower() in ("", "waiting", "pause"):
                    continue  # expected placeholder — not a data error

                sus_date     = _parse_date(sus_raw)
                renewal_date = _parse_date(renewal_raw)
                payment_date = _parse_date(payment_raw)

                if not sus_date or not renewal_date or not payment_date:
                    issues.append(
                        f"[{tab_name}] {name}: unparseable date — "
                        f"sus='{sus_raw}' renewal='{renewal_raw}' payment='{payment_raw}'"
                    )
                    continue

                name_lower = name.strip().lower()

                # Check 1: renewal before sustainability
                if renewal_date < sus_date:
                    issues.append(
                        f"[{tab_name}] {name}: renewal ({renewal_date}) is BEFORE "
                        f"sustainability ({sus_date})"
                    )

                # Check 2: payment or renewal year != current year
                if payment_date.year != current_year:
                    issues.append(
                        f"[{tab_name}] {name}: payment year={payment_date.year} "
                        f"(expected {current_year}) — raw='{payment_raw}'"
                    )
                if renewal_date.year != current_year:
                    issues.append(
                        f"[{tab_name}] {name}: renewal year={renewal_date.year} "
                        f"(expected {current_year}) — raw='{renewal_raw}'"
                    )

                # Check 3: dates copied unchanged from previous tab
                if name_lower in prev_entries:
                    prev  = prev_entries[name_lower]
                    copied = []
                    if prev["first_sustainability_date"] == sus_date:
                        copied.append(f"sustainability ({sus_date})")
                    if prev["renewal_date"] == renewal_date:
                        copied.append(f"renewal ({renewal_date})")
                    if prev["payment_date"] == payment_date:
                        copied.append(f"payment ({payment_date})")
                    if copied:
                        issues.append(
                            f"[{tab_name}] {name}: "
                            + ", ".join(copied)
                            + " — identical to previous tab (possible copy-paste)"
                        )

                tab_entries[name_lower] = {
                    "client_name":               name,
                    "first_sustainability_date": sus_date,
                    "renewal_date":              renewal_date,
                    "payment_date":              payment_date,
                }

            prev_entries = tab_entries
            print(f"  Tab '{tab_name}': {len(tab_entries)} valid rows")

        print()
        if issues:
            print(f"[validate] {len(issues)} issue(s) found:\n")
            for iss in issues:
                print(f"  ⚠  {iss}")
        else:
            print("[validate] All tabs clean — no issues found")

    except Exception as e:
        print(f"[validate] Error: {e}")
        import traceback
        traceback.print_exc()
