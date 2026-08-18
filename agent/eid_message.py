import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import logging
import time
import requests

from good_report_engine import DRIP_DELAY_SECONDS, _check_and_mark_throttle

logger = logging.getLogger("eid_message")

GHL_BASE = "https://services.leadconnectorhq.com"
_CAO_GOOD_REPORT_FIELD_ID = "Y1lUr9X7ACLNmfsXmsCX"

EID_CLIENTS = [
    "Bella Clinic",
    "Rawnaq Medical Clinic",
    "Nova Dental Care",
    "Dr Karam Almishry",
    "Ahmed Alhouli",
    "Row Clinic",
    "Dr Heba Hammad Clinic",
    "WS Clinic Dental Clinic",
    "Dentaface",
    "Vogue Dental Clinic",
    "Dr Ashraf Dababneh",
    "Salamati Medical Center",
]


# ── FUNCTION 1: get_two_month_performance ─────────────────────────────────────

def _is_confirmed_row(row: list) -> bool:
    if len(row) <= 16:
        return False
    return any(str(row[i]).strip() != "" for i in [10, 11, 16])


def get_two_month_performance(client_name: str) -> dict:
    default = {
        "total_new_reviews": 0,
        "total_new_patients": 0,
        "avg_attendance": 0.0,
        "total_website_patients": 0,
        "clean_weeks": 0,
    }
    try:
        from sheets_manager import get_all_sheet_rows
        all_rows = get_all_sheet_rows(client_name)
        if not all_rows:
            return default

        confirmed = [r for r in all_rows if _is_confirmed_row(r)]
        if not confirmed:
            return default

        rows = confirmed[-8:]

        def _safe(row, idx):
            try:
                if len(row) <= idx:
                    return 0.0
                v = row[idx]
                if v in ("", None):
                    return 0.0
                return float(str(v).replace("%", "").strip())
            except (ValueError, TypeError):
                return 0.0

        total_new_reviews = sum(_safe(r, 1) for r in rows)
        total_new_patients = sum(_safe(r, 3) for r in rows)
        total_website_patients = sum(_safe(r, 31) for r in rows)

        att_values = []
        for r in rows:
            if len(r) > 15 and str(r[15]).strip() not in ("", None):
                try:
                    att_values.append((1 - float(str(r[15]).replace("%", "").strip())) * 100)
                except (ValueError, TypeError):
                    pass
        avg_attendance = sum(att_values) / len(att_values) if att_values else 0.0

        clean_weeks = sum(
            1 for r in rows
            if (len(r) > 16
                and str(r[10]).strip().upper() != "TRUE"
                and str(r[11]).strip().upper() != "TRUE"
                and str(r[16]).strip().upper() != "TRUE")
        )

        return {
            "total_new_reviews": total_new_reviews,
            "total_new_patients": total_new_patients,
            "avg_attendance": avg_attendance,
            "total_website_patients": total_website_patients,
            "clean_weeks": clean_weeks,
        }
    except Exception as e:
        logger.warning(f"[eid_message] get_two_month_performance({client_name}): {e}")
        return default


# ── FUNCTION 2: generate_eid_message ─────────────────────────────────────────

def generate_eid_message(client: dict, performance: dict) -> str:
    doctor_name = client.get("doctor_name", client["name"])
    nr  = int(performance["total_new_reviews"])
    np_ = int(performance["total_new_patients"])
    att = round(performance["avg_attendance"], 1)
    cw  = performance["clean_weeks"]
    wp  = int(performance["total_website_patients"])

    data_lines = []
    if nr  > 0: data_lines.append(f"⭐ تقييمات جوجل: حصلنا على {nr} تقييم جديد خلال الفترة الماضية")
    if np_ > 0: data_lines.append(f"👥 المرضى الجدد: استقبلنا {np_} مريض جديد")
    if att > 0: data_lines.append(f"📅 معدل الحضور: {att}٪ من المواعيد تمت بنجاح")
    if wp  > 0: data_lines.append(f"🌐 مرضى الموقع: {wp} مريض جديد عبر الموقع الإلكتروني")
    if cw  > 0: data_lines.append(f"✅ أسابيع بدون مشاكل: {cw} أسبوع نظيف")

    data_section = "\n".join(data_lines) if data_lines else "نعمل باستمرار على تحسين نتائجك"

    message = f"""كلّ عام وأنتم بخير من عائلة ThreeUp 🌙

{doctor_name}، نتمنى لكم ولعائلتكم عيداً سعيداً مباركاً.

نريد أن نشكركم على ثقتكم ودعمكم المتواصل — هذا يعني لنا الكثير 💙

📊 أبرز ما حققناه معاً خلال الشهرين الماضيين:
{data_section}

🔧 ما نعمل عليه قادماً:
- إطلاق نظام المحاسبة الإلكتروني لإدارة أفضل
- ملف المرضى الرقمي لتجربة أكثر احترافية
- تطوير مستمر في استراتيجيات الاستقطاب والتقييمات

━━━━━━━━━━━━━━━━━━━━
🗓️ تنبيه مهم: تم تأجيل دفعة برنامج ThreeUp إلى يوم الأحد الموافق ٣١/٥/٢٠٢٦، وسيتم تذكيركم في وقتها.

القادم أفضل بإذن الله — ونتمنى لكم موسماً صيفياً ناجحاً. يد بيد دائماً في النجاحات والعقبات 🤝

الدفع عبر كليك (Cliq): 00962796876276"""

    return message


# ── FUNCTION 3: generate_payment_reminder ────────────────────────────────────

def generate_payment_reminder(client: dict) -> str:
    doctor_name = client.get("doctor_name", client["name"])

    message = f"""مرحباً {doctor_name} 👋

هذا تذكير بأن موعد دفعة برنامج ThreeUp هو اليوم الأحد ٣١/٥/٢٠٢٦ 🗓️

نشكركم على التزامكم ودعمكم المتواصل 💙

الدفع عبر كليك (Cliq): 00962796876276

في حال وجود أي استفسار، لا تترددوا بالتواصل معنا."""

    return message


# ── FUNCTION 4: send_to_ghl ───────────────────────────────────────────────────

def send_to_ghl(client: dict, message: str, settings: dict) -> bool:
    threeup_api_key = settings.get("threeup_api_key", "")
    contact_id      = client.get("contact_id", "")
    name            = client.get("name", "")

    if not contact_id:
        logger.error(f"[eid_message] {name} — contact_id missing, skipping")
        print(f"[eid_message] {name} — contact_id missing, skipping")
        return False

    if not threeup_api_key:
        logger.error(f"[eid_message] {name} — threeup_api_key missing")
        print(f"[eid_message] {name} — threeup_api_key missing")
        return False

    # Universal 30-min throttle (same enforcement point as good_report_engine.write_to_ghl)
    _check_and_mark_throttle(name)

    headers = {
        "Authorization": f"Bearer {threeup_api_key}",
        "Version": "2021-07-28",
        "Content-Type": "application/json",
    }
    contact_url = f"{GHL_BASE}/contacts/{contact_id}"

    r = requests.put(
        contact_url, headers=headers,
        json={"customFields": [{"id": _CAO_GOOD_REPORT_FIELD_ID, "field_value": message}]},
        timeout=30,
    )
    if r.status_code not in (200, 201):
        logger.error(f"[eid_message] {name} — field write failed: {r.status_code} {r.text[:200]}")
        print(f"[eid_message] {name} — field write failed: {r.status_code}")
        return False

    r2 = requests.post(
        f"{contact_url}/tags", headers=headers,
        json={"tags": ["good-report-ready"]},
        timeout=30,
    )
    print(f"[eid_message] {name} — tag POST {r2.status_code}: {r2.text[:200]}")
    if r2.status_code not in (200, 201):
        logger.error(f"[eid_message] {name} — tag add failed: {r2.status_code} {r2.text[:200]}")
        return False

    logger.info(f"[eid_message] {name} ✅")
    print(f"[eid_message] {name} ✅")
    return True


# ── FUNCTION 5: run_eid ───────────────────────────────────────────────────────

def run_eid(clients: list, settings: dict) -> None:
    due = [c for c in clients if c["name"] in EID_CLIENTS]
    if not due:
        print("[eid_message] No matching Eid clients found")
        return

    print(f"[eid_message] Sending Eid messages to {len(due)} client(s)...")
    sent, failed = 0, 0

    for i, client in enumerate(due):
        name        = client["name"]
        performance = get_two_month_performance(name)
        message     = generate_eid_message(client, performance)
        success     = send_to_ghl(client, message, settings)
        if success:
            sent += 1
        else:
            failed += 1

        if i < len(due) - 1:
            print(f"[eid_message] Waiting 30 min before next client...")
            time.sleep(DRIP_DELAY_SECONDS)

    print(f"[eid_message] Done — {sent} sent, {failed} failed")


# ── FUNCTION 6: run_payment_reminder ─────────────────────────────────────────

def run_payment_reminder(clients: list, settings: dict) -> None:
    due = [c for c in clients if c["name"] in EID_CLIENTS]
    if not due:
        print("[eid_message] No matching clients found for payment reminder")
        return

    print(f"[eid_message] Sending payment reminders to {len(due)} client(s)...")
    sent, failed = 0, 0

    for i, client in enumerate(due):
        message = generate_payment_reminder(client)
        success = send_to_ghl(client, message, settings)
        if success:
            sent += 1
        else:
            failed += 1

        if i < len(due) - 1:
            print(f"[eid_message] Waiting 30 min before next client...")
            time.sleep(DRIP_DELAY_SECONDS)

    print(f"[eid_message] Done — {sent} sent, {failed} failed")
