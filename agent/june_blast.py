import sys
import io
if getattr(sys.stdout, 'encoding', '') != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import logging
import time
import requests

from good_report_engine import DRIP_DELAY_SECONDS, _check_and_mark_throttle

logger = logging.getLogger("june_blast")

GHL_BASE                  = "https://services.leadconnectorhq.com"
_CAO_GOOD_REPORT_FIELD_ID = "Y1lUr9X7ACLNmfsXmsCX"

BLAST_CLIENTS = [
    "Salamati Medical Center",
    "Haitham Adli Hijazi",
    "Dr Ahmad Atout",
    "Abdulhalim Mustafa Shaaban Heba",
    "Row Clinic",
    "Khalidi Rehab Center",
    "Zain Clinic (Dr. Tariq Abdo)",
    "WS Clinic Dental Clinic",
    "Dr. Khaled Hamdan & Dr. Kinda Hamdan",
    "Dr Ashraf Dababneh",
    "Vogue Dental Clinic",
    "Nova Dental Care",
    "Dentaface",
    "Bella Clinic",
    "Rawnaq Medical Clinic",
    "Dr Karam Almishry",
]


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


def _is_confirmed_row(row):
    return len(row) > 16 and any(str(row[i]).strip() != "" for i in [10, 11, 16])


def get_three_month_performance(client_name: str) -> dict:
    default = {
        "total_new_reviews":   0,
        "total_new_patients":  0,
        "avg_attendance":      0.0,
        "total_website_patients": 0,
    }
    try:
        from sheets_manager import get_all_sheet_rows
        rows = get_all_sheet_rows(client_name) or []
        confirmed = [r for r in rows if _is_confirmed_row(r)]
        data = confirmed[-12:]
        if not data:
            return default

        nr  = int(sum(_safe(r, 1)  for r in data))
        np_ = int(sum(_safe(r, 3)  for r in data))
        wp  = int(sum(_safe(r, 31) for r in data))

        att_vals = []
        for r in data:
            if len(r) > 15 and str(r[15]).strip():
                try:
                    att_vals.append((1 - float(str(r[15]).replace("%", "").strip())) * 100)
                except (ValueError, TypeError):
                    pass
        att = round(sum(att_vals) / len(att_vals), 1) if att_vals else 0.0

        return {
            "total_new_reviews":      nr,
            "total_new_patients":     np_,
            "avg_attendance":         att,
            "total_website_patients": wp,
        }
    except Exception as e:
        logger.warning(f"[june_blast] get_three_month_performance({client_name}): {e}")
        return default


def generate_june_message(client: dict, performance: dict) -> str:
    doctor_name = client.get("doctor_name", client["name"])

    nr  = performance["total_new_reviews"]
    np_ = performance["total_new_patients"]
    att = performance["avg_attendance"]
    wp  = performance["total_website_patients"]

    lines = []
    if nr  > 0: lines.append(f"⭐ تقييمات جوجل: حصلنا على {nr} تقييم جديد خلال الفترة الماضية")
    if np_ > 0: lines.append(f"👥 المرضى الجدد: استقبلنا {np_} مريض جديد")
    if att > 0: lines.append(f"📅 معدل الحضور: {att}٪ من المواعيد تمت بنجاح")
    if wp  > 0: lines.append(f"🌐 مرضى الموقع: {wp} مريض جديد عبر الموقع الإلكتروني")

    results_section = "\n".join(lines) if lines else "نعمل باستمرار على تحسين نتائجك"

    message = (
        f"السلام عليكم {doctor_name} 🌹\n"
        "\n"
        "نتمنى أن تكونوا بأفضل حال، ونود أن نشكركم على ثقتكم المستمرة بنا خلال الفترة الماضية. دعمكم وثقتكم هما السبب الرئيسي الذي يدفعنا لمضاعفة جهودنا باستمرار لتحقيق أفضل النتائج لعيادتكم.\n"
        "\n"
        "خلال الشهرين الماضيين، وبفضل التعاون بيننا، حققنا العديد من النتائج الإيجابية، ومن أبرزها:\n"
        "\n"
        f"{results_section}\n"
        "\n"
        "ومن خلال تحليلنا للبيانات والاتجاهات الحديثة، بدأنا نلاحظ تحولًا مهمًا في طريقة بحث المرضى عن الخدمات الطبية، حيث أصبح عدد متزايد منهم يستخدم أدوات الذكاء الاصطناعي مثل ChatGPT وGemini للحصول على ترشيحات للأطباء والعيادات. لذلك سيكون هذا أحد أهم محاور عملنا خلال الأشهر القادمة، بحيث نعمل على تعزيز حضوركم الرقمي ومحتوى موقعكم ليزداد احتمال ظهور عيادتكم ضمن التوصيات التي تقدمها هذه المنصات مستقبلًا.\n"
        "\n"
        "كما نود أن نشارككم أن نظام ملفات المرضى والنظام المحاسبي أصبحا في مراحلهما النهائية، وقريبًا جدًا سيكونان جاهزين للاستخدام، مما سيوفر لكم إدارة أسهل وأكثر كفاءة للعيادة ضمن نظام واحد متكامل.\n"
        "\n"
        "هدفنا لا يقتصر على تقديم خدمة، بل على تحقيق نمو حقيقي ومستدام لعيادتكم، من خلال:\n"
        "- زيادة تتجاوز 50% في تقييمات Google بشكل مستمر.\n"
        "- تحسين ترتيبكم في نتائج البحث شهرًا بعد شهر.\n"
        "- تطوير الموقع الإلكتروني باستمرار وإضافة الكلمات المفتاحية التي يبحث عنها المرضى.\n"
        "- الاستفادة من أحدث التوجهات الرقمية لضمان بقاء عيادتكم في المقدمة.\n"
        "\n"
        "وبهذه المناسبة، نود أيضًا تذكيركم بلطف بأن اشتراك هذا الشهر مستحق، ونكون شاكرين لكم في حال التكرم بإتمام عملية الدفع في الوقت المناسب حتى نواصل العمل على جميع الخطط والتطويرات دون انقطاع.\n"
        "\n"
        "شكرًا مرة أخرى على ثقتكم الغالية، ونتطلع إلى تحقيق إنجازات أكبر معكم خلال الفترة القادمة.\n"
        "\n"
        "نتمنى لكم موسم صيف مليئًا بالنجاح، والمرضى، والبركة. ☀️🌿"
    )

    return message


def send_to_ghl(client: dict, message: str, settings: dict) -> bool:
    threeup_api_key = settings.get("threeup_api_key", "")
    contact_id      = client.get("contact_id", "")
    name            = client.get("name", "")

    if not contact_id:
        logger.error(f"[june_blast] {name} — contact_id missing, skipping")
        print(f"[june_blast] {name} — contact_id missing, skipping")
        return False

    if not threeup_api_key:
        logger.error(f"[june_blast] {name} — threeup_api_key missing")
        print(f"[june_blast] {name} — threeup_api_key missing")
        return False

    # Universal 30-min throttle (same enforcement point as good_report_engine.write_to_ghl)
    _check_and_mark_throttle(name)

    headers = {
        "Authorization": f"Bearer {threeup_api_key}",
        "Version": "2021-07-28",
        "Content-Type": "application/json",
    }
    contact_url = f"{GHL_BASE}/contacts/{contact_id}"

    # Write message to cao_good_report field
    r = requests.put(
        contact_url, headers=headers,
        json={"customFields": [{"id": _CAO_GOOD_REPORT_FIELD_ID, "field_value": message}]},
        timeout=30,
    )
    if r.status_code not in (200, 201):
        logger.error(f"[june_blast] {name} — field write failed: {r.status_code} {r.text[:200]}")
        print(f"[june_blast] {name} — field write failed: {r.status_code}")
        return False

    # Delete stale tag first so the workflow fires on re-add
    requests.delete(
        f"{contact_url}/tags", headers=headers,
        json={"tags": ["good-report-ready"]},
        timeout=30,
    )

    # Add tag to trigger GHL workflow
    r2 = requests.post(
        f"{contact_url}/tags", headers=headers,
        json={"tags": ["good-report-ready"]},
        timeout=30,
    )
    print(f"[june_blast] {name} — tag POST {r2.status_code}: {r2.text[:200]}")
    if r2.status_code not in (200, 201):
        logger.error(f"[june_blast] {name} — tag add failed: {r2.status_code} {r2.text[:200]}")
        return False

    logger.info(f"[june_blast] {name} sent OK")
    print(f"[june_blast] {name} ✅")
    return True


def run_june_blast(clients: list, settings: dict) -> None:
    due = [c for c in clients if c["name"] in BLAST_CLIENTS and not c.get("churned")]
    if not due:
        print("[june_blast] No matching clients found")
        return

    # Preserve order from BLAST_CLIENTS
    order = {name: i for i, name in enumerate(BLAST_CLIENTS)}
    due.sort(key=lambda c: order.get(c["name"], 999))

    print(f"[june_blast] Sending to {len(due)} client(s)...")
    sent, failed = 0, 0

    for i, client in enumerate(due):
        name        = client["name"]
        performance = get_three_month_performance(name)
        message     = generate_june_message(client, performance)
        success     = send_to_ghl(client, message, settings)
        if success:
            sent += 1
        else:
            failed += 1

        if i < len(due) - 1:
            print(f"[june_blast] Waiting 30 min before next ({i+2}/{len(due)})...")
            time.sleep(DRIP_DELAY_SECONDS)

    print(f"[june_blast] Done — {sent} sent, {failed} failed")
