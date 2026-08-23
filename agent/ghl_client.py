"""
ghl_client.py — GoHighLevel API client for ThreeUp CAO Agent.
All calls use sub-account API keys with Bearer auth.
Base URL: https://services.leadconnectorhq.com

Fixes applied:
  FIX 1 — Contacts: GHL ignores date params; removed them, filter client-side.
  FIX 2 — Appointments: use /calendars/events with ms timestamps; fallback to /appointments/.
  FIX 3 — Snippets: use /locations/{id}/customValues (location in PATH, not query param).
  FIX 4 — Trigger links: smarter approach — filter recently-updated contacts then check activity + tags.
"""

import logging
import time
from datetime import datetime
from typing import Optional

import pytz
import requests
from dateutil import parser as dateutil_parser

GHL_BASE         = "https://services.leadconnectorhq.com"
AMMAN_TZ         = pytz.timezone("Asia/Amman")
ARABIC_TAG_PREFIX = "مريض اليوم"
MAX_CONTACT_PAGES = 200   # safety ceiling — 200 × 100 = 20 000 contacts

logger = logging.getLogger("ghl_client")


# ── Shared helpers ────────────────────────────────────────────────────────────

def _headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "Content-Type":  "application/json",
        "Version":       "2021-07-28",
    }


def _ts(client_name: str = "?") -> str:
    return f"[{datetime.now(AMMAN_TZ).strftime('%Y-%m-%d %H:%M:%S')}] [{client_name}]"


def _get(url: str, api_key: str, params: dict = None,
         client_name: str = "?", version: str = None) -> Optional[dict]:
    """GET with 3× retry on 429, full error logging."""
    hdrs = _headers(api_key)
    if version:
        hdrs["Version"] = version
    for attempt in range(3):
        try:
            resp = requests.get(url, headers=hdrs, params=params, timeout=30)
            if resp.status_code == 429:
                logger.warning(f"{_ts(client_name)} [ghl_client] 429 — waiting 2s (attempt {attempt+1}/3)")
                time.sleep(2)
                continue
            if not resp.ok:
                logger.warning(f"{_ts(client_name)} [ghl_client] HTTP {resp.status_code} for {url} | {resp.text[:200]}")
                return None
            logger.info(f"{_ts(client_name)} [ghl_client] GET {url} → {resp.status_code}")
            time.sleep(0.5)   # rate-limit buffer between API calls
            return resp.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"{_ts(client_name)} [ghl_client] GET {url} error: {e}")
            if attempt < 2:
                time.sleep(2)
    return None


def _to_iso(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = AMMAN_TZ.localize(dt)
    return dt.strftime("%Y-%m-%dT%H:%M:%S%z")


def _to_ms(dt: datetime) -> int:
    """Convert datetime to Unix timestamp in milliseconds."""
    if dt.tzinfo is None:
        dt = AMMAN_TZ.localize(dt)
    return int(dt.timestamp() * 1000)


def _parse_dt(value) -> Optional[datetime]:
    """
    Parse a GHL date value to a timezone-aware datetime in Asia/Amman.

    GHL stores dateAdded / dateUpdated as Unix timestamps in MILLISECONDS
    (integer or float), NOT seconds.  ISO strings (with Z or +offset) are
    also accepted for API fields that still return ISO format.

    Rules:
      - int / float          → milliseconds  → UTC → Amman
      - digit-only string    → milliseconds  → UTC → Amman
      - ISO / RFC string     → dateutil.parse → Amman
    Returns None on any failure.
    """
    if value is None:
        return None
    try:
        # Numeric milliseconds (most common GHL format for dateAdded/dateUpdated)
        if isinstance(value, (int, float)):
            return (datetime.fromtimestamp(int(value) / 1000, tz=pytz.UTC)
                            .astimezone(AMMAN_TZ))
        s = str(value).strip()
        if s.isdigit():
            return (datetime.fromtimestamp(int(s) / 1000, tz=pytz.UTC)
                            .astimezone(AMMAN_TZ))
        # ISO 8601 / RFC 2822 string
        return dateutil_parser.parse(s).astimezone(AMMAN_TZ)
    except (ValueError, TypeError, OverflowError, Exception):
        return None


def _normalize_tag(tag: str) -> str:
    """Strip whitespace and invisible Unicode direction/space characters."""
    return (
        tag.strip()
           .replace('\u200f', '')   # RIGHT-TO-LEFT MARK
           .replace('\u200e', '')   # LEFT-TO-RIGHT MARK
           .replace('\u200b', '')   # ZERO-WIDTH SPACE
           .replace('\u200c', '')   # ZERO-WIDTH NON-JOINER
           .replace('\u200d', '')   # ZERO-WIDTH JOINER
           .replace('\xa0',   ' ')  # NON-BREAKING SPACE → regular space
           .strip()
    )


def _has_arabic_tag(contact: dict) -> bool:
    """Return True if any tag starts with 'مريض اليوم' after Unicode normalisation."""
    for tag in (contact.get("tags") or []):
        if isinstance(tag, str):
            norm = _normalize_tag(tag)
            if norm.startswith(ARABIC_TAG_PREFIX) or norm == ARABIC_TAG_PREFIX:
                return True
    return False


# ── FIX 1: Contact pagination — no date params sent to API ───────────────────

def _collect_all_contacts(api_key: str, location_id: str,
                          client_name: str = "?") -> list:
    """
    Paginate through ALL contacts for a location.
    GHL ignores startDate/endDate on the contacts endpoint — so we never send them.
    Handles nextPageUrl, startAfterId cursor, and page-number pagination.
    Deduplicates by contact ID and stops when a page adds no new contacts (cursor cycle).
    """
    all_contacts: list = []
    seen_ids: set = set()
    url = f"{GHL_BASE}/contacts/"
    params = {"locationId": location_id, "limit": 100}
    page_count = 0

    while page_count < MAX_CONTACT_PAGES:
        data = _get(url, api_key, params, client_name)
        if not data:
            break

        contacts = data.get("contacts") or []
        page_count += 1

        if len(contacts) == 0:
            break

        new_this_page = 0
        for c in contacts:
            cid = c.get("id")
            if cid and cid not in seen_ids:
                seen_ids.add(cid)
                all_contacts.append(c)
                new_this_page += 1

        if new_this_page == 0:
            logger.info(
                f"{_ts(client_name)} [_collect_all_contacts] cursor cycle detected "
                f"after {page_count} pages — stopping (total unique: {len(all_contacts)})"
            )
            break

        meta = data.get("meta") or {}

        # Prefer explicit nextPageUrl from top-level or meta
        next_url = data.get("nextPageUrl") or meta.get("nextPageUrl")
        if next_url:
            url = next_url
            params = {}
            continue

        # Cursor-based: startAfterId or nextCursor
        cursor = meta.get("startAfterId") or meta.get("nextCursor") or data.get("nextCursor")
        if cursor and len(contacts) >= 100:
            params = {"locationId": location_id, "limit": 100, "startAfterId": cursor}
            url = f"{GHL_BASE}/contacts/"
            continue

        # Page-number based
        current_page = meta.get("currentPage")
        total_pages  = meta.get("totalPages") or meta.get("total_pages")
        if current_page is not None and total_pages is not None:
            if int(current_page) < int(total_pages):
                params = {"locationId": location_id, "limit": 100,
                          "page": int(current_page) + 1}
                url = f"{GHL_BASE}/contacts/"
                continue

        # No more pages
        break

    # FIX 4 — diagnostic tag summary
    has_any_tags   = sum(1 for c in all_contacts if c.get("tags"))
    has_arabic_tag = sum(1 for c in all_contacts if _has_arabic_tag(c))
    logger.info(
        f"{_ts(client_name)} [_collect_all_contacts] "
        f"total={len(all_contacts)} with_any_tags={has_any_tags} "
        f"with_arabic_tag={has_arabic_tag}"
    )
    logger.info(
        f"{_ts(client_name)} [audit] {client_name} — pagination: "
        f"{page_count} pages fetched"
    )
    return all_contacts


# ── FIX 1: get_new_patients ──────────────────────────────────────────────────

def get_new_patients(api_key: str, location_id: str, start_date: datetime,
                     end_date: datetime, client_name: str = "?",
                     new_patient_tag: str = None) -> int:
    """
    Count contacts whose dateAdded falls within [start_date, end_date]
    AND whose tag exactly matches new_patient_tag (after Unicode normalisation).
    Defaults to the Arabic 'مريض اليوم' tag when new_patient_tag is None/empty.
    GHL ignores date params on /contacts/ — fetch all, filter client-side.
    """
    # Resolve which tag to match — exact match after normalisation
    _raw_tag   = new_patient_tag if new_patient_tag else ARABIC_TAG_PREFIX
    _match_tag = _normalize_tag(_raw_tag).lower()

    def _has_patient_tag(contact: dict) -> bool:
        for tag in (contact.get("tags") or []):
            if isinstance(tag, str) and _normalize_tag(tag).lower() == _match_tag:
                return True
        return False

    try:
        contacts  = _collect_all_contacts(api_key, location_id, client_name)
        start_ts  = start_date.timestamp()
        end_ts    = end_date.timestamp()

        count             = 0
        has_tag_total     = 0
        passes_date_total = 0

        for c in contacts:
            if not _has_patient_tag(c):
                continue
            has_tag_total += 1

            raw_da   = c.get("dateAdded") or c.get("createdAt")
            dt_amman = _parse_dt(raw_da)

            if dt_amman and start_ts <= dt_amman.timestamp() <= end_ts:
                passes_date_total += 1
                count += 1

        logger.info(
            f"{_ts(client_name)} [get_new_patients] tag={_raw_tag!r} "
            f"has_tag={has_tag_total} passes_date={passes_date_total} result={count}"
        )
        logger.info(
            f"{_ts(client_name)} [audit] {client_name} new patients: "
            f"fetched {len(contacts)} contacts, "
            f"{passes_date_total} matched tag+date filter"
        )
        return count

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_new_patients] exception: {e}")
        return 0


# ── FIX 1: get_returning_patients_tagged ─────────────────────────────────────

def get_returning_patients_tagged(api_key: str, location_id: str, start_date: datetime,
                                   end_date: datetime, client_name: str = "?") -> int:
    """
    Count contacts with the Arabic tag whose dateAdded is BEFORE start_date
    (existing contacts that were re-tagged during this week).
    """
    try:
        contacts = _collect_all_contacts(api_key, location_id, client_name)
        start_ts = start_date.timestamp()
        count    = 0

        for c in contacts:
            if not _has_arabic_tag(c):
                continue
            # Must be updated/tagged within the window (dateUpdated in range)
            # but originally created before the window (dateAdded < start)
            dt_added   = _parse_dt(c.get("dateAdded") or c.get("createdAt"))
            dt_updated = _parse_dt(c.get("dateUpdated") or c.get("dateModified"))

            added_before_window   = dt_added   and dt_added.timestamp()   < start_ts
            updated_within_window = dt_updated and (start_date.timestamp()
                                                    <= dt_updated.timestamp()
                                                    <= end_date.timestamp())
            if added_before_window and updated_within_window:
                count += 1

        logger.info(f"{_ts(client_name)} [get_returning_patients_tagged] → {count}")
        return count

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_returning_patients_tagged] exception: {e}")
        return 0


# ── FIX 2: get_trigger_link_clicks — tag-based filter ────────────────────────

TRIGGER_LINK_TAG = "trigger link clicked"   # exact GHL tag name applied on click

def get_trigger_link_clicks(api_key: str, location_id: str, start_date: datetime,
                             end_date: datetime, client_name: str = "?",
                             trigger_link_tag: str = None) -> int:
    """
    Count contacts that:
      1. Have a tag exactly matching trigger_link_tag (default: "trigger link clicked"),
         case-insensitive after Unicode normalisation.
      2. Were updated (dateUpdated, fallback dateAdded) within [start_date, end_date].
    Deduplicates by contact ID.
    """
    _raw_tag   = trigger_link_tag if trigger_link_tag else TRIGGER_LINK_TAG
    _match_tag = _normalize_tag(_raw_tag).lower()

    start_ts = start_date.timestamp()
    end_ts   = end_date.timestamp()

    try:
        contacts   = _collect_all_contacts(api_key, location_id, client_name)
        unique_ids: set = set()

        for c in contacts:
            contact_id = c.get("id")
            if not contact_id:
                continue

            has_tag = any(
                _normalize_tag(tag).lower() == _match_tag
                for tag in (c.get("tags") or [])
                if isinstance(tag, str)
            )
            if not has_tag:
                continue

            # Date filter: contact must have been updated this week
            raw_dt   = c.get("dateUpdated") or c.get("dateAdded") or c.get("createdAt")
            dt_amman = _parse_dt(raw_dt)
            if dt_amman and start_ts <= dt_amman.timestamp() <= end_ts:
                unique_ids.add(contact_id)

        total = len(contacts)
        count = len(unique_ids)

        if total == 0:
            logger.warning(
                f"{_ts(client_name)} [audit] WARNING {client_name} — "
                f"0 contacts fetched. Possible API failure. Trigger links set to 0."
            )
        logger.info(
            f"{_ts(client_name)} [audit] {client_name} trigger links: "
            f"fetched {total} contacts, {count} clicked this week (date-filtered)"
        )
        return count

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_trigger_link_clicks] exception: {e}")
        return 0


# ── Calendar helpers ──────────────────────────────────────────────────────────

def get_all_calendar_ids(api_key: str, location_id: str,
                         client_name: str = "?") -> list:
    """
    Return a list of IDs for all active calendars in this location.
    GET /calendars/?locationId={location_id}
    Falls back to all calendars (active or not) if none are flagged active.
    """
    data = _get(f"{GHL_BASE}/calendars/", api_key,
                {"locationId": location_id}, client_name)
    if not data:
        return []
    calendars = data.get("calendars") or []

    active_ids = [
        cal["id"] for cal in calendars
        if cal.get("id") and (
            cal.get("isActive") or cal.get("isEnabled") or cal.get("active")
        )
    ]

    if not active_ids:
        # Fallback: use every calendar that has an ID
        active_ids = [cal["id"] for cal in calendars if cal.get("id")]

    logger.info(f"{_ts(client_name)} [get_appointments] Found {len(active_ids)} calendars")
    return active_ids


def get_appointments(api_key: str, location_id: str, start_date: datetime,
                     end_date: datetime, client_name: str = "?") -> dict:
    """
    Fetch all calendar events across all active calendars for the date range.
    Auto-discovers calendar IDs via get_all_calendar_ids(), then queries
    GET /calendars/events for each with startTime/endTime ms filters.
    Paginates each calendar until all events are retrieved.
    Returns {booked, confirmed, cancelled, cancellation_rate}.
    """
    empty = {"booked": 0, "confirmed": 0, "cancelled": 0, "cancellation_rate": 0.0}

    _EXCLUDE_STATUSES = {"invalid", "test"}

    def _count_events(events: list) -> tuple:
        """
        Return (booked, confirmed, cancelled, meaningful) for a batch of events.
        booked      = total raw event count
        meaningful  = confirmed + cancelled + showed + new + noshow
                      (excludes invalid/test statuses)
        confirmed   = status "confirmed"
        cancelled   = status "cancelled"
        cancellation_rate uses meaningful as denominator, not booked.
        """
        status_counts = {}
        for event in events:
            s = event.get("appointmentStatus") or event.get("status") or "unknown"
            status_counts[s] = status_counts.get(s, 0) + 1
        logger.info(
            f"[{client_name}] [get_appointments] Status breakdown: {status_counts}"
        )
        booked    = len(events)
        confirmed = 0
        cancelled = 0
        meaningful = 0
        for e in events:
            s = (e.get("appointmentStatus") or e.get("status") or "").lower().strip()
            if any(excl in s for excl in _EXCLUDE_STATUSES):
                continue
            meaningful += 1
            if s == "confirmed":
                confirmed += 1
            elif s == "cancelled":
                cancelled += 1
        return booked, confirmed, cancelled, meaningful

    def _fetch_all_events_paginated(cal_id: str) -> list:
        """Fetch all pages of events for one calendar ID."""
        all_events = []
        params = {
            "locationId": location_id,
            "calendarId": cal_id,
            "startTime":  _to_ms(start_date),
            "endTime":    _to_ms(end_date),
        }
        page = 1
        while True:
            data = _get(f"{GHL_BASE}/calendars/events", api_key, params, client_name)
            if not data:
                break
            batch = data.get("events") or data.get("data") or []
            if isinstance(batch, list):
                all_events.extend(batch)
            next_token = (data.get("nextPageToken") or
                          data.get("meta", {}).get("nextPageToken"))
            if next_token:
                params["pageToken"] = next_token
                page += 1
                logger.info(f"{_ts(client_name)} [get_appointments] "
                            f"cal={cal_id} fetching page {page}")
            else:
                break
        return all_events

    try:
        cal_ids = get_all_calendar_ids(api_key, location_id, client_name)

        if cal_ids:
            grand_booked      = 0
            grand_confirmed   = 0
            grand_cancelled   = 0
            grand_meaningful  = 0

            for cid in cal_ids:
                events                        = _fetch_all_events_paginated(cid)
                book, conf, canc, meaningful  = _count_events(events)
                grand_booked     += book
                grand_confirmed  += conf
                grand_cancelled  += canc
                grand_meaningful += meaningful
                logger.info(
                    f"{_ts(client_name)} [get_appointments] "
                    f"calendarId={cid} booked={book} meaningful={meaningful} "
                    f"confirmed={conf} cancelled={canc}"
                )

            rate = round(grand_cancelled / grand_meaningful, 4) if grand_meaningful > 0 else 0.0
            logger.info(
                f"{_ts(client_name)} [get_appointments] combined "
                f"booked={grand_booked} meaningful={grand_meaningful} "
                f"confirmed={grand_confirmed} cancelled={grand_cancelled} rate={rate}"
            )
            logger.info(
                f"{_ts(client_name)} [audit] {client_name} appointments: "
                f"booked={grand_booked} meaningful={grand_meaningful} "
                f"confirmed={grand_confirmed} cancelled={grand_cancelled} "
                f"rate={round(rate * 100, 1)}%"
            )
            return {
                "booked":            grand_booked,
                "confirmed":         grand_confirmed,
                "cancelled":         grand_cancelled,
                "cancellation_rate": rate,
            }

        logger.warning(f"{_ts(client_name)} [get_appointments] "
                       "No calendars found — falling back to /appointments/")

        # Fallback: /appointments/
        params_fb = {
            "locationId": location_id,
            "startDate":  _to_iso(start_date),
            "endDate":    _to_iso(end_date),
        }
        data_fb = _get(f"{GHL_BASE}/appointments/", api_key, params_fb, client_name)
        if data_fb:
            events_fb = (data_fb.get("appointments") or
                         data_fb.get("events")       or
                         data_fb.get("data")         or [])
            b_fb, conf_fb, canc_fb, meaningful_fb = _count_events(events_fb)
            rate_fb = round(canc_fb / meaningful_fb, 4) if meaningful_fb > 0 else 0.0
            logger.info(
                f"{_ts(client_name)} [get_appointments] /appointments/ fallback "
                f"booked={b_fb} meaningful={meaningful_fb} "
                f"confirmed={conf_fb} cancelled={canc_fb}"
            )
            return {
                "booked":            b_fb,
                "confirmed":         conf_fb,
                "cancelled":         canc_fb,
                "cancellation_rate": rate_fb,
            }

        logger.warning(f"{_ts(client_name)} [get_appointments] Both endpoints failed — returning zeros")
        return empty

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_appointments] exception: {e}")
        return empty


# ── get_loyalty_points (unchanged logic, uses fixed _collect_all_contacts) ───

def get_loyalty_points(api_key: str, location_id: str, start_date: datetime,
                       end_date: datetime, client_name: str = "?") -> int:
    """
    Aggregate loyalty points from custom fields across contacts active in period.
    Looks for fields named 'loyalty points' or containing 'نقاط'.
    """
    try:
        # Discover loyalty field ID
        cf_data = _get(f"{GHL_BASE}/locations/{location_id}/customFields",
                       api_key, {}, client_name)
        loyalty_field_id = None
        if cf_data:
            for field in (cf_data.get("customFields") or []):
                name = (field.get("name") or "").lower()
                if "loyalty points" in name or "نقاط" in name:
                    loyalty_field_id = field.get("id")
                    break

        contacts     = _collect_all_contacts(api_key, location_id, client_name)
        start_ts     = start_date.timestamp()
        end_ts       = end_date.timestamp()
        total_points = 0

        for c in contacts:
            # Only count contacts active/updated in the period
            dt = _parse_dt(c.get("dateUpdated") or c.get("dateAdded"))
            if dt and not (start_ts <= dt.timestamp() <= end_ts):
                continue

            for cf in (c.get("customFields") or []):
                field_id  = cf.get("id") or cf.get("fieldId") or ""
                field_key = (cf.get("key") or "").lower()
                matched   = (
                    (loyalty_field_id and field_id == loyalty_field_id) or
                    "loyalty" in field_key or "نقاط" in field_key
                )
                if matched:
                    val = cf.get("value") or cf.get("fieldValue") or 0
                    try:
                        total_points += int(float(str(val).replace(",", "")))
                    except (ValueError, TypeError):
                        pass
                    break

        logger.info(f"{_ts(client_name)} [get_loyalty_points] → {total_points}")
        return total_points

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_loyalty_points] exception: {e}")
        return 0


# ── get_active_workflow (unchanged) ──────────────────────────────────────────

def get_active_workflow(api_key: str, location_id: str,
                        client_name: str = "?") -> str:
    """
    Return the name of the first workflow containing '(in-use)'.
    Returns "Unknown" if none found.
    """
    try:
        data = _get(f"{GHL_BASE}/workflows/", api_key,
                    {"locationId": location_id}, client_name)
        if not data:
            logger.info(f"{_ts(client_name)} [get_active_workflow] No data — returning Unknown")
            return "Unknown"

        active = [
            w.get("name", "")
            for w in (data.get("workflows") or [])
            if "(in-use)" in (w.get("name") or "").lower()
        ]
        if not active:
            logger.info(f"{_ts(client_name)} [get_active_workflow] No (in-use) workflows — returning Unknown")
            return "Unknown"

        logger.info(f"{_ts(client_name)} [get_active_workflow] → {active[0]!r}")
        return active[0]

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_active_workflow] exception: {e}")
        return "Unknown"


# ── FIX 3: get_active_snippet ────────────────────────────────────────────────

def get_active_snippet(api_key: str, location_id: str, workflow_name: str,
                       client_name: str = "?") -> Optional[dict]:
    """
    Fetch active email templates / custom values for this location.
    Tries three endpoints in order; falls through to the next if the previous
    returns nothing.  Filters entries whose name contains '(in-use)' (case-
    insensitive).  Returns ALL matching names joined with ' | ' and their
    combined body/value text for downstream personalisation analysis.
    Returns {name: None, value: None} if nothing is found.
    """
    no_match = {"name": "Unknown", "value": ""}
    try:
        items: list = []

        # Attempt 1: /templates/?locationId=
        data = _get(f"{GHL_BASE}/templates/", api_key,
                    {"locationId": location_id}, client_name)
        if data:
            raw = data.get("templates") or data.get("data") or []
            if raw:
                items = raw
                logger.info(f"{_ts(client_name)} [get_active_snippet] "
                            f"Got {len(raw)} items from /templates/")

        # Attempt 2: /customValues/?locationId=
        if not items:
            data = _get(f"{GHL_BASE}/customValues/", api_key,
                        {"locationId": location_id}, client_name)
            if data:
                raw = data.get("customValues") or data.get("data") or []
                if raw:
                    items = raw
                    logger.info(f"{_ts(client_name)} [get_active_snippet] "
                                f"Got {len(raw)} items from /customValues/")

        # Attempt 3: /locations/{id}/customValues
        if not items:
            data = _get(f"{GHL_BASE}/locations/{location_id}/customValues",
                        api_key, {}, client_name)
            if data:
                raw = data.get("customValues") or data.get("data") or []
                if raw:
                    items = raw
                    logger.info(f"{_ts(client_name)} [get_active_snippet] "
                                f"Got {len(raw)} items from /locations/.../customValues")

        if not items:
            logger.info(f"{_ts(client_name)} [get_active_snippet] No items from any endpoint")
            return no_match

        # Filter by "(in-use)" in name
        active = [
            item for item in items
            if "(in-use)" in (item.get("name") or "").lower()
        ]

        if not active:
            logger.info(f"{_ts(client_name)} [get_active_snippet] No items with '(in-use)' in name")
            return no_match

        names = [item.get("name", "") for item in active]
        values = [
            item.get("value") or item.get("body") or item.get("html") or ""
            for item in active
        ]
        combined_name  = " | ".join(n for n in names  if n)
        combined_value = " | ".join(v for v in values if v)

        logger.info(
            f"{_ts(client_name)} [get_active_snippet] "
            f"Found {len(active)} active templates: {names}"
        )
        return {"name": combined_name, "value": combined_value}

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_active_snippet] exception: {e}")
        return no_match


# ── get_conversations_sample (unchanged) ─────────────────────────────────────

def get_conversations_sample(api_key: str, location_id: str,
                              client_name: str = "?", limit: int = 30) -> list:
    """Fetch up to `limit` conversations and return flat list of message strings."""
    try:
        # FIX 6 — use /conversations/search (correct endpoint)
        data = _get(f"{GHL_BASE}/conversations/search", api_key,
                    {"locationId": location_id, "limit": limit}, client_name)
        if not data:
            return []

        conversations = data.get("conversations") or []
        all_messages  = []

        for conv in conversations:
            conv_id = conv.get("id")
            if not conv_id:
                continue
            msg_data = _get(f"{GHL_BASE}/conversations/{conv_id}/messages",
                            api_key, {}, client_name)
            if not msg_data:
                continue

            messages = msg_data.get("messages") or msg_data.get("data") or []
            if isinstance(messages, dict):
                messages = messages.get("messages") or []

            for m in messages:
                body = m.get("body") or m.get("message") or m.get("text") or ""
                if body and isinstance(body, str) and body.strip():
                    all_messages.append(body.strip())

        logger.info(f"{_ts(client_name)} [get_conversations_sample] → "
                    f"{len(all_messages)} messages from {len(conversations)} conversations")
        return all_messages

    except Exception as e:
        logger.error(f"{_ts(client_name)} [get_conversations_sample] exception: {e}")
        return []


def get_has_recent_ghl_activity(api_key: str, location_id: str,
                                window_days: int = 14, client_name: str = "?") -> bool:
    """
    Returns True if the GHL sub-account has had any activity in the last window_days.
    Checks conversations (last_message_date) then contacts (dateUpdated/dateAdded).
    Returns True on any API error so we never false-alert due to a broken key.
    """
    from datetime import datetime, timezone, timedelta
    cutoff = datetime.now(timezone.utc) - timedelta(days=window_days)

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type":  "application/json",
        "Version":       "2021-07-28",
    }

    # Signal 1: conversations ordered by last_message_date
    try:
        url = (f"https://services.leadconnectorhq.com/conversations/search"
               f"?locationId={location_id}&limit=50"
               f"&sortBy=last_message_date&sortOrder=desc")
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 200:
            convs = r.json().get("conversations") or []
            for c in convs:
                for key in ("lastMessageDate", "updatedAt", "dateUpdated"):
                    raw = c.get(key)
                    if not raw:
                        continue
                    dt = _parse_dt(raw)
                    if dt and dt >= cutoff:
                        logger.info(f"{_ts(client_name)} [ghl_activity] recent conversation ({key}={raw})")
                        return True
        elif r.status_code not in (400, 404):
            logger.warning(f"{_ts(client_name)} [ghl_activity] conversations status {r.status_code} — treating as active")
            return True
    except Exception as e:
        logger.warning(f"{_ts(client_name)} [ghl_activity] conversations error: {e} — treating as active")
        return True

    # Signal 2: contacts page 1 ordered by default (most recently updated first)
    try:
        url = f"https://services.leadconnectorhq.com/contacts/?locationId={location_id}&limit=100"
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 200:
            contacts = r.json().get("contacts") or []
            for c in contacts:
                for key in ("dateUpdated", "dateAdded"):
                    raw = c.get(key)
                    if not raw:
                        continue
                    dt = _parse_dt(raw)
                    if dt and dt >= cutoff:
                        logger.info(f"{_ts(client_name)} [ghl_activity] recent contact ({key}={raw})")
                        return True
        elif r.status_code not in (400, 404):
            logger.warning(f"{_ts(client_name)} [ghl_activity] contacts status {r.status_code} — treating as active")
            return True
    except Exception as e:
        logger.warning(f"{_ts(client_name)} [ghl_activity] contacts error: {e} — treating as active")
        return True

    logger.info(f"{_ts(client_name)} [ghl_activity] no recent activity in last {window_days}d")
    return False
