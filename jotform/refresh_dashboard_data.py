import json
import os
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv


ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT.parent / ".env"
JOTFORM_API_BASE = "https://api.jotform.com"

# Rego → expiry date lookup (populated from rego_fleet.xlsx via rego_expiry.json)
_REGO_EXPIRY_PATH = ROOT / "rego_expiry.json"
REGO_EXPIRY: dict[str, str] = {}
if _REGO_EXPIRY_PATH.exists():
    with open(_REGO_EXPIRY_PATH, encoding="utf-8") as _f:
        REGO_EXPIRY = {k.upper(): v for k, v in json.load(_f).items()}


@dataclass(frozen=True)
class FieldMap:
    date: str
    driver: str
    site: str | None = None
    attachments: str | None = None
    loads: str | None = None
    hours: str | None = None
    waiting: str | None = None
    rego: str | None = None


RUNSHEET_FIELDS = FieldMap(
    date="24",
    driver="26",
    site="59",
    attachments="41",
    loads="88",
    hours="58",
    waiting="77",
    rego="62",
)

PRESTART_FIELDS = FieldMap(
    date="143",
    driver="76",
    rego="77",
)

INSPECTION_FIELD_IDS = [
    "41", "42", "44", "45", "46", "47", "48", "49", "50", "51", "52",
    "53", "54", "55", "56", "57", "75", "144",
]
INSPECTION_LABELS: dict[str, str] = {
    "41": "Pan internals clean",
    "42": "Tyres roadworthy",
    "44": "Wheels secure",
    "45": "Mud flaps present",
    "46": "Lights & lenses OK",
    "47": "Mirrors OK",
    "48": "Windscreen OK",
    "49": "Wipers & washers OK",
    "50": "Bodywork secure",
    "51": "Fluids OK",
    "52": "Brake pedal firm",
    "53": "Suspension secure",
    "54": "Rego plate attached",
    "55": "All gauges operating",
    "56": "Cabin items secure",
    "57": "Seat belt OK",
    "75": "Load within limits",
    "144": "Load restraint OK",
}
FIT_FOR_DUTY_FIELD_IDS = ["116", "117", "119", "120", "121", "123", "125", "127"]
FFD_LABELS: dict[str, str] = {
    "116": "Not fatigued/unwell",
    "117": "Will report if unfit",
    "119": "24hr stationary rest",
    "120": "10hr rest between shifts",
    "121": "No alcohol or drugs",
    "123": "No drug influence",
    "125": "Fit for tasks",
    "127": "Valid licence held",
}
SERVICE_STICKER_FIELD_ID = "68"
SERVICE_DUE_FIELD_ID = "128"
VCR_FIELD_ID = "80"
DEFECT_PROOF_FIELD_ID = "103"
SERVICE_STICKER_PROOF_FIELD_ID = "147"
ODOMETER_FIELD_ID = "65"
SERVICE_REMAINING_KM_FIELD_ID = "140"
CONTRACT_FIELD_ID = "124"
SHAPE_PRESTART_FIELD_ID = "141"
SHIFT_TIMING_FIELD_ID = "72"
SHAPE_RUNSHEET_FIELD_ID = "91"


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def parse_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, dict):
        lowered = {str(k).lower(): v for k, v in value.items()}
        if lowered.get("year") and lowered.get("month") and lowered.get("day"):
            try:
                return date(int(lowered["year"]), int(lowered["month"]), int(lowered["day"]))
            except ValueError:
                return None
        for key in ("datetime", "date", "prettyformat", "prettyFormat"):
            if lowered.get(key.lower()):
                return parse_date(lowered[key.lower()])
    text = str(value).strip()
    if not text:
        return None
    text = text.split("T", 1)[0]
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%m/%d/%Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d %B %Y",
        "%d %b %Y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_datetime(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text.split("+", 1)[0], fmt)
        except ValueError:
            pass
    return None


def answer(submission: dict[str, Any], field_id: str | None) -> Any:
    if not field_id:
        return None
    item = (submission.get("answers") or {}).get(str(field_id))
    if not item:
        return None
    return item.get("answer")


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        for key in ("prettyFormat", "prettyformat", "datetime", "date", "name", "value"):
            if value.get(key):
                return str(value[key]).strip()
        first = " ".join(str(v).strip() for v in value.values() if str(v).strip())
        return re.sub(r"\s+", " ", first).strip()
    if isinstance(value, list):
        return ", ".join(text_value(item) for item in value if text_value(item))
    return str(value).strip()


def driver_key(driver: str) -> str:
    tokens = re.findall(r"[a-z0-9]+", driver.casefold())
    return " ".join(sorted(tokens))


def number_value(value: Any) -> float:
    text = text_value(value)
    if not text:
        return 0.0
    match = re.search(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    return float(match.group(0)) if match else 0.0


def attachment_count(value: Any) -> int:
    if not value:
        return 0
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return 0
        if stripped.startswith("[") or stripped.startswith("{"):
            try:
                return attachment_count(json.loads(stripped))
            except json.JSONDecodeError:
                pass
        return len([part for part in re.split(r"[\n,]+", stripped) if part.strip().startswith("http")])
    if isinstance(value, list):
        return sum(attachment_count(item) for item in value)
    if isinstance(value, dict):
        if any(value.get(key) for key in ("url", "fileUrl", "file_url", "link")):
            return 1
        return sum(attachment_count(item) for item in value.values())
    return 0


def is_yes(value: Any) -> bool:
    text = text_value(value).casefold()
    return text in {"yes", "y", "true", "checked", "1"} or "yes" in text


def is_no(value: Any) -> bool:
    text = text_value(value).casefold()
    return text in {"no", "n", "false", "unchecked", "0"} or text.startswith("no")


def is_checked(value: Any) -> bool:
    if isinstance(value, list):
        if not value:
            return False
        text = text_value(value).casefold()
        if text in {"no", "false", "unchecked", "0"}:
            return False
        if text in {"yes", "true", "checked", "1"}:
            return True
        return bool(text)
    text = text_value(value).casefold()
    return bool(text) and text not in {"no", "false", "unchecked", "0"}


def jotform_get(api_key: str, endpoint: str, params: dict[str, Any] | None = None) -> Any:
    request_params = dict(params or {})
    request_params["apiKey"] = api_key
    response = requests.get(f"{JOTFORM_API_BASE}{endpoint}", params=request_params, timeout=90)
    response.raise_for_status()
    payload = response.json()
    if payload.get("responseCode") != 200:
        raise RuntimeError(f"Jotform API error: {payload}")
    return payload.get("content")


def fetch_submissions(api_key: str, form_id: str, created_cutoff: date) -> list[dict[str, Any]]:
    submissions: list[dict[str, Any]] = []
    offset = 0
    limit = 1000
    cutoff_dt = datetime.combine(created_cutoff, datetime.min.time())
    created_filter = json.dumps({"created_at:gt": f"{created_cutoff.isoformat()} 00:00:00"})
    while True:
        batch = jotform_get(
            api_key,
            f"/form/{form_id}/submissions",
            {
                "limit": limit,
                "offset": offset,
                "orderby": "created_at",
                "filter": created_filter,
            },
        )
        if not batch:
            break
        for submission in batch:
            created = parse_datetime(submission.get("created_at"))
            if created is None or created >= cutoff_dt:
                submissions.append(submission)
        if len(batch) < limit:
            break
        offset += limit
    return submissions


def week_ending_for(day: date) -> date:
    return day + timedelta(days=(6 - day.weekday()))


def empty_week(week_ending: date) -> dict[str, Any]:
    return {
        "weekEnding": week_ending.isoformat(),
        "weekStart": (week_ending - timedelta(days=6)).isoformat(),
        "runsheets": 0,
        "runsheetPdfs": 0,
        "runsheetDrivers": 0,
        "loads": 0,
        "hours": 0,
        "waiting": 0,
        "missingAttachments": 0,
        "prestarts": 0,
        "prestartDrivers": 0,
        "vcrRaised": 0,
        "defectProofUploads": 0,
        "serviceDue": 0,
        "stickerMissing": 0,
        "inspectionExceptions": 0,
        "fatigueExceptions": 0,
        "dayShifts": 0,
        "nightShifts": 0,
        "submissionDelayTotal": 0,  # sum of hours from shift date to prestart submit
        "submissionDelayCount": 0,
        "inspectionBreakdown": {v: 0 for v in INSPECTION_LABELS.values()},
        "ffdBreakdown": {v: 0 for v in FFD_LABELS.values()},
    }


def empty_month(month_start: date) -> dict[str, Any]:
    return {
        "key": month_start.strftime("%Y-%m"),
        "label": month_start.strftime("%b %Y"),
        "runsheets": 0,
        "runsheetPdfs": 0,
        "runsheetDrivers": 0,
        "loads": 0,
        "hours": 0,
        "waiting": 0,
        "missingAttachments": 0,
        "prestarts": 0,
        "prestartDrivers": 0,
        "vcrRaised": 0,
        "defectProofUploads": 0,
        "serviceDue": 0,
        "stickerMissing": 0,
        "inspectionExceptions": 0,
        "fatigueExceptions": 0,
        "dayShifts": 0,
        "nightShifts": 0,
        "submissionDelayTotal": 0,
        "submissionDelayCount": 0,
        "inspectionBreakdown": {v: 0 for v in INSPECTION_LABELS.values()},
        "ffdBreakdown": {v: 0 for v in FFD_LABELS.values()},
    }


def empty_driver_row(driver: str) -> dict[str, Any]:
    return {
        "driver": driver,
        "runsheetPdfs": 0,
        "runsheets": 0,
        "prestarts": 0,
        "loads": 0,
        "hours": 0,
        "waiting": 0,
        "missingAttachments": 0,
        "serviceDue": 0,
        "inspection": 0,
        "fatigue": 0,
        "vcr": 0,
        "lateSubmissions": 0,
    }


def empty_detail(label: str, start: date, end: date) -> dict[str, Any]:
    return {
        "label": label,
        "start": start.isoformat(),
        "end": end.isoformat(),
        "drivers": {},
        "vehicles": {},
        "sites": {},
        "runsheetDates": defaultdict(set),
        "prestartDates": defaultdict(set),
        "inspectionBreakdown": {v: 0 for v in INSPECTION_LABELS.values()},
        "ffdBreakdown": {v: 0 for v in FFD_LABELS.values()},
    }


def month_start_for(day: date) -> date:
    return date(day.year, day.month, 1)


def iter_month_starts(start: date, end: date) -> list[date]:
    months = []
    cursor = date(start.year, start.month, 1)
    final = date(end.year, end.month, 1)
    while cursor <= final:
        months.append(cursor)
        year = cursor.year + (1 if cursor.month == 12 else 0)
        month = 1 if cursor.month == 12 else cursor.month + 1
        cursor = date(year, month, 1)
    return months


def compliance_score(row: dict[str, Any]) -> float:
    """0–100 composite: pre-start rate, PDF rate, inspection cleanliness, no fatigue, no late submits."""
    rs = row.get("runsheets", 0)
    pre = row.get("prestarts", 0)
    pdfs = row.get("runsheetPdfs", 0)
    insp = row.get("inspection", 0)
    fat = row.get("fatigue", 0)
    late = row.get("lateSubmissions", 0)
    prestart_rate = pre / rs if rs else 0
    pdf_rate = pdfs / rs if rs else 0
    insp_rate = 1 - min(insp / (pre * len(INSPECTION_FIELD_IDS)), 1) if pre else 1
    ffd_rate = 1 - min(fat / (pre * len(FIT_FOR_DUTY_FIELD_IDS)), 1) if pre else 1
    late_rate = 1 - min(late / pre, 1) if pre else 1
    score = (prestart_rate * 0.3 + pdf_rate * 0.25 + insp_rate * 0.2 + ffd_rate * 0.15 + late_rate * 0.1) * 100
    return round(score, 1)


def finalize_detail(detail: dict[str, Any]) -> dict[str, Any]:
    drivers = list(detail["drivers"].values())
    for row in drivers:
        for key in ("loads", "hours", "waiting"):
            row[key] = round(row[key], 1)
        row["complianceScore"] = compliance_score(row)
        row["loadsPerHour"] = round(row["loads"] / row["hours"], 2) if row["hours"] else 0.0
        row["prestartRate"] = round(row["prestarts"] / row["runsheets"] * 100, 1) if row["runsheets"] else 0.0
        row["pdfRate"] = round(row["runsheetPdfs"] / row["runsheets"] * 100, 1) if row["runsheets"] else 0.0
    drivers.sort(key=lambda row: (row["runsheetPdfs"], row["loads"], row["hours"], row["prestarts"]), reverse=True)

    vehicles = []
    for row in detail["vehicles"].values():
        item = dict(row)
        item["drivers"] = len(item["drivers"])
        item["regoExpiry"] = REGO_EXPIRY.get(item["rego"].upper(), "")
        vehicles.append(item)
    vehicles.sort(key=lambda row: (row["serviceDue"] + row["stickerMissing"] + row["inspection"], row["prestarts"]), reverse=True)

    sites = []
    for row in detail["sites"].values():
        item = dict(row)
        item["drivers"] = len(item["drivers"])
        item["loads"] = round(item["loads"], 1)
        item["hours"] = round(item["hours"], 1)
        sites.append(item)
    sites.sort(key=lambda row: row["runsheets"], reverse=True)

    missing = []
    driver_names = {key: row["driver"] for key, row in detail["drivers"].items()}
    for driver_key, dates in detail["runsheetDates"].items():
        missing_dates = dates - detail["prestartDates"].get(driver_key, set())
        if not missing_dates:
            continue
        driver_name = driver_names.get(driver_key, driver_key)
        row = detail["drivers"].get(driver_key)
        if not row:
            continue
        missing.append(
            {
                "driver": driver_name,
                "runsheetPdfs": row["runsheetPdfs"],
                "loads": round(row["loads"], 1),
                "hours": round(row["hours"], 1),
                "missingDays": len(missing_dates),
                "missingDates": [day.strftime("%d/%m/%Y") for day in sorted(missing_dates)],
            }
        )
    missing.sort(key=lambda row: (row["missingDays"], row["runsheetPdfs"], row["loads"]), reverse=True)
    missing_by_driver = {row["driver"]: row for row in missing}

    exceptions = []
    for row in drivers:
        if row["missingAttachments"]:
            exceptions.append(
                {
                    "type": "Missing attachment",
                    "name": row["driver"],
                    "detail": f"{row['missingAttachments']} runsheet submissions without attachments",
                    "count": row["missingAttachments"],
                    "severity": "bad",
                }
            )
        missing_prestart = missing_by_driver.get(row["driver"])
        if missing_prestart:
            exceptions.append(
                {
                    "type": "Missing pre-start",
                    "name": row["driver"],
                    "detail": (
                        f"{missing_prestart['missingDays']} missing day(s): "
                        f"{', '.join(missing_prestart['missingDates'])}"
                    ),
                    "count": missing_prestart["missingDays"],
                    "severity": "bad",
                }
            )
        if row["serviceDue"]:
            exceptions.append(
                {
                    "type": "Service due",
                    "name": row["driver"],
                    "detail": f"{row['serviceDue']} service due declaration(s)",
                    "count": row["serviceDue"],
                    "severity": "warn",
                }
            )
        if row["inspection"]:
            exceptions.append(
                {
                    "type": "Inspection",
                    "name": row["driver"],
                    "detail": f"{row['inspection']} inspection exception answer(s)",
                    "count": row["inspection"],
                    "severity": "warn",
                }
            )
        if row["fatigue"]:
            exceptions.append(
                {
                    "type": "Fatigue",
                    "name": row["driver"],
                    "detail": f"{row['fatigue']} fit-for-duty declaration exception(s)",
                    "count": row["fatigue"],
                    "severity": "bad",
                }
            )
    for row in vehicles:
        if row["stickerMissing"]:
            exceptions.append(
                {
                    "type": "Missing sticker",
                    "name": row["rego"],
                    "detail": f"{row['stickerMissing']} missing service sticker response(s)",
                    "count": row["stickerMissing"],
                    "severity": "warn",
                }
            )
    exceptions.sort(key=lambda row: (0 if row["severity"] == "bad" else 1, -row["count"], row["type"], row["name"]))

    top_insp = sorted(
        [{"item": k, "count": v} for k, v in detail["inspectionBreakdown"].items() if v > 0],
        key=lambda x: -x["count"],
    )[:10]
    top_ffd = sorted(
        [{"item": k, "count": v} for k, v in detail["ffdBreakdown"].items() if v > 0],
        key=lambda x: -x["count"],
    )[:8]

    return {
        "label": detail["label"],
        "start": detail["start"],
        "end": detail["end"],
        "drivers": drivers[:30],
        "vehicles": vehicles[:30],
        "sites": sites,
        "missingPrestarts": missing[:30],
        "exceptions": exceptions[:80],
        "topInspectionFails": top_insp,
        "topFfdFails": top_ffd,
    }


def build_dashboard() -> dict[str, Any]:
    load_dotenv(ENV_PATH)
    api_key = require_env("JOTFORM_API_KEY")
    runsheet_form_id = require_env("JOTFORM_FORM_ID")
    prestart_form_id = require_env("PRESTART_FORM_ID")

    today = datetime.now().date()
    data_start = date(2026, 4, 1)
    data_end = today
    fetch_start = data_start - timedelta(days=14)
    earliest_week_ending = week_ending_for(data_start)
    latest_week_ending = week_ending_for(data_end)
    week_endings = []
    cursor = earliest_week_ending
    while cursor <= latest_week_ending:
        week_endings.append(cursor)
        cursor += timedelta(days=7)

    weeks = {ending: empty_week(ending) for ending in week_endings}
    for week in weeks.values():
        if week["weekStart"] < data_start.isoformat():
            week["weekStart"] = data_start.isoformat()
    month_starts = iter_month_starts(data_start, data_end)
    months = {month: empty_month(month) for month in month_starts}
    period_details = {
        "week": {
            ending.isoformat(): empty_detail(
                f"WE {ending.strftime('%d/%m/%Y')}",
                max(ending - timedelta(days=6), data_start),
                min(ending, data_end),
            )
            for ending in week_endings
        },
        "month": {
            month.strftime("%Y-%m"): empty_detail(
                month.strftime("%b %Y"),
                month,
                min(
                    date(month.year + (1 if month.month == 12 else 0), 1 if month.month == 12 else month.month + 1, 1)
                    - timedelta(days=1),
                    data_end,
                ),
            )
            for month in month_starts
        },
    }
    runsheet_drivers: dict[date, set[str]] = defaultdict(set)
    prestart_drivers: dict[date, set[str]] = defaultdict(set)
    runsheet_pdf_keys: dict[date, set[tuple[str, date]]] = defaultdict(set)
    monthly_runsheet_drivers: dict[date, set[str]] = defaultdict(set)
    monthly_prestart_drivers: dict[date, set[str]] = defaultdict(set)
    monthly_runsheet_pdf_keys: dict[date, set[tuple[str, date]]] = defaultdict(set)
    driver_rollup: dict[str, dict[str, Any]] = {}
    driver_week_rollup: dict[tuple[date, str], dict[str, Any]] = {}
    missing_prestarts: dict[str, dict[str, Any]] = {}
    vehicle_rollup: dict[str, dict[str, Any]] = {}
    site_rollup: dict[str, dict[str, Any]] = {}
    driver_prestart_dates: dict[str, set[date]] = defaultdict(set)
    runsheet_driver_dates: dict[str, set[date]] = defaultdict(set)

    runsheets = fetch_submissions(api_key, runsheet_form_id, fetch_start)
    prestarts = fetch_submissions(api_key, prestart_form_id, fetch_start)

    for submission in runsheets:
        shift_date = parse_date(answer(submission, RUNSHEET_FIELDS.date))
        if shift_date is None or shift_date < data_start or shift_date > data_end:
            continue
        ending = week_ending_for(shift_date)
        if ending not in weeks:
            continue
        month_key = month_start_for(shift_date)
        driver = text_value(answer(submission, RUNSHEET_FIELDS.driver)) or "Unknown"
        driver_id = driver_key(driver)
        site = text_value(answer(submission, RUNSHEET_FIELDS.site)) or "Unassigned"
        attachments = attachment_count(answer(submission, RUNSHEET_FIELDS.attachments))
        loads = number_value(answer(submission, RUNSHEET_FIELDS.loads)) or attachments
        hours = number_value(answer(submission, RUNSHEET_FIELDS.hours))
        waiting = number_value(answer(submission, RUNSHEET_FIELDS.waiting))
        shift_timing = text_value(answer(submission, SHIFT_TIMING_FIELD_ID)).upper()
        is_night = "NIGHT" in shift_timing

        week = weeks[ending]
        week["runsheets"] += 1
        week["loads"] += loads
        week["hours"] += hours
        week["waiting"] += waiting
        week["nightShifts" if is_night else "dayShifts"] += 1
        month = months[month_key]
        month["runsheets"] += 1
        month["loads"] += loads
        month["hours"] += hours
        month["waiting"] += waiting
        month["nightShifts" if is_night else "dayShifts"] += 1
        if attachments <= 0:
            week["missingAttachments"] += 1
            month["missingAttachments"] += 1
        else:
            runsheet_pdf_keys[ending].add((driver_id, shift_date))
            monthly_runsheet_pdf_keys[month_key].add((driver_id, shift_date))
        runsheet_drivers[ending].add(driver_id)
        monthly_runsheet_drivers[month_key].add(driver_id)
        runsheet_driver_dates[driver_id].add(shift_date)

        current = driver_rollup.setdefault(driver_id, empty_driver_row(driver))
        current_week = driver_week_rollup.setdefault(
            (ending, driver_id),
            empty_driver_row(driver),
        )
        week_detail = period_details["week"][ending.isoformat()]
        month_detail = period_details["month"][month_key.strftime("%Y-%m")]
        week_driver = week_detail["drivers"].setdefault(driver_id, empty_driver_row(driver))
        month_driver = month_detail["drivers"].setdefault(driver_id, empty_driver_row(driver))
        for target in (current, current_week, week_driver, month_driver):
            target["runsheets"] += 1
            target["runsheetPdfs"] += 1 if attachments > 0 else 0
            target["loads"] += loads
            target["hours"] += hours
            target["waiting"] += waiting
            target["missingAttachments"] += 1 if attachments <= 0 else 0
        week_detail["runsheetDates"][driver_id].add(shift_date)
        month_detail["runsheetDates"][driver_id].add(shift_date)

        site_item = site_rollup.setdefault(site, {"site": site, "runsheets": 0, "drivers": set(), "loads": 0, "hours": 0})
        site_item["runsheets"] += 1
        site_item["drivers"].add(driver_id)
        site_item["loads"] += loads
        site_item["hours"] += hours
        for detail in (week_detail, month_detail):
            detail_site = detail["sites"].setdefault(site, {"site": site, "runsheets": 0, "drivers": set(), "loads": 0, "hours": 0})
            detail_site["runsheets"] += 1
            detail_site["drivers"].add(driver_id)
            detail_site["loads"] += loads
            detail_site["hours"] += hours

    for submission in prestarts:
        shift_date = parse_date(answer(submission, PRESTART_FIELDS.date))
        if shift_date is None or shift_date < data_start or shift_date > data_end:
            continue
        ending = week_ending_for(shift_date)
        if ending not in weeks:
            continue
        month_key = month_start_for(shift_date)
        driver = text_value(answer(submission, PRESTART_FIELDS.driver)) or "Unknown"
        driver_id = driver_key(driver)
        rego = text_value(answer(submission, PRESTART_FIELDS.rego)) or "Unknown"
        service_due = 1 if is_checked(answer(submission, SERVICE_DUE_FIELD_ID)) else 0
        sticker_missing = 1 if is_no(answer(submission, SERVICE_STICKER_FIELD_ID)) else 0
        vcr = 1 if is_yes(answer(submission, VCR_FIELD_ID)) else 0
        defect_upload = 1 if attachment_count(answer(submission, DEFECT_PROOF_FIELD_ID)) else 0
        inspection_exceptions = sum(1 for field_id in INSPECTION_FIELD_IDS if is_no(answer(submission, field_id)))
        fatigue_exceptions = sum(1 for field_id in FIT_FOR_DUTY_FIELD_IDS if not is_checked(answer(submission, field_id)))
        failed_items = {INSPECTION_LABELS[fid] for fid in INSPECTION_FIELD_IDS if is_no(answer(submission, fid))}
        failed_ffd = {FFD_LABELS[fid] for fid in FIT_FOR_DUTY_FIELD_IDS if not is_checked(answer(submission, fid))}
        odometer = int(number_value(answer(submission, ODOMETER_FIELD_ID))) or None
        service_remaining_km = int(number_value(answer(submission, SERVICE_REMAINING_KM_FIELD_ID))) or None

        # Submission timeliness: hours between shift date and form created_at
        created_at = parse_datetime(submission.get("created_at"))
        submission_delay_hours: float | None = None
        late_submission = 0
        if created_at and shift_date:
            shift_dt = datetime.combine(shift_date, datetime.min.time())
            delay_h = (created_at - shift_dt).total_seconds() / 3600
            if 0 <= delay_h <= 72:  # ignore outliers / backdated entries
                submission_delay_hours = round(delay_h, 1)
                late_submission = 1 if delay_h > 24 else 0

        week = weeks[ending]
        week["prestarts"] += 1
        week["serviceDue"] += service_due
        week["stickerMissing"] += sticker_missing
        week["vcrRaised"] += vcr
        week["defectProofUploads"] += defect_upload
        week["inspectionExceptions"] += inspection_exceptions
        week["fatigueExceptions"] += fatigue_exceptions
        for item in failed_items:
            week["inspectionBreakdown"][item] += 1
        for item in failed_ffd:
            week["ffdBreakdown"][item] += 1
        if submission_delay_hours is not None:
            week["submissionDelayTotal"] += submission_delay_hours
            week["submissionDelayCount"] += 1

        month = months[month_key]
        month["prestarts"] += 1
        month["serviceDue"] += service_due
        month["stickerMissing"] += sticker_missing
        month["vcrRaised"] += vcr
        month["defectProofUploads"] += defect_upload
        month["inspectionExceptions"] += inspection_exceptions
        month["fatigueExceptions"] += fatigue_exceptions
        for item in failed_items:
            month["inspectionBreakdown"][item] += 1
        for item in failed_ffd:
            month["ffdBreakdown"][item] += 1
        if submission_delay_hours is not None:
            month["submissionDelayTotal"] += submission_delay_hours
            month["submissionDelayCount"] += 1

        prestart_drivers[ending].add(driver_id)
        monthly_prestart_drivers[month_key].add(driver_id)
        driver_prestart_dates[driver_id].add(shift_date)

        current = driver_rollup.setdefault(driver_id, empty_driver_row(driver))
        current_week = driver_week_rollup.setdefault((ending, driver_id), empty_driver_row(driver))
        week_detail = period_details["week"][ending.isoformat()]
        month_detail = period_details["month"][month_key.strftime("%Y-%m")]
        week_driver = week_detail["drivers"].setdefault(driver_id, empty_driver_row(driver))
        month_driver = month_detail["drivers"].setdefault(driver_id, empty_driver_row(driver))
        for target in (current, current_week, week_driver, month_driver):
            target["prestarts"] += 1
            target["serviceDue"] += service_due
            target["inspection"] += inspection_exceptions
            target["fatigue"] += fatigue_exceptions
            target["vcr"] += vcr
            target["lateSubmissions"] += late_submission
        week_detail["prestartDates"][driver_id].add(shift_date)
        month_detail["prestartDates"][driver_id].add(shift_date)
        for item in failed_items:
            week_detail["inspectionBreakdown"][item] += 1
            month_detail["inspectionBreakdown"][item] += 1
        for item in failed_ffd:
            week_detail["ffdBreakdown"][item] += 1
            month_detail["ffdBreakdown"][item] += 1

        vehicle = vehicle_rollup.setdefault(
            rego,
            {"rego": rego, "prestarts": 0, "drivers": set(), "serviceDue": 0, "stickerMissing": 0, "inspection": 0, "vcrCount": 0, "latestOdometer": None, "serviceKmsRemaining": None},
        )
        vehicle["prestarts"] += 1
        vehicle["drivers"].add(driver_id)
        vehicle["serviceDue"] += service_due
        vehicle["stickerMissing"] += sticker_missing
        vehicle["inspection"] += inspection_exceptions
        vehicle["vcrCount"] += vcr
        if odometer and (vehicle["latestOdometer"] is None or odometer > vehicle["latestOdometer"]):
            vehicle["latestOdometer"] = odometer
        if service_remaining_km is not None:
            vehicle["serviceKmsRemaining"] = service_remaining_km

        for detail in (week_detail, month_detail):
            detail_vehicle = detail["vehicles"].setdefault(
                rego,
                {"rego": rego, "prestarts": 0, "drivers": set(), "serviceDue": 0, "stickerMissing": 0, "inspection": 0, "vcrCount": 0, "latestOdometer": None, "serviceKmsRemaining": None},
            )
            detail_vehicle["prestarts"] += 1
            detail_vehicle["drivers"].add(driver_id)
            detail_vehicle["serviceDue"] += service_due
            detail_vehicle["stickerMissing"] += sticker_missing
            detail_vehicle["inspection"] += inspection_exceptions
            detail_vehicle["vcrCount"] += vcr
            if odometer and (detail_vehicle["latestOdometer"] is None or odometer > detail_vehicle["latestOdometer"]):
                detail_vehicle["latestOdometer"] = odometer
            if service_remaining_km is not None:
                detail_vehicle["serviceKmsRemaining"] = service_remaining_km

    def compute_rates(p: dict[str, Any]) -> None:
        rs = p["runsheets"]
        pre = p["prestarts"]
        pdfs = p["runsheetPdfs"]
        hrs = p["hours"]
        p["prestartRate"] = round(pre / rs * 100, 1) if rs else 0.0
        p["pdfRate"] = round(pdfs / rs * 100, 1) if rs else 0.0
        p["loadsPerHour"] = round(p["loads"] / hrs, 2) if hrs else 0.0
        p["avgHoursPerShift"] = round(hrs / rs, 1) if rs else 0.0
        p["avgWaitPerShift"] = round(p["waiting"] / rs, 1) if rs else 0.0
        p["avgSubmissionDelay"] = round(
            p["submissionDelayTotal"] / p["submissionDelayCount"], 1
        ) if p.get("submissionDelayCount") else None
        # Sort inspection breakdown descending, keep top 10
        p["topInspectionFails"] = sorted(
            [{"item": k, "count": v} for k, v in p["inspectionBreakdown"].items() if v > 0],
            key=lambda x: -x["count"],
        )[:10]
        p["topFfdFails"] = sorted(
            [{"item": k, "count": v} for k, v in p["ffdBreakdown"].items() if v > 0],
            key=lambda x: -x["count"],
        )[:8]
        del p["inspectionBreakdown"], p["ffdBreakdown"]
        del p["submissionDelayTotal"], p["submissionDelayCount"]

    for ending, week in weeks.items():
        week["runsheetDrivers"] = len(runsheet_drivers[ending])
        week["prestartDrivers"] = len(prestart_drivers[ending])
        week["runsheetPdfs"] = len(runsheet_pdf_keys[ending])
        for key in ("loads", "hours", "waiting"):
            week[key] = round(week[key], 1)
        compute_rates(week)

    for month_key, month in months.items():
        month["runsheetDrivers"] = len(monthly_runsheet_drivers[month_key])
        month["prestartDrivers"] = len(monthly_prestart_drivers[month_key])
        month["runsheetPdfs"] = len(monthly_runsheet_pdf_keys[month_key])
        for key in ("loads", "hours", "waiting"):
            month[key] = round(month[key], 1)
        compute_rates(month)

    current_start = latest_week_ending - timedelta(days=6)
    current_driver_rows = [row for (ending, _), row in driver_week_rollup.items() if ending == latest_week_ending]
    current_driver_rows.sort(key=lambda row: (row["runsheetPdfs"], row["loads"], row["hours"]), reverse=True)
    current_driver_rows = current_driver_rows[:20]

    for driver_id, dates in runsheet_driver_dates.items():
        missing_dates = dates - driver_prestart_dates.get(driver_id, set())
        current_missing_dates = {d for d in missing_dates if current_start <= d <= data_end}
        if not current_missing_dates:
            continue
        row = driver_rollup.get(driver_id)
        week_row = driver_week_rollup.get((latest_week_ending, driver_id))
        if row and week_row:
            label = row["driver"]
            missing_prestarts[label] = {
                "driver": label,
                "runsheetPdfs": week_row["runsheetPdfs"],
                "loads": round(week_row["loads"], 1),
                "hours": round(week_row["hours"], 1),
                "missingDays": len(current_missing_dates),
                "missingDates": [day.strftime("%d/%m/%Y") for day in sorted(current_missing_dates)],
            }

    vehicle_rows = []
    for row in vehicle_rollup.values():
        row = dict(row)
        row["drivers"] = len(row["drivers"])
        row["regoExpiry"] = REGO_EXPIRY.get(row["rego"].upper(), "")
        vehicle_rows.append(row)
    vehicle_rows.sort(key=lambda row: (row["serviceDue"] + row["stickerMissing"] + row["inspection"] + row["vcrCount"], row["prestarts"]), reverse=True)

    site_rows = []
    for row in site_rollup.values():
        row = dict(row)
        row["drivers"] = len(row["drivers"])
        row["loads"] = round(row["loads"], 1)
        row["hours"] = round(row["hours"], 1)
        site_rows.append(row)
    site_rows.sort(key=lambda row: row["runsheets"], reverse=True)
    finalized_details = {
        kind: {key: finalize_detail(detail) for key, detail in details.items()}
        for kind, details in period_details.items()
    }

    return {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "dataStart": data_start.isoformat(),
        "dataEnd": data_end.isoformat(),
        "weeks": [weeks[ending] for ending in week_endings],
        "monthlySummary": [months[month] for month in month_starts],
        "currentDriverRows": current_driver_rows,
        "missingPrestarts": list(missing_prestarts.values()),
        "vehicleRows": vehicle_rows[:20],
        "siteRows": site_rows,
        "periodDetails": finalized_details,
    }


def js_const(name: str, value: Any) -> str:
    return f"    const {name} = {json.dumps(value, indent=6)};"


def replace_const(source: str, name: str, value: Any) -> str:
    # \n\s* before ]; is optional to handle empty arrays written as []
    pattern = re.compile(rf"    const {name} = \[.*?(?:\n\s*)?\];", re.DOTALL)
    replacement = js_const(name, value)
    source, count = pattern.subn(replacement, source, count=1)
    if count != 1:
        raise RuntimeError(f"Could not replace const {name}")
    return source


def replace_or_insert_const(source: str, name: str, value: Any, after_name: str) -> str:
    pattern = re.compile(rf"    const {name} = \[.*?(?:\n\s*)?\];", re.DOTALL)
    replacement = js_const(name, value)
    source, count = pattern.subn(replacement, source, count=1)
    if count == 1:
        return source

    after_pattern = re.compile(rf"(    const {after_name} = \[.*?\n\s*\];)", re.DOTALL)
    source, count = after_pattern.subn(rf"\1\n\n{replacement}", source, count=1)
    if count != 1:
        raise RuntimeError(f"Could not insert const {name} after {after_name}")
    return source


def replace_or_insert_string_const(source: str, name: str, value: str, marker: str) -> str:
    replacement = f'    const {name} = "{value}";'
    pattern = re.compile(rf"    const {name} = \".*?\";")
    source, count = pattern.subn(replacement, source, count=1)
    if count == 1:
        return source

    marker_index = source.find(marker)
    if marker_index == -1:
        raise RuntimeError(f"Could not find marker to insert const {name}")
    insert_at = source.find("\n", marker_index)
    if insert_at == -1:
        raise RuntimeError(f"Could not insert const {name}")
    return source[: insert_at + 1] + replacement + "\n" + source[insert_at + 1 :]


def replace_or_insert_object_const(source: str, name: str, value: Any, after_name: str) -> str:
    replacement = f"    const {name} = {json.dumps(value, indent=6)};"
    pattern = re.compile(rf"    const {name} = \{{.*?\n\s*\}};", re.DOTALL)
    source, count = pattern.subn(replacement, source, count=1)
    if count == 1:
        return source

    after_pattern = re.compile(rf"(    const {after_name} = \[.*?\n\s*\];)", re.DOTALL)
    source, count = after_pattern.subn(rf"\1\n\n{replacement}", source, count=1)
    if count != 1:
        raise RuntimeError(f"Could not insert const {name} after {after_name}")
    return source


def main() -> int:
    dashboard = build_dashboard()
    index_path = ROOT / "index.html"
    source = index_path.read_text(encoding="utf-8")
    source = replace_const(source, "weeks", dashboard["weeks"])
    source = replace_or_insert_const(source, "monthlySummary", dashboard["monthlySummary"], "weeks")
    source = replace_const(source, "currentDriverRows", dashboard["currentDriverRows"])
    source = replace_const(source, "missingPrestarts", dashboard["missingPrestarts"])
    source = replace_const(source, "vehicleRows", dashboard["vehicleRows"])
    source = replace_const(source, "siteRows", dashboard["siteRows"])
    source = replace_or_insert_object_const(source, "periodDetails", dashboard["periodDetails"], "siteRows")
    source = replace_or_insert_string_const(
        source,
        "dashboardDataEnd",
        dashboard["dataEnd"],
        "const dateFmt = value =>",
    )
    source = re.sub(
        r"Last (?:five|\d+) weeks by shift date",
        f"Last {len(dashboard['weeks'])} weeks by shift date",
        source,
        count=1,
    )
    source = re.sub(
        r"(?:April partial vs May month-to-date from available sample data|Two-month snapshot generated [^<]+|March-April-May snapshot generated [^<]+)",
        f"March-April-May snapshot generated {dashboard['generatedAt']}",
        source,
        count=1,
    )
    source = re.sub(
        r"title=\"Static (?:sample data only|two-month Jotform snapshot)\">(?:Static snapshot|2-month snapshot)</button>",
        f"title=\"Static March-April-May Jotform snapshot\">Mar-May snapshot</button>",
        source,
        count=1,
    )
    index_path.write_text(source, encoding="utf-8")
    print(
        f"Updated dashboard: {len(dashboard['weeks'])} weeks, "
        f"{dashboard['dataStart']} to {dashboard['dataEnd']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
