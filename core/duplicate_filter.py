import json
from datetime import datetime


_DATE_FORMATS = (
    "%d-%m-%Y",
    "%d-%b-%y",
    "%d-%b-%Y",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d/%m/%y",
)


def normalize_date(value):
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    return None


def normalize_ledger(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_amount(value):
    if value is None:
        return None

    text = str(value).replace(",", "").strip()
    text = text.replace("(", "-").replace(")", "")
    if not text:
        return None

    try:
        return abs(float(text))
    except ValueError:
        return None


def _extract_contra_details(payload):
    details = payload.get("lvbody", {}).get("dspvchdetail", [])

    if isinstance(details, dict):
        return [details]
    if isinstance(details, list):
        return details
    return []


def load_existing_contras(json_path):
    payload = _load_json_payload(json_path)

    existing = set()

    for entry in _extract_contra_details(payload):
        voucher_type = str(entry.get("dspvchtype", "")).strip().upper()
        if voucher_type not in {"CNTRA", "CTRA"}:
            continue

        date_key = normalize_date(entry.get("dspvchdate"))
        ledger_key = normalize_ledger(entry.get("dspvchledaccount"))

        amount_value = entry.get("dspvchcramt")
        if amount_value is None:
            amount_value = entry.get("dspvchdramt")
        amount_key = normalize_amount(amount_value)

        if date_key is None or not ledger_key or amount_key is None:
            continue

        existing.add((date_key, ledger_key, amount_key))

    return existing


def _load_json_payload(json_path):
    with open(json_path, "rb") as file_obj:
        raw_bytes = file_obj.read()

    for encoding in ("utf-8-sig", "utf-16", "utf-16-le", "utf-16-be"):
        try:
            return json.loads(raw_bytes.decode(encoding))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue

    return json.loads(raw_bytes.decode("latin-1"))
