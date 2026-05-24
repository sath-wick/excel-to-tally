import json
from collections import Counter, defaultdict
from datetime import datetime
import pandas as pd


_DATE_FORMATS = (
    "%d-%m-%Y",
    "%d-%m-%y",
    "%d-%b-%y",
    "%d-%b-%Y",
    "%d-%B-%y",
    "%d-%B-%Y",
    "%d %b %y",
    "%d %b %Y",
    "%d %B %y",
    "%d %B %Y",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d/%m/%y",
)

JSON_DUPLICATE_COLUMNS = [
    "TxnJson_VoucherType",
    "TxnJson_Date",
    "TxnJson_Ledger",
    "TxnJson_DrAmt",
    "TxnJson_CrAmt",
    "TxnJson_VoucherNo",
]


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
        return round(abs(float(text)), 2)
    except ValueError:
        return None


def _extract_voucher_details(payload):
    details = payload.get("lvbody", {}).get("dspvchdetail", [])

    if isinstance(details, dict):
        return [details]
    if isinstance(details, list):
        return details
    return []


def _load_json_payload(json_path):
    with open(json_path, "rb") as file_obj:
        raw_bytes = file_obj.read()

    for encoding in ("utf-8-sig", "utf-16", "utf-16-le", "utf-16-be"):
        try:
            return json.loads(raw_bytes.decode(encoding))
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue

    return json.loads(raw_bytes.decode("latin-1"))


def _to_signed_float(value):
    if value is None:
        return None

    text = str(value).replace(",", "").strip()
    text = text.replace("(", "-").replace(")", "")
    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def _normalize_side(value):
    signed_amount = _to_signed_float(value)
    if signed_amount is None:
        return None
    return "withdrawal" if signed_amount >= 0 else "deposit"


def _extract_signed_amount(entry):
    for field_name in ("dspvchdramt", "dspvchcramt"):
        value = entry.get(field_name)
        signed_amount = _to_signed_float(value)
        if signed_amount is not None:
            return value, field_name, signed_amount

    return None, None, None


def _derive_json_side(entry, amount_field, signed_amount):
    voucher_type = str(entry.get("dspvchtype", "")).strip().upper()

    if voucher_type in {"PYMT", "PAYMENT"}:
        return "withdrawal"
    if voucher_type in {"RCPT", "RECEIPT"}:
        return "deposit"

    if amount_field == "dspvchdramt":
        return "withdrawal"
    if amount_field == "dspvchcramt":
        return _normalize_side(signed_amount)

    return None


def _extract_text(value):
    if value is None:
        return ""

    if isinstance(value, dict):
        for preferred_key in ("dspexplvchnumber", "dspvchnumber"):
            if preferred_key in value:
                text = _extract_text(value.get(preferred_key))
                if text:
                    return text

        for child in value.values():
            text = _extract_text(child)
            if text:
                return text
        return ""

    if isinstance(value, list):
        for item in value:
            text = _extract_text(item)
            if text:
                return text
        return ""

    return str(value).strip()


def _format_json_duplicate_details(entry):
    return {
        "TxnJson_VoucherType": str(entry.get("dspvchtype", "")).strip(),
        "TxnJson_Date": str(entry.get("dspvchdate", "")).strip(),
        "TxnJson_Ledger": str(entry.get("dspvchledaccount", "")).strip(),
        "TxnJson_DrAmt": str(entry.get("dspvchdramt", "")).strip(),
        "TxnJson_CrAmt": str(entry.get("dspvchcramt", "")).strip(),
        "TxnJson_VoucherNo": _extract_text(entry.get("dspvchnumber")),
    }


def load_existing_transaction_entries(json_path):
    payload = _load_json_payload(json_path)
    records = []

    for index, entry in enumerate(_extract_voucher_details(payload), start=1):
        date_key = normalize_date(entry.get("dspvchdate"))
        ledger_key = normalize_ledger(entry.get("dspvchledaccount"))
        amount_source, amount_field, signed_amount = _extract_signed_amount(entry)
        amount_key = normalize_amount(amount_source)
        side_key = _derive_json_side(entry, amount_field, signed_amount)

        records.append(
            {
                "match_id": index,
                "date_key": date_key,
                "ledger_key": ledger_key,
                "amount_key": amount_key,
                "side_key": side_key,
                "json_details": _format_json_duplicate_details(entry),
            }
        )

    return records


class ExistingTransactionMatcher:
    def __init__(self, json_path):
        self.records = load_existing_transaction_entries(json_path)
        self.consumed_ids = set()

        self.by_ledger_key = defaultdict(list)
        self.by_side_key = defaultdict(list)
        self.by_amount_key = defaultdict(list)

        for record in self.records:
            date_key = record["date_key"]
            amount_key = record["amount_key"]
            ledger_key = record["ledger_key"]
            side_key = record["side_key"]

            if date_key is None or amount_key is None or amount_key == 0:
                continue

            self.by_amount_key[(date_key, amount_key)].append(record)

            if side_key is not None:
                self.by_side_key[(date_key, side_key, amount_key)].append(record)

            if ledger_key:
                self.by_ledger_key[(date_key, ledger_key, amount_key)].append(record)

    def _claim_first(self, candidates):
        for record in candidates:
            match_id = record["match_id"]
            if match_id in self.consumed_ids:
                continue
            self.consumed_ids.add(match_id)
            return record
        return None

    def find_statement_match(self, value_date, withdrawals, deposits):
        date_key = normalize_date(value_date)
        if date_key is None:
            return None

        withdrawal_amount = normalize_amount(withdrawals)
        deposit_amount = normalize_amount(deposits)

        statement_side = None
        amount_key = None

        if withdrawal_amount is not None and withdrawal_amount > 0:
            statement_side = "withdrawal"
            amount_key = withdrawal_amount
        elif deposit_amount is not None and deposit_amount > 0:
            statement_side = "deposit"
            amount_key = deposit_amount

        if amount_key is None:
            return None

        if statement_side is not None:
            side_record = self._claim_first(
                self.by_side_key.get((date_key, statement_side, amount_key), [])
            )
            if side_record:
                return side_record

        return self._claim_first(self.by_amount_key.get((date_key, amount_key), []))

    def find_voucher_match(self, voucher_date, voucher_amount, ledgers):
        date_key = normalize_date(voucher_date)
        amount_key = normalize_amount(voucher_amount)

        if date_key is None or amount_key is None or amount_key == 0:
            return None

        for ledger in ledgers:
            ledger_key = normalize_ledger(ledger)
            if not ledger_key:
                continue

            record = self._claim_first(
                self.by_ledger_key.get((date_key, ledger_key, amount_key), [])
            )
            if record:
                return record

        return None


def load_existing_contras(json_path):
    matcher = ExistingTransactionMatcher(json_path)
    existing = set()

    for record in matcher.records:
        date_key = record["date_key"]
        ledger_key = record["ledger_key"]
        amount_key = record["amount_key"]

        if date_key is None or not ledger_key or amount_key is None:
            continue

        existing.add((date_key, ledger_key, amount_key))

    return existing


def load_existing_contra_amount_counter(json_path):
    matcher = ExistingTransactionMatcher(json_path)
    side_amount_counter = Counter()
    amount_counter = Counter()

    for record in matcher.records:
        date_key = record["date_key"]
        amount_key = record["amount_key"]
        side_key = record["side_key"]

        if date_key is None or amount_key is None or amount_key == 0:
            continue

        amount_counter[(date_key, amount_key)] += 1
        if side_key is not None:
            side_amount_counter[(date_key, side_key, amount_key)] += 1

    return side_amount_counter, amount_counter


def split_statement_duplicates(statement_df, json_path):
    matcher = ExistingTransactionMatcher(json_path)

    if not matcher.by_amount_key:
        return statement_df.copy(), statement_df.iloc[0:0].copy()

    duplicate_indices = []
    duplicate_json_matches = []

    for index, row in statement_df.iterrows():
        matched_record = matcher.find_statement_match(
            row.get("Value Date"),
            row.get("Withdrawals"),
            row.get("Deposits"),
        )

        if not matched_record:
            continue

        duplicate_indices.append(index)
        duplicate_json_matches.append(matched_record["json_details"])

    duplicate_df = statement_df.loc[duplicate_indices].copy()
    filtered_df = statement_df.drop(index=duplicate_indices).copy()

    if not duplicate_df.empty:
        duplicate_meta_df = json_details_dataframe(duplicate_json_matches)
        duplicate_df.reset_index(drop=True, inplace=True)
        duplicate_df = duplicate_df.join(duplicate_meta_df)
    else:
        duplicate_df = statement_df.iloc[0:0].copy()
        for column in JSON_DUPLICATE_COLUMNS:
            duplicate_df[column] = ""

    filtered_df.reset_index(drop=True, inplace=True)
    return filtered_df, duplicate_df


def json_details_dataframe(records):
    if not records:
        return json_empty_dataframe()

    return json_to_dataframe(records).reindex(columns=JSON_DUPLICATE_COLUMNS)


def json_empty_dataframe():
    return pd.DataFrame(columns=JSON_DUPLICATE_COLUMNS)


def json_to_dataframe(records):
    return pd.DataFrame(records if records else [])
