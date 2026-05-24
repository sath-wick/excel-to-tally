def apply_bank_specific_rule_override(rule, transaction, bank_ledger):
    """
    BANK 3332 / 2085 special rule:
    - Labour Charges with withdrawal >= 500 -> Vehicle Maintainance
    - Labour Charges with withdrawal < 500  -> Staff Welfare
    """
    if not _is_special_bank(bank_ledger):
        return rule

    if str(rule.get("voucher_type", "")).strip() != "Payment":
        return rule

    if _normalize_text(rule.get("ledger")) != "labour charges":
        return rule

    withdrawal_amount = _to_float(getattr(transaction, "withdrawal", 0))
    if withdrawal_amount <= 0:
        return rule

    overridden_rule = dict(rule)
    if withdrawal_amount >= 2000:
        overridden_rule["ledger"] = "Labour Charges"
    elif withdrawal_amount >= 500:
        overridden_rule["ledger"] = "Vehicle Maintainance"
    else:
        overridden_rule["ledger"] = "Staff Welfare"

    return overridden_rule


def _is_special_bank(bank_ledger):
    digits = "".join(ch for ch in str(bank_ledger) if ch.isdigit())
    return digits in {"3332", "2085"}


def _normalize_text(value):
    return str(value).strip().lower()


def _to_float(value):
    text = str(value).replace(",", "").strip()
    if not text:
        return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0
