import os
import json
import re
from datetime import datetime
from collections import defaultdict
from core.client_config import resolve_duplicate_json_path

COMMON_STOP_WORDS = {
    "neft", "rtgs", "upi", "imps", "transfer", "to", "by", "for", "payment",
    "dr", "cr", "cash", "withdrawal", "deposit", "txn", "ref", "in", "out",
    "charges", "commission", "chg", "chgs", "value", "date", "withdrawn", "deposited"
}

def parse_date(date_str):
    """Safely parse various date formats into a standard date object or string YYYY-MM-DD"""
    if not date_str:
        return None
    for fmt in ("%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None

def clean_description_tokens(description):
    """Splits description into meaningful candidate keywords by removing stop words and non-alpha tokens"""
    # Replace non-alphanumeric chars with spaces
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', description.lower())
    tokens = cleaned.split()
    candidates = []
    for token in tokens:
        token = token.strip()
        if len(token) > 2 and token not in COMMON_STOP_WORDS and not token.isdigit():
            candidates.append(token.upper())
    return candidates

def mine_rules_from_history(unclassified_txns, client_name):
    """
    Mines ledger mappings by matching unclassified statement rows with Tally history (Transactions.json)
    on Date and Amount, then extracting high-correlation keyword patterns.
    """
    dup_path = resolve_duplicate_json_path(client_name)
    if not os.path.exists(dup_path):
        return []

    # 1. Load historical Tally vouchers
    tally_vouchers = []
    try:
        with open(dup_path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
            tally_vouchers = data.get("lvbody", {}).get("dspvchdetail", [])
    except Exception as e:
        print(f"Warning: Failed to load duplicate JSON path: {e}")
        return []

    # Index historical vouchers by (Date, Amount) for fast lookup
    # Note: Tally amounts are strings, we convert to float
    history_index = defaultdict(list)
    for vch in tally_vouchers:
        date_parsed = parse_date(vch.get("dspvchdate"))
        if not date_parsed:
            continue
        
        # In duplicate JSON, amounts can be in dspvchdramt or dspvchcramt
        try:
            dr_amt = float(str(vch.get("dspvchdramt", "0")).replace(",", "").strip() or 0)
            cr_amt = float(str(vch.get("dspvchcramt", "0")).replace(",", "").strip() or 0)
            amt = dr_amt if dr_amt > 0 else cr_amt
            if amt > 0:
                history_index[(date_parsed, round(amt, 2))].append(vch.get("dspvchledaccount"))
        except (ValueError, TypeError):
            continue

    # 2. Match unclassified rows against history
    keyword_ledger_counts = defaultdict(lambda: defaultdict(int))
    total_keyword_matches = defaultdict(int)

    for txn in unclassified_txns:
        txn_date = parse_date(txn.get("date"))
        if not txn_date:
            continue
        
        try:
            txn_amt = float(str(txn.get("amount", "0")).replace(",", "").strip() or 0)
        except ValueError:
            continue

        if txn_amt <= 0:
            continue

        # Look for exact date and amount matches
        matched_ledgers = history_index.get((txn_date, round(txn_amt, 2)), [])
        if not matched_ledgers:
            continue

        # Extract tokens from transaction description
        tokens = clean_description_tokens(txn.get("description", ""))
        for token in tokens:
            for ledger in matched_ledgers:
                keyword_ledger_counts[token][ledger] += 1
                total_keyword_matches[token] += 1

    # 3. Formulate proposed rules where confidence >= 85%
    proposed_rules = []
    for token, ledgers in keyword_ledger_counts.items():
        total = total_keyword_matches[token]
        if total < 2:  # Must have matched at least twice to be a reliable pattern
            continue
        
        # Sort ledgers by match count descending
        sorted_ledgers = sorted(ledgers.items(), key=lambda x: x[1], reverse=True)
        top_ledger, count = sorted_ledgers[0]
        confidence_ratio = count / total

        if confidence_ratio >= 0.85:
            proposed_rules.append({
                "pattern": [token],
                "ledger": top_ledger,
                "priority": 90,
                "confidence": "HIGH" if confidence_ratio >= 0.95 else "MEDIUM",
                "reasoning": f"Mined from historical Tally match: {count}/{total} instances mapped to '{top_ledger}'"
            })

    return proposed_rules
