import os
import json
import glob
from core.client_config import get_client, resolve_duplicate_json_path

def extract_chart_of_accounts(client_name):
    """
    Harvests all unique whitelisted Tally ledger names for a given client from:
    1. The client's rules/*.json files
    2. The client's exports/Transactions.json (historical vouchers)
    """
    ledgers = set()
    
    # 1. Harvest from clients.json bank rules config and description_rules directories
    try:
        client = get_client(client_name)
        # Scan the rules folder relative to the client structure
        rules_dir = f"./rules/{client_name}"
        if os.path.exists(rules_dir):
            for file_path in glob.glob(os.path.join(rules_dir, "*.json")):
                if "ignored_descriptions" in file_path:
                    continue
                try:
                    with open(file_path, "r", encoding="utf-8-sig") as f:
                        rules = json.load(f)
                        if isinstance(rules, list):
                            for r in rules:
                                if isinstance(r, dict):
                                    for key in ("ledger", "payment_ledger", "receipt_ledger", "out_ledger", "in_ledger"):
                                        val = r.get(key)
                                        if val:
                                            ledgers.add(str(val).strip())
                                    # Harvest from amount_tiers
                                    if "amount_tiers" in r and isinstance(r["amount_tiers"], list):
                                        for tier in r["amount_tiers"]:
                                            for key in ("ledger", "payment_ledger", "receipt_ledger", "out_ledger", "in_ledger"):
                                                val = tier.get(key)
                                                if val:
                                                    ledgers.add(str(val).strip())
                except Exception as e:
                    print(f"Warning: Error reading rule file {file_path}: {e}")
    except Exception as e:
        print(f"Warning: Error accessing rules for {client_name}: {e}")

    # 2. Harvest from exports/Transactions.json (historical vouchers)
    try:
        dup_path = resolve_duplicate_json_path(client_name)
        if os.path.exists(dup_path):
            with open(dup_path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
                # Structure: { "lvbody": { "dspvchdetail": [ { "dspvchledaccount": "...", ... } ] } }
                vch_list = data.get("lvbody", {}).get("dspvchdetail", [])
                if not isinstance(vch_list, list):
                    vch_list = []
                for entry in vch_list:
                    led_acc = entry.get("dspvchledaccount")
                    if led_acc:
                        ledgers.add(str(led_acc).strip())
    except Exception as e:
        print(f"Warning: Error harvesting ledgers from Transaction JSON: {e}")

    return sorted(list(ledgers))
