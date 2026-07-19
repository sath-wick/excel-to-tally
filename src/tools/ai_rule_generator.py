import os
import sys
import argparse
import pandas as pd
import json

# Add project root and src root to sys.path so it works when run from any folder
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
src_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from core.client_config import get_client, resolve_rule_path, extract_bank_code
from tools.rule_generator.chart_of_accounts import extract_chart_of_accounts
from tools.rule_generator.gemini_engine import generate_rules_with_gemini
from tools.rule_generator.tally_miner import mine_rules_from_history
from tools.rule_generator.rule_applier import append_rules

def load_unclassified_from_excel(excel_path, bank_ledger):
    """
    Reads the Unclassified sheet from Statement_Import.xlsx and formats
    transactions as a list of dicts.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found at: {excel_path}")
        
    try:
        df = pd.read_excel(excel_path, sheet_name="Unclassified")
    except ValueError:
        print("No 'Unclassified' sheet found or file is empty.")
        return []

    if df.empty:
        return []

    # Map columns to transaction structure
    txns = []
    for _, row in df.iterrows():
        desc = str(row.get("Description", "")).strip()
        if not desc or desc == "nan":
            continue
            
        amt = 0.0
        try:
            amt = float(str(row.get("Amount", "0")).replace(",", "").strip() or 0)
        except ValueError:
            pass

        # Inferred direction based on ledger placements
        cr_led = str(row.get("Cr_Ledger", "")).strip()
        dr_led = str(row.get("Dr_Ledger", "")).strip()
        
        direction = "OUT" if cr_led.lower() == bank_ledger.lower() else "IN"
        
        txns.append({
            "date": str(row.get("Date", "")).strip(),
            "description": desc,
            "amount": amt,
            "direction": direction
        })
        
    return txns

def main():
    parser = argparse.ArgumentParser(description="Tally Framework AI & Historical Rule Generator")
    parser.add_argument("excel_path", nargs="?", default="./output/Statement_Import.xlsx", 
                        help="Path to Statement_Import.xlsx (default: ./output/Statement_Import.xlsx)")
    parser.add_argument("bank_ledger", nargs="?", help="Name of the bank ledger (e.g. 'YES 2085')")
    parser.add_argument("client_name", nargs="?", help="Client Name (e.g. 'Arjun Rao')")
    parser.add_argument("--mode", choices=["gemini", "mine", "both"], default="both", 
                        help="Rule generation mode: gemini, mine, or both (default: both)")
    parser.add_argument("--apply", action="store_true", help="Apply proposed rules automatically without prompt")
    parser.add_argument("--model", default="gemini-2.5-flash", help="Gemini Model to use (default: gemini-2.5-flash)")
    
    args = parser.parse_args()

    # Fallback/Interactive setup if args are missing
    client_name = args.client_name
    if not client_name:
        try:
            from core.client_config import get_default_client_name
            client_name = get_default_client_name()
            print(f"Defaulting to client: {client_name}")
        except Exception:
            client_name = input("Enter Client Name: ").strip()

    bank_ledger = args.bank_ledger
    if not bank_ledger:
        try:
            client = get_client(client_name)
            banks = client.get("banks", {})
            if banks:
                bank_ledger = list(banks.keys())[0]
                print(f"Defaulting to bank: {bank_ledger}")
        except Exception:
            bank_ledger = input("Enter Bank Ledger Name (e.g. YES 2085): ").strip()

    if not client_name or not bank_ledger:
        print("Error: Client name and bank ledger are required.")
        sys.exit(1)

    print(f"\n--- Loading Unclassified Transactions from {args.excel_path} ---")
    try:
        unclassified_txns = load_unclassified_from_excel(args.excel_path, bank_ledger)
    except Exception as e:
        print(f"Error reading Excel sheet: {e}")
        sys.exit(1)

    if not unclassified_txns:
        print("No unclassified transactions found to process.")
        sys.exit(0)

    print(f"Found {len(unclassified_txns)} unclassified transactions.")

    # 1. Harvest Chart of Accounts
    print("\n--- Harvesting whitelisted Tally Ledgers ---")
    coa = extract_chart_of_accounts(client_name)
    print(f"Loaded {len(coa)} unique ledgers in the Chart of Accounts.")

    proposed_rules = []

    # 2. Historical Mining
    if args.mode in ("mine", "both"):
        print("\n--- Running Historical Data Miner ---")
        mined_rules = mine_rules_from_history(unclassified_txns, client_name)
        print(f"Mined {len(mined_rules)} rules from historical Transactions.json.")
        proposed_rules.extend(mined_rules)

    # 3. Gemini Generation
    if args.mode in ("gemini", "both"):
        print("\n--- Requesting Gemini Pro AI Suggestions ---")
        try:
            gemini_rules = generate_rules_with_gemini(
                unclassified_txns[:80], # Limit to first 80 to fit token budgets nicely
                coa,
                model_name=args.model
            )
            print(f"Gemini suggested {len(gemini_rules)} rules.")
            proposed_rules.extend(gemini_rules)
        except Exception as e:
            print(f"Warning: Gemini AI rule generation failed: {e}")

    if not proposed_rules:
        print("No rules were proposed.")
        sys.exit(0)

    # Deduplicate proposed rules locally before applying
    unique_proposed = []
    seen_patterns = set()
    for rule in proposed_rules:
        patterns = tuple(sorted(rule.get("pattern", [])))
        if patterns not in seen_patterns:
            seen_patterns.add(patterns)
            unique_proposed.append(rule)

    print("\n================ PROPOSED NEW RULES ================")
    for idx, rule in enumerate(unique_proposed, 1):
        patterns = ", ".join(rule.get("pattern", []))
        ledger = rule.get("ledger")
        conf = rule.get("confidence", "HIGH")
        reason = rule.get("reasoning", "N/A")
        print(f"{idx}. [{conf}] Pattern: '{patterns}' -> Ledger: '{ledger}'")
        print(f"   Reason: {reason}")
        if "min_amount" in rule and rule["min_amount"] is not None:
            print(f"   Condition: Amount >= {rule['min_amount']}")
        if "max_amount" in rule and rule["max_amount"] is not None:
            print(f"   Condition: Amount <= {rule['max_amount']}")
        if "amount_tiers" in rule and rule["amount_tiers"]:
            print("   Condition Slabs:")
            for tier in rule["amount_tiers"]:
                print(f"     - Min: {tier.get('min')}, Max: {tier.get('max')} -> {tier.get('ledger')}")
    print("====================================================")

    apply = args.apply
    if not apply:
        choice = input(f"\nDo you want to append these {len(unique_proposed)} rules to rules file? [y/N]: ").strip().lower()
        apply = choice in ("y", "yes")

    if apply:
        rule_file_path = resolve_rule_path(bank_ledger, client_name)
        print(f"\nSaving rules to: {rule_file_path}")
        added, updated = append_rules(rule_file_path, unique_proposed)
        print(f"Success: Added {added} new rules, updated {updated} existing rules.")
    else:
        print("\nCancelled. No rules were modified.")

if __name__ == "__main__":
    main()
