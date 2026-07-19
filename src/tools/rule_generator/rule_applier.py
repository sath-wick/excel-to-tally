import os
import json
import shutil

def append_rules(rule_path, new_rules):
    """
    Safely merges, deduplicates, sorts, and writes rules to the specified JSON rules file.
    Creates a backup copy before modifications.
    """
    if not new_rules:
        return 0, 0

    existing_rules = []
    if os.path.exists(rule_path):
        # 1. Back up existing rules file
        backup_path = rule_path + ".bak"
        try:
            shutil.copyfile(rule_path, backup_path)
        except Exception as e:
            print(f"Warning: Failed to create backup of rules file: {e}")

        # 2. Read existing rules
        try:
            with open(rule_path, "r", encoding="utf-8-sig") as f:
                existing_rules = json.load(f)
                if not isinstance(existing_rules, list):
                    existing_rules = []
        except Exception as e:
            print(f"Warning: Failed to read existing rules file: {e}. Starting fresh.")
            existing_rules = []

    # Helper to clean/hash pattern arrays for comparison
    def get_pattern_set(rule_obj):
        patterns = rule_obj.get("pattern", [])
        if isinstance(patterns, str):
            patterns = [patterns]
        return {str(p).strip().lower() for p in patterns}

    # 3. Merge and deduplicate
    added_count = 0
    updated_count = 0
    
    for new_r in new_rules:
        new_pattern_set = get_pattern_set(new_r)
        if not new_pattern_set:
            continue
        
        # Check if any existing rule covers these exact patterns
        duplicate_found = False
        for idx, exist_r in enumerate(existing_rules):
            exist_pattern_set = get_pattern_set(exist_r)
            # If the sets overlap exactly or the new pattern set is a subset
            if new_pattern_set == exist_pattern_set:
                # Update ledger if changed/null in existing
                if exist_r.get("ledger") != new_r.get("ledger"):
                    existing_rules[idx]["ledger"] = new_r.get("ledger")
                    updated_count += 1
                duplicate_found = True
                break
        
        if not duplicate_found:
            # Clean up AI/miner specific keys like 'confidence' and 'reasoning'
            clean_rule = {
                "pattern": new_r.get("pattern"),
                "ledger": new_r.get("ledger"),
                "priority": new_r.get("priority", 100)
            }
            # Copy other optional standard fields if present
            for opt_key in ("payment_ledger", "receipt_ledger", "out_ledger", "in_ledger", 
                            "min_amount", "max_amount", "amount_tiers", "is_contra", "contra", "voucher_type"):
                if opt_key in new_r:
                    clean_rule[opt_key] = new_r[opt_key]
                    
            existing_rules.append(clean_rule)
            added_count += 1

    # 4. Sort all rules by priority descending
    existing_rules.sort(key=lambda r: r.get("priority", 0), reverse=True)

    # 5. Write back to rules file with UTF-8 BOM encoding
    try:
        os.makedirs(os.path.dirname(rule_path), exist_ok=True)
        with open(rule_path, "w", encoding="utf-8-sig") as f:
            json.dump(existing_rules, f, indent=2, ensure_ascii=False)
    except Exception as e:
        raise IOError(f"Failed to write to rules file {rule_path}: {e}")

    return added_count, updated_count
