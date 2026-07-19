import json
import re


def load_ignore_rules(json_path):
    with open(json_path, "r", encoding="utf-8") as file_obj:
        payload = json.load(file_obj)

    rules = payload.get("description_ignore_rules", [])
    normalized_rules = []

    for index, rule in enumerate(rules, start=1):
        if isinstance(rule, str):
            normalized_rules.append(
                {
                    "name": f"rule_{index}",
                    "pattern": rule,
                    "mode": "contains",
                    "case_sensitive": False,
                }
            )
            continue

        if not isinstance(rule, dict):
            continue

        pattern = str(rule.get("pattern", "")).strip()
        if not pattern:
            continue

        mode = str(rule.get("mode", "contains")).strip().lower()
        if mode not in {"contains", "exact", "regex"}:
            mode = "contains"

        normalized_rules.append(
            {
                "name": str(rule.get("name", f"rule_{index}")).strip() or f"rule_{index}",
                "pattern": pattern,
                "mode": mode,
                "case_sensitive": bool(rule.get("case_sensitive", False)),
            }
        )

    return normalized_rules


def split_ignored_descriptions(statement_df, rule_json_path):
    ignore_rules = load_ignore_rules(rule_json_path)
    if not ignore_rules:
        return statement_df.copy(), statement_df.iloc[0:0].copy()

    ignored_indices = []
    ignored_rule_names = []

    for index, row in statement_df.iterrows():
        description = str(row.get("Description", "")).strip()
        matched_rule_name = _match_rule_name(description, ignore_rules)
        if not matched_rule_name:
            continue

        ignored_indices.append(index)
        ignored_rule_names.append(matched_rule_name)

    ignored_df = statement_df.loc[ignored_indices].copy()
    filtered_df = statement_df.drop(index=ignored_indices).copy()

    if not ignored_df.empty:
        ignored_df.insert(0, "Ignore_Rule", ignored_rule_names)

    ignored_df.reset_index(drop=True, inplace=True)
    filtered_df.reset_index(drop=True, inplace=True)
    return filtered_df, ignored_df


def _match_rule_name(description, rules):
    for rule in rules:
        pattern = rule["pattern"]
        mode = rule["mode"]
        case_sensitive = rule["case_sensitive"]
        source_text = description if case_sensitive else description.lower()
        target_text = pattern if case_sensitive else pattern.lower()

        if mode == "contains" and target_text in source_text:
            return rule["name"]

        if mode == "exact" and source_text == target_text:
            return rule["name"]

        if mode == "regex":
            flags = 0 if case_sensitive else re.IGNORECASE
            if re.search(pattern, description, flags):
                return rule["name"]

    return None
