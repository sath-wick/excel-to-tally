import json
import re


class RuleEngine:
    def __init__(self, rule_path):
        with open(rule_path, "r") as f:
            self.rules = json.load(f)

        # Highest priority first
        self.rules.sort(key=lambda r: r.get("priority", 0), reverse=True)

    def match(self, transaction):
        description = transaction.description.strip()

        for rule in self.rules:
            patterns = rule["pattern"]

            if isinstance(patterns, str):
                patterns = [patterns]

            for pattern in patterns:
                if re.search(pattern, description, re.IGNORECASE):
                    return rule  # ← only rule

        return None
