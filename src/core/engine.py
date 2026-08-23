from core.duplicate_filter import (
    ExistingTransactionMatcher,
)
from core.bank_3332_rules import apply_bank_specific_rule_override
from core.client_config import get_client_bank_categories


class VoucherEngine:
    def __init__(
        self,
        rule_engine,
        builder_registry,
        duplicate_json_path=None,
        bank_ledger=None,
        client_name=None
    ):
        self.rule_engine = rule_engine
        self.builder_registry = builder_registry
        self.bank_ledger = bank_ledger
        self.unclassified = []
        self.duplicates = []
        self.duplicate_matcher = None

        if duplicate_json_path:
            self.duplicate_matcher = ExistingTransactionMatcher(duplicate_json_path)

        self.client_banks = set()
        if client_name:
            try:
                categories = get_client_bank_categories(client_name)
                for banks in categories.values():
                    for b in banks:
                        self.client_banks.add(str(b).strip().lower())
            except Exception as e:
                print(f"Error loading client bank accounts in VoucherEngine: {e}")

    def _resolve_voucher_type(self, rule, transaction):
        """
        Voucher type source of truth:
        - Contra can be explicitly forced by rule.
        - If mapped ledger belongs to the client's other bank accounts, treat as Contra.
        - Otherwise decide from statement direction.
        """
        if bool(rule.get("is_contra")) or bool(rule.get("contra")):
            return "Contra"

        configured_type = str(rule.get("voucher_type", "")).strip().lower()
        if configured_type == "contra":
            return "Contra"

        # Check if target ledger is one of the client's own banks
        resolved_ledger = rule.get("ledger")
        if resolved_ledger:
            resolved_ledger_clean = str(resolved_ledger).strip().lower()
            if resolved_ledger_clean in self.client_banks:
                return "Contra"

        direction = getattr(transaction, "direction", None)
        if direction == "OUT":
            return "Payment"
        if direction == "IN":
            return "Receipt"
        return None

    def _resolve_directional_rule(self, rule, transaction):
        """
        Optional extension:
        - payment_ledger / out_ledger => OUT transactions
        - receipt_ledger / in_ledger => IN transactions
        - fall back to any available ledger key if direction-specific is missing.
        """
        resolved_rule = dict(rule)
        direction = getattr(transaction, "direction", None)

        if direction == "OUT":
            resolved_rule["ledger"] = (
                rule.get("payment_ledger")
                or rule.get("out_ledger")
                or rule.get("ledger")
                or rule.get("receipt_ledger")
                or rule.get("in_ledger")
            )
        elif direction == "IN":
            resolved_rule["ledger"] = (
                rule.get("receipt_ledger")
                or rule.get("in_ledger")
                or rule.get("ledger")
                or rule.get("payment_ledger")
                or rule.get("out_ledger")
            )
        else:
            resolved_rule["ledger"] = (
                rule.get("ledger")
                or rule.get("payment_ledger")
                or rule.get("receipt_ledger")
                or rule.get("out_ledger")
                or rule.get("in_ledger")
            )

        # Evaluate amount_tiers if defined
        if "amount_tiers" in rule and isinstance(rule["amount_tiers"], list):
            txn_amount = getattr(transaction, "amount", 0.0)
            for tier in rule["amount_tiers"]:
                t_min = float(tier.get("min", float("-inf")))
                t_max = float(tier.get("max", float("inf")))
                if t_min <= txn_amount <= t_max:
                    if direction == "OUT":
                        resolved_rule["ledger"] = (
                            tier.get("payment_ledger")
                            or tier.get("out_ledger")
                            or tier.get("ledger")
                            or tier.get("receipt_ledger")
                            or tier.get("in_ledger")
                        )
                    elif direction == "IN":
                        resolved_rule["ledger"] = (
                            tier.get("receipt_ledger")
                            or tier.get("in_ledger")
                            or tier.get("ledger")
                            or tier.get("payment_ledger")
                            or tier.get("out_ledger")
                        )
                    else:
                        resolved_rule["ledger"] = (
                            tier.get("ledger")
                            or tier.get("payment_ledger")
                            or tier.get("receipt_ledger")
                            or tier.get("out_ledger")
                            or tier.get("in_ledger")
                        )
                    
                    if "voucher_type" in tier:
                        resolved_rule["voucher_type"] = tier["voucher_type"]
                        if tier["voucher_type"].lower() == "contra":
                            resolved_rule["is_contra"] = True
                    break

        if "voucher_type" not in resolved_rule or not resolved_rule["voucher_type"]:
            resolved_rule["voucher_type"] = self._resolve_voucher_type(resolved_rule, transaction)

        return resolved_rule

    def process(self, transactions):
        vouchers = []

        for txn in transactions:
            rule = self.rule_engine.match(txn)

            if not rule:
                txn_amount = getattr(txn, "amount", 0.0)
                if txn_amount >= 500000:
                    setattr(txn, "unclassified_reason", "amount greater than 500000 but no ledger found")
                    self.unclassified.append(txn)
                    continue
                else:
                    direction = getattr(txn, "direction", None)
                    rule = {}
                    if direction == "OUT":
                        rule["ledger"] = "Labour Charges"
                    elif direction == "IN":
                        rule["ledger"] = "Receipts"
                    else:
                        rule["ledger"] = "Labour Charges"

            resolved_rule = self._resolve_directional_rule(rule, txn)
            resolved_rule = apply_bank_specific_rule_override(
                resolved_rule,
                txn,
                self.bank_ledger
            )

            # Safeguard: if resolved ledger is empty, mark as unclassified
            if not resolved_rule.get("ledger"):
                self.unclassified.append(txn)
                continue

            builder = self.builder_registry.get(resolved_rule.get("voucher_type"))

            if not builder:
                self.unclassified.append(txn)
                continue

            voucher = builder.build(txn, resolved_rule)

            if voucher:
                if self.duplicate_matcher:
                    matched_record = self.duplicate_matcher.find_voucher_match(
                        voucher.get("Date"),
                        voucher.get("Amount"),
                        [voucher.get("Cr_Ledger"), voucher.get("Dr_Ledger")],
                    )

                    if matched_record:
                        duplicate_row = dict(voucher)
                        duplicate_row.update(matched_record["json_details"])
                        self.duplicates.append(duplicate_row)
                        continue

                vouchers.append(voucher)
            else:
                self.unclassified.append(txn)

        return vouchers
