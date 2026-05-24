from core.duplicate_filter import (
    ExistingTransactionMatcher,
)
from core.bank_3332_rules import apply_bank_specific_rule_override


class VoucherEngine:
    def __init__(
        self,
        rule_engine,
        builder_registry,
        duplicate_json_path=None,
        bank_ledger=None
    ):
        self.rule_engine = rule_engine
        self.builder_registry = builder_registry
        self.bank_ledger = bank_ledger
        self.unclassified = []
        self.duplicates = []
        self.duplicate_matcher = None

        if duplicate_json_path:
            self.duplicate_matcher = ExistingTransactionMatcher(duplicate_json_path)

    def _resolve_directional_rule(self, rule, transaction):
        """
        Optional extension:
        - payment_ledger / out_ledger => OUT transactions as Payment
        - receipt_ledger / in_ledger => IN transactions as Receipt
        Falls back to original rule behavior when these keys are absent.
        """
        direction = getattr(transaction, "direction", None)
        has_directional_ledger = any(
            key in rule
            for key in ("payment_ledger", "out_ledger", "receipt_ledger", "in_ledger")
        )

        if not has_directional_ledger:
            return rule

        resolved_rule = dict(rule)

        if direction == "OUT":
            resolved_rule["voucher_type"] = "Payment"
            resolved_rule["ledger"] = (
                rule.get("payment_ledger")
                or rule.get("out_ledger")
                or rule.get("ledger")
            )
            return resolved_rule

        if direction == "IN":
            resolved_rule["voucher_type"] = "Receipt"
            resolved_rule["ledger"] = (
                rule.get("receipt_ledger")
                or rule.get("in_ledger")
                or rule.get("ledger")
            )
            return resolved_rule

        return rule

    def process(self, transactions):
        vouchers = []

        for txn in transactions:
            rule = self.rule_engine.match(txn)

            if not rule:
                self.unclassified.append(txn)
                continue

            resolved_rule = self._resolve_directional_rule(rule, txn)
            resolved_rule = apply_bank_specific_rule_override(
                resolved_rule,
                txn,
                self.bank_ledger
            )
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
