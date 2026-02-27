from core.duplicate_filter import (
    load_existing_contras,
    normalize_amount,
    normalize_date,
    normalize_ledger,
)


class VoucherEngine:
    def __init__(self, rule_engine, builder_registry, duplicate_json_path=None):
        self.rule_engine = rule_engine
        self.builder_registry = builder_registry
        self.unclassified = []
        self.duplicates = []
        self.existing_contras = set()

        if duplicate_json_path:
            self.existing_contras = load_existing_contras(duplicate_json_path)

    def process(self, transactions):
        vouchers = []

        for txn in transactions:
            rule = self.rule_engine.match(txn)

            if not rule:
                self.unclassified.append(txn)
                continue

            builder = self.builder_registry.get(rule["voucher_type"])

            if not builder:
                self.unclassified.append(txn)
                continue

            voucher = builder.build(txn, rule)

            if voucher:
                if voucher.get("Voucher_Type") == "Contra" and self.existing_contras:
                    duplicate_key = (
                        normalize_date(voucher.get("Date")),
                        normalize_ledger(voucher.get("Cr_Ledger")),
                        normalize_amount(voucher.get("Amount")),
                    )

                    if (
                        duplicate_key[0] is not None
                        and duplicate_key[1]
                        and duplicate_key[2] is not None
                        and duplicate_key in self.existing_contras
                    ):
                        self.duplicates.append(voucher)
                        continue

                vouchers.append(voucher)
            else:
                self.unclassified.append(txn)

        return vouchers
