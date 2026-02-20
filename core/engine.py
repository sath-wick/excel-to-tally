class VoucherEngine:
    def __init__(self, rule_engine, builder_registry):
        self.rule_engine = rule_engine
        self.builder_registry = builder_registry
        self.unclassified = []

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
                vouchers.append(voucher)
            else:
                self.unclassified.append(txn)

        return vouchers
