class BaseBuilder:
    def __init__(self, bank_ledger):
        self.bank_ledger = bank_ledger

    def _format(self, transaction, dr, cr, voucher_type):
        return {
            "Voucher_Type": voucher_type,
            "Date": transaction.date,
            "Description": transaction.description,
            "Narration": "",  # reverted
            "Cr_Ledger": cr,
            "Amount": transaction.amount,
            "Cr": "CR",
            "Dr_Ledger": dr,
            "Dr_Amount": transaction.amount,
            "Dr": "DR"
        }


class ContraBuilder(BaseBuilder):
    def build(self, transaction, rule):
        counter_ledger = rule["ledger"]

        if transaction.direction == "IN":
            dr_ledger = self.bank_ledger
            cr_ledger = counter_ledger
        elif transaction.direction == "OUT":
            dr_ledger = counter_ledger
            cr_ledger = self.bank_ledger
        else:
            return None

        return self._format(transaction, dr_ledger, cr_ledger, "Contra")


class PaymentBuilder(BaseBuilder):
    def build(self, transaction, rule):
        counter_ledger = rule["ledger"]

        if transaction.direction == "OUT":
            dr_ledger = counter_ledger
            cr_ledger = self.bank_ledger
        else:
            return None

        return self._format(transaction, dr_ledger, cr_ledger, "Payment")


class ReceiptBuilder(BaseBuilder):
    def build(self, transaction, rule):
        counter_ledger = rule["ledger"]

        if transaction.direction == "IN":
            dr_ledger = self.bank_ledger
            cr_ledger = counter_ledger
        else:
            return None

        return self._format(transaction, dr_ledger, cr_ledger, "Receipt")
