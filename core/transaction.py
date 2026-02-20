import pandas as pd


class Transaction:
    def __init__(self, row):
        # --- Date handling (FINAL FIX) ---
        raw_date = pd.to_datetime(
            row.get("Value Date"),
            errors="coerce",
            dayfirst=False  # bank statement is yyyy-mm-dd
        )

        if pd.notnull(raw_date):
            self.date = raw_date.strftime("%d-%m-%Y")
        else:
            self.date = ""

        # --- Description ---
        self.description = str(row.get("Description", "")).strip()

        # --- Reference ---
        self.reference = str(row.get("Reference Number", "")).strip()

        # --- Amount Handling ---
        withdrawal = row.get("Withdrawals")
        deposit = row.get("Deposits")

        self.withdrawal = self._parse_amount(withdrawal)
        self.deposit = self._parse_amount(deposit)

        if self.deposit > 0:
            self.amount = self.deposit
            self.direction = "IN"
        elif self.withdrawal > 0:
            self.amount = self.withdrawal
            self.direction = "OUT"
        else:
            self.amount = 0
            self.direction = None

    def _parse_amount(self, value):
        if pd.isna(value) or value == "":
            return 0

        value = str(value).replace(",", "").strip()
        try:
            return float(value)
        except ValueError:
            return 0
