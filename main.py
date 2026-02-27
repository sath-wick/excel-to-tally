import sys
import pandas as pd

from extract.pdf_extractor import extract_bank_statement
from core.transaction import Transaction
from core.rule_engine import RuleEngine
from core.builders import ContraBuilder, PaymentBuilder, ReceiptBuilder
from core.engine import VoucherEngine
from utils.file_writer import safe_excel_write


# -------- Runtime Arguments --------
if len(sys.argv) > 1:
    PDF_PATH = sys.argv[1]
else:
    PDF_PATH = "./input/Statements/Nov_Statement.pdf"

if len(sys.argv) > 2:
    BANK_LEDGER = sys.argv[2]
else:
    BANK_LEDGER = "494"

if len(sys.argv) > 3:
    DUPLICATE_JSON_PATH = sys.argv[3]
else:
    DUPLICATE_JSON_PATH = None

EXCEL_PATH = "./output/bank_statement.xlsx"
FINAL_OUTPUT = "./output/statement_import_ready.xlsx"
UNCLASSIFIED_OUTPUT = "./output/unclassified.xlsx"
DUPLICATE_OUTPUT = "./output/duplicate_entries.xlsx"
RULE_PATH = "./rules/description_rules.json"


def confirm_step(message):
    print("\n" + "=" * 50)
    print(message)
    print("=" * 50)
    choice = input("Continue? (Y/N): ").strip().lower()
    return choice == "y"


def main():

    # 1️⃣ Extract Bank Statement
    safe_excel_write(
        lambda: extract_bank_statement(PDF_PATH, EXCEL_PATH),
        EXCEL_PATH
    )

    print("\nBank statement generated successfully.")
    
    # 2️⃣ Human Confirmation Step
    if not confirm_step("Please review the generated bank_statement.xlsx file before proceeding."):
        print("\nProcess stopped by user after bank statement generation.")
        return

    # 3️⃣ Load statement
    df = pd.read_excel(EXCEL_PATH)
    transactions = [Transaction(row) for _, row in df.iterrows()]

    # 4️⃣ Setup rule engine
    rule_engine = RuleEngine(RULE_PATH)

    builder_registry = {
        "Contra": ContraBuilder(BANK_LEDGER),
        "Payment": PaymentBuilder(BANK_LEDGER),
        "Receipt": ReceiptBuilder(BANK_LEDGER)
    }

    engine = VoucherEngine(
        rule_engine,
        builder_registry,
        duplicate_json_path=DUPLICATE_JSON_PATH
    )

    # 5️⃣ Process transactions
    vouchers = engine.process(transactions)

    df_output = pd.DataFrame(vouchers)

    if not df_output.empty:

        payments_df = df_output[df_output["Voucher_Type"] == "Payment"].copy()
        receipts_df = df_output[df_output["Voucher_Type"] == "Receipt"].copy()
        contra_df = df_output[df_output["Voucher_Type"] == "Contra"].copy()

        def write_voucher_file():
            with pd.ExcelWriter(FINAL_OUTPUT, engine="openpyxl") as writer:

                if not payments_df.empty:
                    payments_df.reset_index(drop=True, inplace=True)
                    payments_df.insert(0, "Voucher_Num", payments_df.index + 1)
                    payments_df.to_excel(writer, sheet_name="Payment", index=False)

                if not receipts_df.empty:
                    receipts_df.reset_index(drop=True, inplace=True)
                    receipts_df.insert(0, "Voucher_Num", receipts_df.index + 1)
                    receipts_df.to_excel(writer, sheet_name="Receipt", index=False)

                if not contra_df.empty:
                    contra_df.reset_index(drop=True, inplace=True)
                    contra_df.insert(0, "Voucher_Num", contra_df.index + 1)
                    contra_df.to_excel(writer, sheet_name="Contra", index=False)

        safe_excel_write(write_voucher_file, FINAL_OUTPUT)

        print("\nVoucher file generated successfully.")

    else:
        print("\nNo vouchers generated.")

    # 6️⃣ Export Unclassified
    if engine.duplicates:
        duplicate_df = pd.DataFrame(engine.duplicates).copy()
        duplicate_df.reset_index(drop=True, inplace=True)
        duplicate_df.insert(0, "Voucher_Num", duplicate_df.index + 1)

        safe_excel_write(
            lambda: duplicate_df.to_excel(DUPLICATE_OUTPUT, index=False),
            DUPLICATE_OUTPUT
        )

        print(f"\nDuplicate contra vouchers skipped: {len(engine.duplicates)}")

    if engine.unclassified:

        unclassified_data = [
            {
                "Value Date": txn.date,
                "Description": txn.description,
                "Withdrawal": txn.withdrawal,
                "Deposit": txn.deposit,
                "Reference": txn.reference
            }
            for txn in engine.unclassified
        ]

        safe_excel_write(
            lambda: pd.DataFrame(unclassified_data).to_excel(
                UNCLASSIFIED_OUTPUT, index=False
            ),
            UNCLASSIFIED_OUTPUT
        )

        print(f"\nUnclassified transactions: {len(engine.unclassified)}")


if __name__ == "__main__":
    main()
