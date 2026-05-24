import os
import sys
import pandas as pd

from extract.pdf_extractor import extract_bank_statement
from core.duplicate_filter import JSON_DUPLICATE_COLUMNS, split_statement_duplicates
from core.ignore_filter import split_ignored_descriptions
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

DUPLICATE_JSON_PATH = "./exports/Transactions.json"
IGNORE_JSON_PATH = "./rules/ignored_descriptions.json"

FINAL_OUTPUT = "./output/Statement_Import.xlsx"
RULE_PATH = "./rules/description_rules_494.json"


def resolve_rule_path(bank_ledger):
    ledger_text = str(bank_ledger).strip()
    digits_only = "".join(ch for ch in ledger_text if ch.isdigit())
    slug = "_".join(ledger_text.lower().split())

    if digits_only == "494":
        return RULE_PATH

    candidates = []
    if ledger_text:
        candidates.append(ledger_text)
    if slug and slug != ledger_text:
        candidates.append(slug)
    if digits_only and digits_only not in candidates:
        candidates.append(digits_only)

    for candidate in candidates:
        bank_specific_rule_path = f"./rules/description_rules_{candidate}.json"
        if os.path.exists(bank_specific_rule_path):
            return bank_specific_rule_path

    return RULE_PATH


def confirm_step(message):
    print("\n" + "=" * 50)
    print(message)
    print("=" * 50)
    choice = input("Continue? (Y/N): ").strip().lower()
    return choice == "y"


def prepare_voucher_sheet(df_output, voucher_type):
    voucher_columns = [
        "Voucher_Type",
        "Date",
        "Description",
        "Narration",
        "Cr_Ledger",
        "Amount",
        "Cr",
        "Dr_Ledger",
        "Dr_Amount",
        "Dr",
    ]

    typed_df = df_output[df_output["Voucher_Type"] == voucher_type].copy()

    if typed_df.empty:
        return pd.DataFrame(columns=["Voucher_Num"] + voucher_columns)

    typed_df = typed_df.reindex(columns=voucher_columns)
    typed_df.reset_index(drop=True, inplace=True)
    typed_df.insert(0, "Voucher_Num", typed_df.index + 1)
    return typed_df


def build_unclassified_df(unclassified_transactions):
    columns = ["Value Date", "Description", "Withdrawal", "Deposit", "Reference"]

    if not unclassified_transactions:
        return pd.DataFrame(columns=columns)

    rows = [
        {
            "Value Date": txn.date,
            "Description": txn.description,
            "Withdrawal": txn.withdrawal,
            "Deposit": txn.deposit,
            "Reference": txn.reference,
        }
        for txn in unclassified_transactions
    ]

    return pd.DataFrame(rows, columns=columns)


def with_json_duplicate_columns(df):
    output = df.copy()
    for column in JSON_DUPLICATE_COLUMNS:
        if column not in output.columns:
            output[column] = ""
    return output


def main():
    statement_df = extract_bank_statement(PDF_PATH, bank_ledger=BANK_LEDGER)
    print("\nBank statement extracted successfully.")

    if not confirm_step("Proceed with duplicate filtering and voucher generation?"):
        print("\nProcess stopped by user.")
        return

    filtered_statement_df = statement_df.copy()
    statement_duplicates_df = statement_df.iloc[0:0].copy()
    for column in JSON_DUPLICATE_COLUMNS:
        statement_duplicates_df[column] = ""

    ignored_descriptions_df = statement_df.iloc[0:0].copy()

    if os.path.exists(DUPLICATE_JSON_PATH):
        filtered_statement_df, statement_duplicates_df = split_statement_duplicates(
            filtered_statement_df,
            DUPLICATE_JSON_PATH,
        )
    else:
        print(f"\nDuplicate JSON not found: {DUPLICATE_JSON_PATH}")

    if os.path.exists(IGNORE_JSON_PATH):
        filtered_statement_df, ignored_descriptions_df = split_ignored_descriptions(
            filtered_statement_df,
            IGNORE_JSON_PATH,
        )
    else:
        print(f"\nIgnore rules JSON not found: {IGNORE_JSON_PATH}")

    transactions = [Transaction(row) for _, row in filtered_statement_df.iterrows()]

    active_rule_path = resolve_rule_path(BANK_LEDGER)
    print(f"Using rules file: {active_rule_path}")
    rule_engine = RuleEngine(active_rule_path)

    builder_registry = {
        "Contra": ContraBuilder(BANK_LEDGER),
        "Payment": PaymentBuilder(BANK_LEDGER),
        "Receipt": ReceiptBuilder(BANK_LEDGER),
    }

    duplicate_json_path = DUPLICATE_JSON_PATH if os.path.exists(DUPLICATE_JSON_PATH) else None
    engine = VoucherEngine(
        rule_engine,
        builder_registry,
        duplicate_json_path=duplicate_json_path,
        bank_ledger=BANK_LEDGER,
    )

    vouchers = engine.process(transactions)
    df_output = pd.DataFrame(vouchers)

    if df_output.empty:
        df_output = pd.DataFrame(
            columns=[
                "Voucher_Type",
                "Date",
                "Description",
                "Narration",
                "Cr_Ledger",
                "Amount",
                "Cr",
                "Dr_Ledger",
                "Dr_Amount",
                "Dr",
            ]
        )

    payment_import_df = prepare_voucher_sheet(df_output, "Payment")
    receipt_import_df = prepare_voucher_sheet(df_output, "Receipt")
    contra_import_df = prepare_voucher_sheet(df_output, "Contra")

    statement_duplicate_count = len(statement_duplicates_df)
    voucher_duplicate_count = len(engine.duplicates)
    ignored_description_count = len(ignored_descriptions_df)
    total_duplicate_count = statement_duplicate_count + voucher_duplicate_count

    print(f"\nDuplicate transactions identified and ignored: {total_duplicate_count}")
    print(f"Ignored by description rules: {ignored_description_count}")
    print(f"Unclassified transactions: {len(engine.unclassified)}")

    duplicate_frames = []

    if statement_duplicate_count > 0:
        statement_duplicate_sheet = with_json_duplicate_columns(statement_duplicates_df)
        statement_duplicate_sheet.insert(0, "Source", "Statement")
        duplicate_frames.append(statement_duplicate_sheet)

    if voucher_duplicate_count > 0:
        voucher_duplicate_sheet = pd.DataFrame(engine.duplicates)
        voucher_duplicate_sheet = with_json_duplicate_columns(voucher_duplicate_sheet)
        voucher_duplicate_sheet.insert(0, "Source", "Voucher")
        duplicate_frames.append(voucher_duplicate_sheet)

    if duplicate_frames:
        duplicates_df = pd.concat(duplicate_frames, ignore_index=True, sort=False)
    else:
        duplicates_df = pd.DataFrame(columns=["Source"] + JSON_DUPLICATE_COLUMNS)

    duplicates_df.reset_index(drop=True, inplace=True)
    duplicates_df.insert(0, "Duplicate_Num", duplicates_df.index + 1)

    ignored_sheet_df = ignored_descriptions_df.copy()
    ignored_sheet_df.reset_index(drop=True, inplace=True)
    ignored_sheet_df.insert(0, "Ignored_Num", ignored_sheet_df.index + 1)

    unclassified_df = build_unclassified_df(engine.unclassified)
    unclassified_df.reset_index(drop=True, inplace=True)
    unclassified_df.insert(0, "Unclassified_Num", unclassified_df.index + 1)

    def write_statement_workbook():
        with pd.ExcelWriter(FINAL_OUTPUT, engine="openpyxl") as writer:
            statement_df.to_excel(writer, sheet_name="Bank statement", index=False)
            payment_import_df.to_excel(writer, sheet_name="Import Payment", index=False)
            receipt_import_df.to_excel(writer, sheet_name="Import Receipt", index=False)
            contra_import_df.to_excel(writer, sheet_name="Import Contra", index=False)
            ignored_sheet_df.to_excel(writer, sheet_name="Ignored", index=False)
            duplicates_df.to_excel(writer, sheet_name="Duplicates", index=False)
            unclassified_df.to_excel(writer, sheet_name="Unclassified", index=False)

    safe_excel_write(write_statement_workbook, FINAL_OUTPUT)

    print(f"\nStatement import workbook generated: {FINAL_OUTPUT}")


if __name__ == "__main__":
    main()
