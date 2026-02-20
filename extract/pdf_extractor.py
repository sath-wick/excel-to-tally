import camelot
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


STANDARD_COLUMNS = [
    "Transaction Date",
    "Value Date",
    "Description",
    "Reference Number",
    "Withdrawals",
    "Deposits",
    "Running Balance"
]


# -----------------------------------------------------
# Header Utilities
# -----------------------------------------------------

def normalize_header(text):
    return (
        str(text)
        .lower()
        .replace("\n", " ")
        .replace("/", " ")
        .strip()
    )


def map_column(col):
    col_lower = normalize_header(col)

    if "transaction" in col_lower:
        return "Transaction Date"

    if "value" in col_lower:
        return "Value Date"

    if "description" in col_lower:
        return "Description"

    if "reference" in col_lower or "cheque" in col_lower or "ref" in col_lower:
        return "Reference Number"

    if "withdraw" in col_lower:
        return "Withdrawals"

    if "deposit" in col_lower:
        return "Deposits"

    if "balance" in col_lower:
        return "Running Balance"

    return None


# -----------------------------------------------------
# Repair Merged Deposit / Balance Columns
# -----------------------------------------------------

def repair_merged_amounts(df):

    pattern = re.compile(r"^\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$")

    for i in range(len(df)):

        deposit = str(df.loc[i, "Deposits"]).strip()
        balance = str(df.loc[i, "Running Balance"]).strip()

        # Case 1: Deposits contains both values
        match = pattern.match(deposit)
        if match and balance == "":
            df.loc[i, "Deposits"] = match.group(1)
            df.loc[i, "Running Balance"] = match.group(2)
            continue

        # Case 2: Running Balance contains both values
        match = pattern.match(balance)
        if match and deposit == "":
            df.loc[i, "Deposits"] = match.group(1)
            df.loc[i, "Running Balance"] = match.group(2)

    return df


# -----------------------------------------------------
# Merge Spillover Rows (Multiline Transactions)
# -----------------------------------------------------

def merge_spillover_rows(df):

    rows_to_drop = []

    for i in range(1, len(df)):

        txn_date = str(df.loc[i, "Transaction Date"]).strip()
        val_date = str(df.loc[i, "Value Date"]).strip()
        withdrawal = str(df.loc[i, "Withdrawals"]).strip()
        deposit = str(df.loc[i, "Deposits"]).strip()
        desc = str(df.loc[i, "Description"]).strip()

        # Spillover condition:
        if (
            desc != "" and
            txn_date == "" and
            val_date == "" and
            withdrawal == "" and
            deposit == ""
        ):
            prev_desc = str(df.loc[i - 1, "Description"]).strip()
            df.loc[i - 1, "Description"] = f"{prev_desc} {desc}"
            rows_to_drop.append(i)

    df = df.drop(index=rows_to_drop).reset_index(drop=True)

    return df


# -----------------------------------------------------
# Main Extraction Function
# -----------------------------------------------------

def extract_bank_statement(pdf_path, output_file):

    tables = camelot.read_pdf(pdf_path, pages="all")

    print(f"Total tables detected: {tables.n}")

    dataframes = []
    accepted_tables = 0
    ignored_tables = 0

    for table in tables:

        df = table.df

        if df.empty or df.shape[0] < 2:
            ignored_tables += 1
            continue

        # First row as header
        df.columns = df.iloc[0]
        df = df[1:]

        # Dynamic column mapping
        mapped_columns = {}
        for col in df.columns:
            mapped = map_column(col)
            if mapped:
                mapped_columns[col] = mapped

        df = df.rename(columns=mapped_columns)

        # Required fields
        required = {
            "Transaction Date",
            "Value Date",
            "Description",
            "Withdrawals",
            "Deposits"
        }

        if not required.issubset(set(df.columns)):
            ignored_tables += 1
            continue

        # Add optional columns if missing
        if "Reference Number" not in df.columns:
            df["Reference Number"] = ""

        if "Running Balance" not in df.columns:
            df["Running Balance"] = ""

        # Standardize column order
        df = df[STANDARD_COLUMNS]

        accepted_tables += 1
        dataframes.append(df)

    if not dataframes:
        raise ValueError("No valid tables found in PDF. Extraction aborted.")

    final_df = pd.concat(dataframes, ignore_index=True)

    # Clean all cell values
    for col in final_df.columns:
        final_df[col] = (
            final_df[col]
            .astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace("\r", "", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    # Repair merged columns
    final_df = repair_merged_amounts(final_df)

    # Merge multiline spillovers
    final_df = merge_spillover_rows(final_df)

    # Save Excel
    final_df.to_excel(output_file, index=False)

    # Format as Excel table
    wb = load_workbook(output_file)
    ws = wb.active

    last_row = ws.max_row
    last_col = ws.max_column
    table_range = f"A1:{ws.cell(row=last_row, column=last_col).coordinate}"

    excel_table = Table(displayName="CombinedTable", ref=table_range)

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    excel_table.tableStyleInfo = style
    ws.add_table(excel_table)

    wb.save(output_file)

    print(f"Tables accepted: {accepted_tables}")
    print(f"Tables ignored : {ignored_tables}")
    print("Bank statement extraction completed.")

    return {
        "accepted": accepted_tables,
        "ignored": ignored_tables
    }
