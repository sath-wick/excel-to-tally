import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog


BANK_CATEGORIES = {
    "YES": [
        "YES 1060",
        "YES 2085",
        "YES BANK"
    ],
    "SBI": [
        "SBI 4452",
        "SBI 9382"
    ],
    "Kotak": [
        "Kotak 2449",
        "KOTAK 2755",
        "Kotak 44000",
        "KOTAK 555",
        "KOTAK 8113",
        "KOTAK 9699",
        "Kotk 11053"
    ],
    "DBS": [
        "DBS 232",
        "DBS361"
    ],
    "Others": [
        "Bank 3332",
        "Cash"
    ]
}


def clear():
    os.system("cls")


def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Bank Statement PDF",
        filetypes=[("PDF Files", "*.pdf")]
    )
    return file_path


def select_from_list(options, title):
    clear()
    print("============================================")
    print(f"        {title}")
    print("============================================\n")

    for idx, item in enumerate(options, start=1):
        print(f"{idx}. {item}")

    print()
    choice = input("Enter option number: ").strip()

    if not choice.isdigit():
        return None

    choice = int(choice)

    if 1 <= choice <= len(options):
        return options[choice - 1]

    return None


def main():
    clear()
    print("============================================")
    print("         TALLY IMPORT GENERATOR")
    print("============================================\n")

    print("Select Import Type:")
    print("1. Sales")
    print("2. Purchases")
    print("3. Statements\n")

    choice = input("Enter option number: ").strip()

    if choice != "3":
        print("\nOnly 'Statements' is implemented right now.")
        input("Press Enter to exit...")
        sys.exit()

    # Step 1: Select Bank Category
    category = select_from_list(list(BANK_CATEGORIES.keys()), "SELECT BANK CATEGORY")

    if not category:
        print("\nInvalid selection. Exiting.")
        sys.exit()

    # Step 2: Select Bank Ledger
    bank_ledger = select_from_list(BANK_CATEGORIES[category], "SELECT BANK")

    if not bank_ledger:
        print("\nInvalid selection. Exiting.")
        sys.exit()

    # Step 3: Select PDF File
    print("\nSelect the bank statement PDF...")
    file_path = select_file()

    if not file_path:
        print("\nNo file selected. Exiting.")
        sys.exit()

    print(f"\nSelected Bank: {bank_ledger}")
    print(f"Selected File: {file_path}")
    print("\nProcessing...\n")

    # Pass file path and bank ledger to main.py
    subprocess.run(["python", "main.py", file_path, bank_ledger])


if __name__ == "__main__":
    main()
