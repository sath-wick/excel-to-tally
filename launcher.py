import os
import sys
import subprocess
import msvcrt
import tkinter as tk
from tkinter import filedialog


BANK_CATEGORIES = {
    "YES": [
        "YES 1060",
        "YES 2085",
        "494"
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

BACK = "__BACK__"


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


def select_sales_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Sales Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    return file_path


def select_purchases_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Purchases Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    return file_path


def read_menu_choice(max_option, allow_back=False):
    while True:
        key = msvcrt.getch()

        # Ignore special keys (arrows, function keys, etc.)
        if key in (b"\x00", b"\xe0"):
            msvcrt.getch()
            continue

        if allow_back and key == b"\x1b":
            print("Esc")
            return BACK

        if key.isdigit():
            choice = int(key.decode())
            if 1 <= choice <= max_option:
                print(choice)
                return choice


def select_from_list(options, title, allow_back=False):
    clear()
    print("============================================")
    print(f"        {title}")
    print("============================================\n")

    for idx, item in enumerate(options, start=1):
        print(f"{idx}. {item}")

    print("\nPress number key to select.")
    if allow_back:
        print("Press Esc to go back.")

    choice = read_menu_choice(len(options), allow_back=allow_back)
    if choice == BACK:
        return BACK

    if 1 <= choice <= len(options):
        return options[choice - 1]

    return None


def extract_bank_code(bank_ledger):
    ledger_text = str(bank_ledger).strip()
    digits_only = "".join(ch for ch in ledger_text if ch.isdigit())
    return digits_only if digits_only else ledger_text


def run_excel_import_flow(module_name, script_name, file_selector):
    while True:
        category = select_from_list(
            list(BANK_CATEGORIES.keys()),
            "SELECT BANK CATEGORY",
            allow_back=True
        )

        if category == BACK:
            return False

        while True:
            bank_ledger = select_from_list(
                BANK_CATEGORIES[category],
                "SELECT BANK",
                allow_back=True
            )

            if bank_ledger == BACK:
                break

            if extract_bank_code(bank_ledger) != "494":
                print(f"\n{module_name} is currently available only for bank 494.")
                print("Press any key to choose another bank...")
                msvcrt.getch()
                continue

            print(f"\nSelect the {module_name.lower()} input Excel file...")
            file_path = file_selector()

            if not file_path:
                print("\nNo file selected. Exiting.")
                sys.exit()

            print(f"\nSelected Bank: {bank_ledger}")
            print(f"Selected {module_name} File: {file_path}")
            print("\nProcessing...\n")
            subprocess.run(["python", script_name, file_path, bank_ledger])
            return True


def main():
    while True:
        clear()
        print("============================================")
        print("         TALLY IMPORT GENERATOR")
        print("============================================\n")

        print("Select Import Type:")
        print("1. Sales")
        print("2. Purchases")
        print("3. Statements\n")
        print("Press number key to select.")

        choice = read_menu_choice(3)

        if choice == 1:
            if run_excel_import_flow("Sales", "sales_main.py", select_sales_file):
                return
            continue

        if choice == 2:
            if run_excel_import_flow("Purchases", "purchases_main.py", select_purchases_file):
                return
            continue

        while True:
            category = select_from_list(
                list(BANK_CATEGORIES.keys()),
                "SELECT BANK CATEGORY",
                allow_back=True
            )

            if category == BACK:
                break

            while True:
                bank_ledger = select_from_list(
                    BANK_CATEGORIES[category],
                    "SELECT BANK",
                    allow_back=True
                )

                if bank_ledger == BACK:
                    break

                print("\nSelect the bank statement PDF...")
                file_path = select_file()

                if not file_path:
                    print("\nNo file selected. Exiting.")
                    sys.exit()

                print(f"\nSelected Bank: {bank_ledger}")
                print(f"Selected File: {file_path}")
                print("\nProcessing...\n")
                subprocess.run(["python", "main.py", file_path, bank_ledger])
                return


if __name__ == "__main__":
    main()
