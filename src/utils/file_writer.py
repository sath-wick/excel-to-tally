import os
import sys
from copy import copy

from openpyxl import load_workbook


def safe_excel_write(write_function, file_path, max_retries=3):
    """
    Attempts to execute write_function().
    If file is open (PermissionError), prompts user to close and retry.
    Retries only max_retries times, then exits with error.
    """

    attempts = 0

    while attempts < max_retries:
        try:
            write_function()
            _apply_standard_excel_format(file_path)
            return  # success

        except PermissionError:
            attempts += 1
            print(f"\nFile '{file_path}' is currently open.")
            print(f"Attempt {attempts} of {max_retries}.")
            
            if attempts >= max_retries:
                print("\nCould not proceed because the file is still open.")
                sys.exit(1)

            try:
                input("Please close the file and press Enter to retry...")
            except EOFError:
                print("Non-interactive mode. Waiting 3 seconds for file to be closed before retrying...")
                import time
                time.sleep(3)

        except Exception as e:
            print(f"\nUnexpected error while writing '{file_path}': {e}")
            sys.exit(1)


def _apply_standard_excel_format(file_path):
    if not _is_excel_file(file_path):
        return

    if not os.path.exists(file_path):
        return

    workbook = load_workbook(file_path)

    for worksheet in workbook.worksheets:
        max_row = worksheet.max_row
        max_col = worksheet.max_column

        for row in worksheet.iter_rows(
            min_row=1,
            max_row=max_row,
            min_col=1,
            max_col=max_col
        ):
            for cell in row:
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                cell.alignment = alignment

        for column_cells in worksheet.iter_cols(
            min_col=1,
            max_col=max_col,
            min_row=1,
            max_row=max_row
        ):
            column_letter = column_cells[0].column_letter
            max_length = 0

            for cell in column_cells:
                if cell.value is None:
                    continue
                max_length = max(max_length, len(str(cell.value)))

            worksheet.column_dimensions[column_letter].width = min(
                100,
                max(8, max_length + 2)
            )

    workbook.save(file_path)


def _is_excel_file(file_path):
    file_name = str(file_path).lower()
    return file_name.endswith((".xlsx", ".xlsm", ".xltx", ".xltm"))
