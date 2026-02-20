import sys
import time


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
            return  # success

        except PermissionError:
            attempts += 1
            print(f"\nFile '{file_path}' is currently open.")
            print(f"Attempt {attempts} of {max_retries}.")
            
            if attempts >= max_retries:
                print("\nCould not proceed because the file is still open.")
                sys.exit(1)

            input("Please close the file and press Enter to retry...")

        except Exception as e:
            print(f"\nUnexpected error while writing '{file_path}': {e}")
            sys.exit(1)
