"""
Utility script to repair or recreate corrupted Excel file.
Run this script if your Excel file is not opening properly.
"""

from pathlib import Path
from utils.excel_storage import EXCEL_PATH, _create_new_workbook, _ensure_workbook
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

def repair_excel_file():
    """Repair or recreate the Excel file if it's corrupted."""
    print("Checking Excel file...")
    
    if not EXCEL_PATH.exists():
        print("Excel file does not exist. Creating new file...")
        _create_new_workbook()
        print("[OK] New Excel file created successfully!")
        return
    
    # Try to open the file
    try:
        wb = load_workbook(EXCEL_PATH, data_only=False, keep_links=False)
        print("[OK] Excel file is valid and can be opened!")
        wb.close()
        return
    except (BadZipFile, InvalidFileException, KeyError, OSError, IOError) as e:
        print(f"[ERROR] Excel file is corrupted: {e}")
        print("Creating backup and recreating file...")
        
        # Backup corrupted file
        backup_path = EXCEL_PATH.with_suffix('.corrupted_backup.xlsx')
        try:
            EXCEL_PATH.rename(backup_path)
            print(f"  Backed up corrupted file to: {backup_path}")
        except Exception as backup_error:
            print(f"  Warning: Could not backup file: {backup_error}")
            try:
                EXCEL_PATH.unlink()
            except Exception:
                pass
        
        # Create new file
        try:
            _create_new_workbook()
            print("[OK] New Excel file created successfully!")
            print("  Note: Previous data was lost. Check backup file if needed.")
        except Exception as create_error:
            print(f"[ERROR] Error creating new file: {create_error}")
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")

if __name__ == "__main__":
    repair_excel_file()

