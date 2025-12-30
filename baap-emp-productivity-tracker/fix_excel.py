"""
Utility script to fix existing Excel file by updating headers, column widths, and formatting.
Run this script once to migrate your existing activity_log.xlsx to the new format.
"""

from utils.excel_storage import fix_existing_excel_file

if __name__ == "__main__":
    print("Fixing existing Excel file...")
    try:
        fix_existing_excel_file()
        print("[OK] Excel file has been updated successfully!")
        print("  - Column widths have been increased")
        print("  - Headers have been formatted")
        print("  - Missing columns have been added (webcam_photo)")
        print("  - metadata_json column width increased to 120 characters")
    except Exception as e:
        print(f"[X] Error fixing Excel file: {e}")

