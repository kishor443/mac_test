@echo off
echo Opening Excel file...
cd /d "%~dp0"
start excel.exe "data\activity_log.xlsx"
echo.
echo If file doesn't open, run: python fix_excel_now.py
pause

