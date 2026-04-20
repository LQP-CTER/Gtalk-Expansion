import openpyxl
from datetime import datetime

file_path = r"d:\LQP\Project\ghn-gtalk-user\[Gtalk Expansion] Workforce Analysis.xlsx"
print("Loading workbook...")
wb = openpyxl.load_workbook(file_path)
ws = wb["data history"]

print("Updating cells...")
for cell in ws[1]:
    val = cell.value
    if val is None:
        continue
        
    if isinstance(val, datetime):
        # Excel interpreted DD-MM as MM-DD
        # e.g. 01-04 (April 1st) became Jan 4th -> month=1, day=4
        dt_str = f"{val.month:02d}/{val.day:02d}/2026"
        cell.value = dt_str
        cell.number_format = '@' # enforce text format
    elif isinstance(val, str) and "-" in val:
        parts = val.split("-")
        if len(parts) == 2:
            dt_str = f"{parts[0]}/{parts[1]}/2026"
            cell.value = dt_str
            cell.number_format = '@'

print("Saving workbook... this may take a few seconds.")
wb.save(file_path)
print("Successfully fixed dates in Excel file!")
