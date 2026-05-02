# Excel Table Removal and Workbook Cleanup Automation using Python

## IMPORT REQUIRED LIBRARIES


```python
import os                                                                 # For file and directory operations
from openpyxl import load_workbook                                        # For reading/modifying Excel files
```

## SET WORKING DIRECTORY


```python
# Define the folder path where all input/output files are located
path = 'C:/New folder'

# Change the current working directory to the specified path
# This ensures that all relative file operations (read/write) happen in this folder
os.chdir(path)

# -----------------------------
# PURPOSE
# -----------------------------
# - Centralizes file access location
# - Avoids writing full file paths repeatedly
# - Useful in automation scripts handling multiple Excel/files
# - Ensures consistency in file reading/writing operations
```

## Excel Table Removal Automation with Flexible Sheet and Save Control


```python
def remove_excel_tables(
    file_path,
    output_path=None,

    sheets=None,
    overwrite=False
):
    """
    Removes Excel Tables from worksheets and converts them back to normal ranges.

    Supports:
    - All sheets
    - Single sheet
    - List of sheets
    - Custom output / overwrite logic

    Parameters:
        file_path (str): Input Excel file
        output_path (str): Custom output file path
        sheets (None | str | list): Sheets to process
        overwrite (bool): Overwrite original file
    """

    # -----------------------------
    # LOAD WORKBOOK
    # -----------------------------
    wb = load_workbook(file_path)

    # -----------------------------
    # SHEET SELECTION LOGIC
    # -----------------------------
    if sheets is None:
        target_sheets = wb.sheetnames

    elif isinstance(sheets, str):
        target_sheets = [sheets]

    else:
        target_sheets = sheets

    # -----------------------------
    # PROCESS SHEETS
    # -----------------------------
    for sheet_name in target_sheets:

        if sheet_name not in wb.sheetnames:
            print(f"⚠️ Missing sheet: {sheet_name}")
            continue

        ws = wb[sheet_name]

        # -----------------------------
        # REMOVE ALL TABLES
        # -----------------------------
        if ws._tables:
            ws._tables.clear()
            print(f"✅ Tables removed from: {sheet_name}")
        else:
            print(f"ℹ️ No tables found in: {sheet_name}")

    # -----------------------------
    # SAVE LOGIC (same as your script)
    # -----------------------------
    if overwrite:
        save_path = file_path

    elif output_path:
        save_path = output_path

    else:
        base, ext = os.path.splitext(file_path)
        save_path = f"{base}_clean{ext}"

    wb.save(save_path)

    print(f"\n📁 Saved: {save_path}")
```

## EXCEL TABLE REMOVAL - USAGE EXAMPLES


```python
# =========================================================
# EXCEL TABLE REMOVAL - USAGE EXAMPLES
# =========================================================

# INPUT FILE:
# "CS_Data_Traffic_table.xlsx"
# → Excel file containing structured tables

# =========================================================
# SHEET SELECTION OPTIONS
# =========================================================

# sheets="NW_Traffic"
# → Process ONLY one sheet

# sheets=["s1", "s2"]
# → Process SPECIFIC list of sheets

# sheets=None
# → Process ALL sheets in workbook

# =========================================================
# FUNCTION CALL EXAMPLE
# =========================================================

remove_excel_tables(
    "CS_Data_Traffic_table.xlsx",
    sheets="NW_Traffic"   # 1 sheet only
)

# =========================================================
# OUTPUT / SAVE OPTIONS (same logic as function)
# =========================================================

# overwrite=True
# → Save back into SAME input file
#   CS_Data_Traffic_table.xlsx will be replaced

# output_path="Clean_File.xlsx"
# → Save as NEW custom file name

# default (if neither given)
# → Automatically saves as:
#   CS_Data_Traffic_table_clean.xlsx
```
