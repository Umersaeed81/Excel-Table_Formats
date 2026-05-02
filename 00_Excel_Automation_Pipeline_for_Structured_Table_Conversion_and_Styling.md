# Excel Automation Pipeline for Structured Table Conversion and Styling

## IMPORT REQUIRED LIBRARIES


```python
import os                                                                 # For file and directory operations
import re                                                                 # For pattern matching and string cleaning
import pandas as pd                                                       # For data manipulation and analysis
import win32com.client as win32                                           # For Excel automation using COM interface (Windows only)
from openpyxl import load_workbook                                        # For reading/modifying Excel files
from openpyxl.worksheet.table import Table, TableStyleInfo                # For creating styled Excel tables


# -----------------------------
# PURPOSE OF THIS SCRIPT
# -----------------------------
# This script is designed to:
# 1. Handle file system operations (os)
# 2. Clean and manipulate text patterns (re)
# 3. Process tabular data (pandas)
# 4. Automate Excel using COM (win32com)
# 5. Enhance Excel files with structured tables and styling (openpyxl)


# -----------------------------
# NOTES
# -----------------------------
# - win32com requires Microsoft Excel installed (Windows only)
# - openpyxl works independently without Excel GUI
# - This combination is powerful for:
#     • Report automation
#     • Dashboard generation
#     • Excel formatting and table creation
#     • Batch processing of Excel files
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

## SAFE TABLE NAME CLEANER


```python
def clean_table_name(name):
    """
    Cleans and standardizes a string to make it a valid Excel table name.

    Excel table naming rules:
    - Must start with a letter or underscore
    - Cannot contain spaces or special characters
    - Should only include letters, numbers, and underscores

    Parameters:
        name (str): Original table name (may contain invalid characters)

    Returns:
        str: A cleaned and Excel-compatible table name
    """

    # Convert input to string and replace all invalid characters
    # Only allows letters, numbers, and underscore (_)
    name = re.sub(r'[^A-Za-z0-9_]', '_', str(name))

    # Ensure name is not empty and starts with a letter
    # If it starts with a number or is empty, prefix with "T_"
    if not name or not name[0].isalpha():
        name = "T_" + name

    return name
```

## MAIN FUNCTION


```python
def convert_to_excel_table(
    file_path,
    output_path=None,

    sheets=None,
    table_names=None,

    style_name="TableStyleMedium9",

    show_first_column=False,
    show_last_column=False,
    show_row_stripes=True,
    show_column_stripes=False,

    auto_prefix=True,
    overwrite=False
):
    """
    Converts Excel worksheets into structured Excel Tables with styling and saves the file.

    This function:
    - Loads an Excel file
    - Converts selected sheets into Excel Tables
    - Applies consistent table styling
    - Allows custom table naming per sheet
    - Supports output file control (overwrite or new file)

    Parameters:
        file_path (str):
            Path of the input Excel file.

        output_path (str, optional):
            Custom output file path. If not provided, a default "_table" file is created.

        sheets (str or list, optional):
            Sheet name(s) to process.
            - None → all sheets
            - str → single sheet
            - list → multiple sheets

        table_names (dict, optional):
            Custom table names per sheet:
            Example: {"Sheet1": "SalesTable"}

        style_name (str):
            Excel table style name (default: TableStyleMedium9)

        show_first_column (bool):
            Highlight first column in table style.

        show_last_column (bool):
            Highlight last column in table style.

        show_row_stripes (bool):
            Enable alternating row shading.

        show_column_stripes (bool):
            Enable alternating column shading.

        auto_prefix (bool):
            If True, adds "tbl_" prefix to table names.

        overwrite (bool):
            If True, overwrites original file.

    Returns:
        None
        (Saves the modified Excel file directly)

    """

    # -----------------------------
    # LOAD EXCEL WORKBOOK
    # -----------------------------
    wb = load_workbook(file_path)

    # -----------------------------
    # SHEET SELECTION LOGIC
    # -----------------------------
    # If no sheet is provided → process all sheets
    if sheets is None:
        target_sheets = wb.sheetnames

    # If single sheet name is provided → convert to list
    elif isinstance(sheets, str):
        target_sheets = [sheets]

    # If list of sheets provided → use as-is
    else:
        target_sheets = sheets

    # -----------------------------
    # PROCESS EACH SHEET
    # -----------------------------
    for sheet_name in target_sheets:

        # Skip if sheet does not exist
        if sheet_name not in wb.sheetnames:
            print(f"⚠️ Missing sheet: {sheet_name}")
            continue

        ws = wb[sheet_name]

        # Remove existing tables to avoid conflicts
        ws._tables.clear()

        # Get data boundaries
        last_row = ws.max_row
        last_col = ws.max_column

        # Skip empty sheets
        if last_row < 2:
            continue

        # Define full data range for table
        table_ref = f"A1:{ws.cell(row=last_row, column=last_col).coordinate}"

        # -----------------------------
        # TABLE NAME HANDLING
        # -----------------------------
        # Use custom name if provided, otherwise sheet name
        base_name = (
            table_names[sheet_name]
            if table_names and sheet_name in table_names
            else sheet_name
        )

        # Clean table name to make it Excel-safe
        base_name = clean_table_name(base_name)

        # Add prefix if enabled
        table_name = f"tbl_{base_name}" if auto_prefix else base_name

        # -----------------------------
        # CREATE EXCEL TABLE
        # -----------------------------
        table = Table(displayName=table_name, ref=table_ref)

        # Define table styling options
        style = TableStyleInfo(
            name=style_name,
            showFirstColumn=show_first_column,
            showLastColumn=show_last_column,
            showRowStripes=show_row_stripes,
            showColumnStripes=show_column_stripes
        )

        # Attach style to table
        table.tableStyleInfo = style

        # Add table to worksheet
        ws.add_table(table)

        print(f"✅ {sheet_name} → {table_name}")

    # -----------------------------
    # SAVE FILE LOGIC
    # -----------------------------
    if overwrite:
        # Overwrite original file
        save_path = file_path

    elif output_path:
        # Save to user-defined path
        save_path = output_path

    else:
        # Default: create new file with suffix "_table"
        base, ext = os.path.splitext(file_path)
        save_path = f"{base}_table{ext}"

    # Save workbook
    wb.save(save_path)

    print(f"\n📁 Saved: {save_path}")
```

## EXECUTION: CONVERT EXCEL SHEETS INTO FORMATTED TABLES


```python
convert_to_excel_table(

    # ==========================================
    # INPUT FILE
    # ==========================================
    "CS_Data_Traffic.xlsx",   # Input Excel file
    # ==========================================
    # SHEET SELECTION
    # ==========================================
    #sheets=["NW_Traffic", "C1_Traffic"],
    sheets= None,
    # Convert only these sheets into Excel tables

    # sheets="NW_Traffic"
    # Convert only one sheet

    # sheets=None
    # Convert ALL sheets in workbook


    # ==========================================
    # TABLE STYLE
    # ==========================================
    style_name="TableStyleMedium2",

    # ===============================
    # EXCEL TABLE STYLE OPTIONS
    # (openpyxl / Excel built-in)
    # ===============================

    # LIGHT STYLES
    # TableStyleLight1   to TableStyleLight21

    # MEDIUM STYLES (Recommended)
    # TableStyleMedium1 to TableStyleMedium28

    # DARK STYLES
    # TableStyleDark1   to TableStyleDark11

    # NOTE:
    # - Must use exact names
    # - openpyxl supports built-in Excel styles only


    # ==========================================
    # TABLE VISUAL OPTIONS
    # ==========================================

    show_first_column=True,
    # True  = highlight first column
    # False = normal formatting

    show_last_column=True,
    # True  = highlight last column
    # False = normal formatting

    show_row_stripes=True,
    # True  = alternating row colors (banded rows)
    # False = no row stripes

    show_column_stripes=True,
    # True  = alternating column colors (banded columns)
    # False = no column stripes
    # ==========================================
    # TABLE NAME CONTROL
    # ==========================================
    table_names={
        "NW_Traffic": "NW_CS_Traffic",
        "C1_Traffic": "C1_CS_Traffic"
    },
    # Custom table names for specific sheets

    # If table_names=None:
    # Table name automatically becomes:
    # tbl_<sheet_name>

    # Example:
    # NW_Traffic -> tbl_NW_Traffic
    # C1_Traffic -> tbl_C1_Traffic


    # ==========================================
    # OUTPUT FILE CONTROL
    # ==========================================

    # output_path="Final_Report.xlsx",
    # Save as a new custom file

    # If output_path is not given:
    # Automatically saves as:
    # CS_Data_Traffic_table.xlsx


    # overwrite=True
    # Save changes in SAME original file
    # Warning: original file will be overwritten


)
```


