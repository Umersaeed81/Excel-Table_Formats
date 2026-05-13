## Import Required Libraries

## 🐍 pythoncom
- 🧠 Used for initializing and managing COM threading
- 🔄 Ensures proper communication between Python and Windows COM system
- ⚙️ Required for stable Excel automation when using COM objects
- 🧩 Prevents crashes and instability in multi-threaded Excel operations

## 📊 win32com.client

- 🔌 Provides interface to communicate with Microsoft Excel via Windows COM
- 🤖 Enables full Excel automation from Python
- 📂 Supports operations like:
    - 📁 Opening and managing workbooks
    - 📊 Creating and modifying Pivot Tables
    - 🎨 Formatting sheets and ranges
    - ⚡ Running Excel actions programmatically (fast, no manual work needed)


```python
import pythoncom
import win32com.client as win32
```

## EXCEL CONSTANTS (COM / WIN32 INTERNAL VALUES)


```python
# -----------------------------------------------------
# Pivot cache source type
# -----------------------------------------------------
# 1 = xlDatabase
# Used when pivot source is a worksheet range / table
XL_DATABASE = 1


# -----------------------------------------------------
# Pivot field orientations
# -----------------------------------------------------
# These define where a field appears in Pivot Table

# 1 = xlRowField
# Field appears in ROW area (vertical grouping)
XL_ROW_FIELD = 1

# 2 = xlColumnField
# Field appears in COLUMN area (horizontal grouping)
XL_COLUMN_FIELD = 2

# 3 = xlPageField (Filter area in modern Excel UI)
# Field is used as FILTER (top dropdown filter)
XL_PAGE_FIELD = 3


# -----------------------------------------------------
# Pivot layout style
# -----------------------------------------------------
# 1 = xlTabularRow
# Shows pivot in tabular (table-like) format
# - Each row field gets its own column
# - Easier for exporting / reporting
XL_TABULAR = 1


# -----------------------------------------------------
# Excel navigation directions (used in .End())
# -----------------------------------------------------

# -4162 = xlUp
# Moves upward to find last used cell in a column
XL_UP = -4162

# -4159 = xlToLeft
# Moves left to find last used cell in a row
XL_TO_LEFT = -4159


# -----------------------------------------------------
# Aggregation functions for Pivot Values
# -----------------------------------------------------
# These are Excel XlConsolidationFunction enum values

AGG_MAP = {

    # -4157 = xlSum
    # Adds all numeric values
    "sum": -4157,

    # -4112 = xlCount
    # Counts number of records (non-empty cells)
    "count": -4112,

    # -4106 = xlAverage
    # Calculates mean of values
    "average": -4106,

    # -4136 = xlMax
    # Returns maximum value
    "max": -4136,

    # -4139 = xlMin
    # Returns minimum value
    "min": -4139
}
```

## INITIALIZE EXCEL APPLICATION


```python
def init_excel():
    """
    Initialize Excel COM application.

    Returns
    -------
    excel : COM Object
        Excel application object with optimized settings.

    Notes
    -----
    - Excel runs in background mode
    - Screen updating is disabled for better performance
    - Alerts/events are disabled for smoother automation
    """

    # Initialize COM library
    pythoncom.CoInitialize()

    # Create isolated Excel instance
    excel = win32.DispatchEx("Excel.Application")

    # Performance optimization
    excel.Visible = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    return excel

```

## CLOSE EXCEL APPLICATION


```python
def close_excel(excel):
    """
    Properly close Excel application.

    Parameters
    ----------
    excel : COM Object
        Excel application instance.
    """

    if excel:
        excel.Quit()

    # Uninitialize COM library
    pythoncom.CoUninitialize()
```

## CONVERT COLUMN NUMBER TO LETTER


```python
def col_letter(col_num):
    """
    Convert Excel column number to column letter.

    Parameters
    ----------
    col_num : int
        Excel column number.

    Returns
    -------
    str
        Excel column letter.

    Examples
    --------
    1  -> A
    26 -> Z
    27 -> AA
    """

    result = ""

    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result

    return result
```

## GET LAST USED ROW


```python

def get_last_row(sheet, col):
    """
    Find last used row in a worksheet column.

    Parameters
    ----------
    sheet : Worksheet
        Excel worksheet object.

    col : int
        Column number to inspect.

    Returns
    -------
    int
        Last used row number.
    """

    return sheet.Cells(
        sheet.Rows.Count,
        col
    ).End(XL_UP).Row

```

## GET LAST USED COLUMN


```python
def get_last_col(sheet, row):
    """
    Find last used column in a worksheet row.

    Parameters
    ----------
    sheet : Worksheet
        Excel worksheet object.

    row : int
        Row number to inspect.

    Returns
    -------
    int
        Last used column number.
    """

    return sheet.Cells(
        row,
        sheet.Columns.Count
    ).End(XL_TO_LEFT).Column
```

## CLEAN HEADERS


```python
def clean_headers(sheet, header_row=1):
    """
    Remove extra spaces from Excel headers.

    Parameters
    ----------
    sheet : Worksheet
        Source worksheet.

    header_row : int, default=1
        Header row number.

    Notes
    -----
    This avoids pivot errors caused by:
    - Leading spaces
    - Trailing spaces
    """

    # Find last header column
    last_col = get_last_col(sheet, header_row)

    # Clean each header cell
    for col in range(1, last_col + 1):

        cell = sheet.Cells(header_row, col)

        if cell.Value:
            cell.Value = str(cell.Value).strip()

```

## BUILD SOURCE DATA RANGE


```python
def build_source_range(sheet, header_row=1, start_col=1):
    """
    Build Excel source range dynamically.

    Parameters
    ----------
    sheet : Worksheet
        Source worksheet.

    header_row : int, default=1
        Header row number.

    start_col : int, default=1
        Starting column number.

    Returns
    -------
    str
        Excel range reference.

    Example
    -------
    'Sheet1'!$A$1:$H$500
    """

    # Detect data boundaries
    last_row = get_last_row(sheet, start_col)
    last_col = get_last_col(sheet, header_row)

    # Build Excel range string
    return (
        f"'{sheet.Name}'!"
        f"${col_letter(start_col)}${header_row}:"
        f"${col_letter(last_col)}${last_row}"
    )

```

## DELETE EXISTING SHEET


```python
def delete_sheet_if_exists(wb, sheet_name):
    """
    Delete worksheet if already exists.

    Parameters
    ----------
    wb : Workbook
        Excel workbook object.

    sheet_name : str
        Worksheet name to delete.
    """

    for sh in wb.Worksheets:

        if sh.Name == sheet_name:
            sh.Delete()
            break

```

## CREATE PIVOT TABLE


```python
def create_pivot_table(
    wb,
    source_data,
    pivot_sheet_name,
    pivot_table_name
):
    """
    Create a new Excel pivot table.

    Parameters
    ----------
    wb : Workbook
        Excel workbook object.

    source_data : str
        Excel source data range.

    pivot_sheet_name : str
        Pivot worksheet name.

    pivot_table_name : str
        Pivot table name.

    Returns
    -------
    tuple
        (pivot_sheet, pivot_table)
    """

    # Remove old pivot sheet if exists
    delete_sheet_if_exists(wb, pivot_sheet_name)

    # Create new worksheet
    pivot_sheet = wb.Worksheets.Add()
    pivot_sheet.Name = pivot_sheet_name

    # Create pivot cache
    cache = wb.PivotCaches().Create(
        SourceType=XL_DATABASE,
        SourceData=source_data
    )

    # Create pivot table
    pivot_table = cache.CreatePivotTable(
        TableDestination=pivot_sheet.Range("A1"),
        TableName=pivot_table_name
    )

    return pivot_sheet, pivot_table

```

## ADD PIVOT FIELDS


```python
def add_fields(pivot_table, fields, orientation):
    """
    Add fields to pivot table.

    Parameters
    ----------
    pivot_table : PivotTable
        Excel pivot table object.

    fields : list
        List of field names.

    orientation : int
        Pivot field orientation.

    Notes
    -----
    Orientation types:
    - XL_ROW_FIELD
    - XL_COLUMN_FIELD
    - XL_PAGE_FIELD
    """

    for position, field_name in enumerate(fields or [], start=1):

        try:

            # Get pivot field
            field = pivot_table.PivotFields(field_name)

            # Apply orientation
            field.Orientation = orientation
            field.Position = position

            # Disable subtotals for row fields
            if orientation == XL_ROW_FIELD:
                field.Subtotals = (False,) * 12

        except Exception as e:
            print(f"Field Error ({field_name}): {e}")

```

## ADD VALUE FIELDS


```python
def add_value_fields(
    pivot_table,
    values,
    value_formats=None
):
    """
    Add value fields with aggregation.

    Parameters
    ----------
    pivot_table : PivotTable
        Excel pivot table object.

    values : dict
        Dictionary of value fields.

    value_formats : dict, optional
        Excel number formats.

    Example
    -------
    values = {
        "Traffic": ["sum", "average"],
        "Users": ["count"]
    }
    """

    for field_name, aggregations in (values or {}).items():

        for agg in aggregations:

            try:

                # Add data field
                data_field = pivot_table.AddDataField(
                    pivot_table.PivotFields(field_name),
                    f"{agg.title()} of {field_name}",
                    AGG_MAP.get(agg.lower(), -4106)
                )

                # Apply number format
                data_field.NumberFormat = (
                    value_formats.get(field_name, "0.00")
                    if value_formats else "0.00"
                )

            except Exception as e:
                print(f"Value Error ({field_name}): {e}")

```

## APPLY FILTER VALUES


```python
def apply_filter_values(
    pivot_table,
    filter_values
):
    """
    Apply filters to pivot fields.

    Parameters
    ----------
    pivot_table : PivotTable
        Excel pivot table object.

    filter_values : dict
        Dictionary containing filter selections.

    Examples
    --------
    Single filter:
    {"Zone": "South"}

    Multiple filters:
    {"Zone": ["South", "North"]}
    """

    for field_name, selected_values in (filter_values or {}).items():

        try:

            field = pivot_table.PivotFields(field_name)

            # Clear existing filters
            field.ClearAllFilters()

            # Single filter selection
            if isinstance(selected_values, str):

                field.CurrentPage = selected_values

            # Multiple filter selections
            else:

                field.EnableMultiplePageItems = True

                selected_values = set(map(str, selected_values))

                for item in field.PivotItems():

                    try:
                        item.Visible = (
                            str(item.Name) in selected_values
                        )
                    except:
                        pass

        except Exception as e:
            print(f"Filter Error ({field_name}): {e}")


```

## SORT PIVOT FIELDS


```python
def sort_pivot_fields(
    pivot_table,
    sort_fields
):
    """
    Sort pivot table fields.

    Parameters
    ----------
    pivot_table : PivotTable
        Excel pivot table object.

    sort_fields : dict
        Dictionary of sorting configuration.

    Example
    -------
    {
        "Zone": "ascending",
        "Traffic": "descending"
    }
    """

    for field_name, order in (sort_fields or {}).items():

        try:

            field = pivot_table.PivotFields(field_name)

            field.AutoSort(
                1 if order.lower() == "ascending" else 2,
                field_name
            )

        except Exception as e:
            print(f"Sort Error ({field_name}): {e}")

```

## UPDATE PIVOT LAYOUT


```python
def update_pivot_layout(
    pivot_table,
    repeat_item_labels=True,
    grand_total="both",
    subtotals=True
):
    """
    Configure pivot table layout settings.

    Parameters
    ----------
    pivot_table : PivotTable
        Excel pivot table object.

    repeat_item_labels : bool, default=True
        Repeat row labels.

    grand_total : str, default="both"
        Grand total visibility.

        Options:
        - both
        - rows
        - columns
        - none

    subtotals : bool, default=True
        Enable/disable subtotals.
    """

    # Apply tabular layout
    pivot_table.RowAxisLayout(XL_TABULAR)

    # Enable classic pivot layout
    pivot_table.InGridDropZones = True
    pivot_table.DisplayFieldCaptions = True

    # Repeat row labels
    for i in range(1, pivot_table.RowFields.Count + 1):

        try:
            pivot_table.RowFields.Item(i).RepeatLabels = (
                repeat_item_labels
            )
        except:
            pass

    # Grand total settings
    grand_map = {
        "both": (True, True),
        "rows": (True, False),
        "columns": (False, True),
        "none": (False, False)
    }

    pivot_table.RowGrand, pivot_table.ColumnGrand = (
        grand_map.get(grand_total, (True, True))
    )

    # Subtotal settings
    for i in range(1, pivot_table.PivotFields().Count + 1):

        try:

            field = pivot_table.PivotFields().Item(i)

            field.Subtotals = [subtotals] * 12

        except:
            pass

```

## MAIN PIVOT TABLE ENGINE


```python
def create_pivot_table_com(

    file_path,
    source_sheet_name,

    pivot_sheet_name="Pivot_Table",
    pivot_table_name="PivotTable1",

    filters=None,
    rows=None,
    columns=None,

    values=None,

    filter_values=None,
    sort_fields=None,
    value_formats=None,

    repeat_item_labels=True,
    grand_total="both",
    subtotals=True,

    header_row=1,
    start_col=1
):
    """
    Fully automated Excel Pivot Table creation engine.

    Parameters
    ----------
    file_path : str
        Excel file path.

    source_sheet_name : str
        Source worksheet name.

    pivot_sheet_name : str, default="Pivot_Table"
        Pivot worksheet name.

    pivot_table_name : str, default="PivotTable1"
        Pivot table name.

    filters : list, optional
        Pivot filter fields.

    rows : list, optional
        Pivot row fields.

    columns : list, optional
        Pivot column fields.

    values : dict, optional
        Pivot value fields with aggregations.

    filter_values : dict, optional
        Filter selections.

    sort_fields : dict, optional
        Sorting configuration.

    value_formats : dict, optional
        Excel number formats.

    repeat_item_labels : bool, default=True
        Repeat row labels.

    grand_total : str, default="both"
        Grand total configuration.

    subtotals : bool, default=True
        Enable/disable subtotals.

    header_row : int, default=1
        Header row number.

    start_col : int, default=1
        Starting column number.

    Notes
    -----
    This function performs:
    - Excel automation
    - Dynamic source detection
    - Pivot table creation
    - Field placement
    - Filtering
    - Sorting
    - Formatting
    - Layout customization
    """

    excel = None

    try:

        # =================================================
        # INITIALIZE EXCEL
        # =================================================
        excel = init_excel()

        # =================================================
        # OPEN WORKBOOK
        # =================================================
        wb = excel.Workbooks.Open(file_path)

        # =================================================
        # ACCESS SOURCE SHEET
        # =================================================
        source_sheet = wb.Worksheets(source_sheet_name)

        # =================================================
        # CLEAN HEADERS
        # =================================================
        clean_headers(source_sheet, header_row)

        # =================================================
        # BUILD SOURCE RANGE
        # =================================================
        source_data = build_source_range(
            source_sheet,
            header_row,
            start_col
        )

        # =================================================
        # CREATE PIVOT TABLE
        # =================================================
        pivot_sheet, pivot_table = create_pivot_table(
            wb,
            source_data,
            pivot_sheet_name,
            pivot_table_name
        )

        # =================================================
        # ADD FILTER FIELDS
        # =================================================
        add_fields(
            pivot_table,
            filters,
            XL_PAGE_FIELD
        )

        # =================================================
        # ADD ROW FIELDS
        # =================================================
        add_fields(
            pivot_table,
            rows,
            XL_ROW_FIELD
        )

        # =================================================
        # ADD COLUMN FIELDS
        # =================================================
        add_fields(
            pivot_table,
            columns,
            XL_COLUMN_FIELD
        )

        # =================================================
        # ADD VALUE FIELDS
        # =================================================
        add_value_fields(
            pivot_table,
            values,
            value_formats
        )

        # =================================================
        # APPLY FILTERS
        # =================================================
        apply_filter_values(
            pivot_table,
            filter_values
        )

        # =================================================
        # APPLY SORTING
        # =================================================
        sort_pivot_fields(
            pivot_table,
            sort_fields
        )

        # =================================================
        # APPLY LAYOUT SETTINGS
        # =================================================
        update_pivot_layout(
            pivot_table,
            repeat_item_labels,
            grand_total,
            subtotals
        )

        # =================================================
        # REFRESH PIVOT
        # =================================================
        pivot_table.RefreshTable()

        # =================================================
        # AUTO FIT COLUMNS
        # =================================================
        pivot_sheet.Columns.AutoFit()

        # =================================================
        # SAVE & CLOSE WORKBOOK
        # =================================================
        wb.Save()
        wb.Close(True)

        print("Pivot table created successfully.")

    finally:

        # =================================================
        # CLOSE EXCEL
        # =================================================
        close_excel(excel)
```

## EXAMPLE: CREATE + FORMAT PIVOT TABLE


```python
create_pivot_table_com(

    # =================================================
    # FILE + SOURCE INFORMATION
    # =================================================

    # 📂 Excel file path (source dataset)
    file_path=r"D:\Advance_Data_Sets\FSD_MOCN\Cell_List\GSM_Test.xlsx",

    # 📄 Source worksheet containing raw KPI data
    source_sheet_name="GSM_KPIs",

    # =================================================
    # PIVOT TABLE INFORMATION
    # =================================================

    # 📊 Name of the new sheet where Pivot Table will be created
    pivot_sheet_name="Pivot_CSSR",

    # 🧾 Name of the Pivot Table object inside Excel
    pivot_table_name="PT_CSSR",

    # =================================================
    # FILTER FIELDS
    # =================================================

    # 🔽 Fields used as Pivot Filters
    filters=[
        "Zone",
        "Date_Mapped"
    ],

    # =================================================
    # FILTER VALUES
    # =================================================

    # 🎯 Default filter selections
    filter_values={

        # Multiple selection filter
        "Zone": [
            "Buffer_Tier_0",
            "Buffer_Tier_1"
        ],

        # Multiple selection filter
        "Date_Mapped": [
            "Pre",
            "Post"
        ]
    },

    # =================================================
    # ROW FIELDS
    # =================================================

    # 📌 Fields displayed vertically
    rows=[
        "Zone",
        "Time"
    ],

    # =================================================
    # COLUMN FIELDS
    # =================================================

    # 📊 Fields displayed horizontally
    columns=[
        "Date"
    ],

    # =================================================
    # VALUE FIELDS
    # =================================================

    # 📈 KPI aggregation logic
    #
    # Supported:
    # - sum
    # - average
    # - count
    # - max
    # - min
    # =================================================

    values={

        # KPI 1
        "CSSR": [
            "sum"
        ],

        # KPI 2
        "DCR": [
            "average"
        ]
    },

    # =================================================
    # SORTING
    # =================================================

    # 🔢 Sorting rules
    #
    # Supported:
    # - ascending
    # - descending
    # =================================================

    sort_fields={
        # Latest dates first
        "Date": "ascending",
        # Reverse time order
        "Time": "ascending"
    },

    # =================================================
    # VALUE FORMATTING
    # =================================================

    # # =====================================================
    # # DECIMAL FORMATS
    # # =====================================================
    
    # "0"              # integer
    # "0.0"            # 1 decimal
    # "0.00"           # 2 decimals
    # "0.000"          # 3 decimals
    
    # # =====================================================
    # # PERCENTAGE FORMATS
    # # =====================================================
    
    # "0%"             # percentage no decimal
    # "0.0%"           # percentage 1 decimal
    # "0.00%"          # percentage 2 decimals
    
    # # =====================================================
    # # COMMA SEPARATED FORMATS
    # # =====================================================
    
    # "#,##0"          # 1,000
    # "#,##0.00"       # 1,000.00
    
    # # =====================================================
    # # CURRENCY FORMATS
    # # =====================================================
    
    # "$#,##0"         # $1,000
    # "$#,##0.00"      # $1,000.00
    
    # # =====================================================
    # # NEGATIVE NUMBER FORMATS
    # # =====================================================
    
    # "#,##0;(#,##0)"                 # negatives in brackets
    # "#,##0.00;(#,##0.00)"          # decimals + brackets
    
    # # =====================================================
    # # CONDITIONAL COLORS
    # # =====================================================
    
    # "[Green]0.00%;[Red]-0.00%"     # positive green, negative red
    
    # # =====================================================
    # # SCIENTIFIC FORMAT
    # # =====================================================
    
    # "0.00E+00"       # scientific notation
    
    # # =====================================================
    # # DATE FORMATS
    # # =====================================================
    
    # "dd-mm-yyyy"
    # "dd-mmm-yyyy"
    # "yyyy-mm-dd"
    
    # # =====================================================
    # # TIME FORMATS
    # # =====================================================
    
    # "hh:mm"
    # "hh:mm:ss"
    
    # # =====================================================
    # # TEXT FORMAT
    # # =====================================================
    
    # "@"              # force text
    # # =================================================

    value_formats={

        "CSSR": "0.000",
        "DCR": "0.000"
    },

    # =================================================
    # LAYOUT OPTIONS
    # =================================================

    # 🔁 Repeat row labels
    repeat_item_labels=True,

    # =================================================
    # GRAND TOTAL OPTIONS
    # =================================================

    # Supported:
    # -----------------------------------------------
    # "both"
    # "rows"
    # "columns"
    # "none"
    # =================================================

    grand_total="both",

    # =================================================
    # SUBTOTAL OPTIONS
    # =================================================

    # True  → Enable subtotals
    # False → Disable subtotals
    # =================================================

    subtotals=False,

    # =================================================
    # SOURCE DATA OPTIONS
    # =================================================

    # 📅 Header row position in source data
    header_row=5,

    # 📍 Starting column index
    start_col=3
)
```


```python
import os
os.system("taskkill /f /im excel.exe")
```


```python
# re-set all the variable from the RAM
%reset -f
```
