# 📊 Excel Pivot Table Automation Using Python COM (End-to-End Engine)
- 🧠 This code builds a complete Excel automation system using Python COM interface to generate Pivot Tables dynamically
- ⚙️ It removes manual Excel work by fully automating data range detection, Pivot creation, and formatting
- 📂 Works directly with Excel files using `win32com.client` for deep Excel control

- 🚀 Initializes Excel in hidden mode for fast processing and disables alerts/events for smooth automation
- 📊 Dynamically detects dataset boundaries (last row and last column) to avoid hardcoding ranges
- 🔄 Automatically builds Excel source range for Pivot Table creation

- 🧩 Creates and manages Pivot Table structure using Pivot Cache and assigns it to a new worksheet
- 🗂️ Deletes existing Pivot sheets to avoid duplication and ensures clean output each time

📌 Supports flexible Pivot configuration including:

    - 📍 Row fields
    - 📍 Column fields
    - 📍 Filter fields
    - 📍 Multiple value aggregations (sum, average, min, max, count)

- 📊 Enables advanced analytics by allowing multiple aggregations per metric in a single Pivot Table
- 🎯 Applies filtering logic including single-select and multi-select filter support

- 📉 Provides sorting functionality for Pivot fields in ascending or descending order
- 🎨 Controls Pivot layout styles (compact, outline, tabular) for better readability

- 📐 Applies formatting enhancements like number formats, grand totals, and repeated labels
- 🔄 Refreshes Pivot Table automatically after configuration changes

- 💾 Finalizes output by auto-fitting columns, saving workbook, and closing Excel safely
- 🧹 Ensures proper cleanup of COM objects to avoid memory leaks or hanging Excel processes

- 🧾 Overall, this script acts as a fully automated Pivot Table generation engine that transforms raw Excel data into structured, formatted, and analysis-ready reports without manual effort 🚀

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

## 🟢 `init_excel()`

- Starts Excel using COM automation
- Runs Excel in background (hidden mode)
- Disables screen updates, alerts, and events
- ⚡ **Purpose**: Creates a fast and stable Excel automation session


```python
# =====================================================
# INIT EXCEL APPLICATION
# =====================================================
def init_excel():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    return excel
```

##🔵 `col_letter(n)`

- Converts column number to Excel letters (1 → A, 27 → AA)
- Uses loop-based base-26 conversion
- ⚡ **Purpose**: Helps build dynamic Excel ranges


```python
# =====================================================
# COLUMN NUMBER TO LETTER
# =====================================================
def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

```

## 🟡 `get_last_row(sheet, start_col)`

- Finds last filled row in a given column
- Uses Excel’s .End() method (like Ctrl+Down)
-⚡ **Purpose**: Detects dataset height automatically


```python
# =====================================================
# GET LAST ROW
# =====================================================
def get_last_row(sheet, start_col):
    return sheet.Cells(sheet.Rows.Count, start_col).End(-4162).Row

```

## 🟣 `get_last_col(sheet, header_row, start_col)`

-  Finds last used column in header row
-  Uses .End() method (like Ctrl+Right)
- ⚡ **Purpose**: Detects dataset width dynamically


```python
# =====================================================
# GET LAST COLUMN
# =====================================================
def get_last_col(sheet, header_row, start_col):
    return sheet.Cells(header_row, sheet.Columns.Count).End(-4159).Column
```

## 🟠 `build_source_range(sheet, header_row, start_col)`

- Calculates last row and last column
- Converts column numbers to letters
- Builds Excel range string like `'Sheet1'!$A$1:$F$500`
- ⚡ **Purpose**: Defines Pivot source range dynamically


```python
# =====================================================
# BUILD SOURCE RANGE
# =====================================================
def build_source_range(sheet, header_row, start_col):

    last_row = get_last_row(sheet, start_col)
    last_col = get_last_col(sheet, header_row, start_col)

    start_col_letter = col_letter(start_col)
    last_col_letter = col_letter(last_col)

    return (
        f"'{sheet.Name}'!"
        f"${start_col_letter}${header_row}:"
        f"${last_col_letter}${last_row}"
    )

```

## 🔴 `delete_sheet_if_exists(wb, name)`

- Loops through workbook sheets
- Deletes sheet if name matches
- ⚡ **Purpose**: Prevents duplicate Pivot sheets


```python
# =====================================================
# DELETE SHEET IF EXISTS
# =====================================================
def delete_sheet_if_exists(wb, name):
    for sh in wb.Worksheets:
        if sh.Name == name:
            sh.Delete()
            return
```

## 🟤 `create_pivot(wb, source_data, sheet, table_name)`

- Creates Pivot Cache from source data
- Builds Pivot Table on target sheet
- Assigns table name
- ⚡ **Purpose**: Core Pivot Table creation engine


```python
# =====================================================
# CREATE PIVOT
# =====================================================
def create_pivot(wb, source_data, sheet, table_name):

    cache = wb.PivotCaches().Create(
        SourceType=1,
        SourceData=source_data
    )

    return cache.CreatePivotTable(
        TableDestination=sheet.Range("A1"),
        TableName=table_name
    )
```

## ⚫ `add_fields(pivot_table, fields, orientation)`

- Adds fields to Pivot areas (rows/columns/filters)
- Sets field position order
- Disables subtotals for row fields
- ⚡ **Purpose**: Structures Pivot layout


```python
# =====================================================
# ADD FIELDS
# =====================================================
def add_fields(pivot_table, fields, orientation):

    pos = 1

    for f in fields or []:

        try:
            pf = pivot_table.PivotFields(f)
            pf.Orientation = orientation
            pf.Position = pos

            if orientation == 1:
                pf.Subtotals = (False,) * 12

            pos += 1

        except:
            pass
```

## 🟢 `add_multi_values(pivot_table, values_dict, value_formats=None)`

- Loops through fields and aggregation types
- Supports multiple aggregations (sum, avg, min, max, count)
- Creates multiple value fields per metric
- Applies number formatting if provided
- ⚡ **Purpose**: Enables advanced KPI analysis in Pivot


```python
# =====================================================
#  MULTI AGG VALUES SUPPORT
# =====================================================
def add_multi_values(pivot_table, values_dict, value_formats=None):

    agg_map = {
        "sum": -4157,
        "count": -4112,
        "average": -4106,
        "max": -4136,
        "min": -4139
    }

    for field, aggs in (values_dict or {}).items():

        for agg in aggs:

            try:
                agg_code = agg_map.get(agg.lower(), -4106)

                df = pivot_table.AddDataField(
                    pivot_table.PivotFields(field),
                    f"{agg.title()} of {field}",
                    agg_code
                )

                fmt = "0.00"
                if value_formats and field in value_formats:
                    fmt = value_formats[field]

                df.NumberFormat = fmt

            except:
                pass

```

## 🔵 `apply_filter_values(pivot_table, filters, filter_values)`

- Clears existing filters
- Applies single or multi-select filters
- Controls Pivot item visibility
- ⚡ **Purpose**: Automates filtering logic


```python
# =====================================================
# APPLY FILTERS
# =====================================================
def apply_filter_values(pivot_table, filters, filter_values):

    for f in filters or []:

        if f not in (filter_values or {}):
            continue

        try:
            pf = pivot_table.PivotFields(f)
            pf.ClearAllFilters()

            selected = filter_values[f]

            if isinstance(selected, str):
                pf.CurrentPage = selected

            elif isinstance(selected, (list, tuple, set)):

                pf.EnableMultiplePageItems = True
                selected = set(map(str, selected))

                for item in pf.PivotItems():
                    item.Visible = True

                for item in pf.PivotItems():
                    try:
                        item.Visible = str(item.Name) in selected
                    except:
                        pass

        except:
            pass

```

## 🟡 `sort_pivot_field(pivot_table, field_name, sort_order)`

- Selects Pivot field
- Applies ascending or descending sort
- ⚡ **Purpose**: Controls data ordering in Pivot


```python
# =====================================================
# SORT FUNCTION
# =====================================================
def sort_pivot_field(pivot_table, field_name, sort_order="ascending"):

    try:
        pf = pivot_table.PivotFields(field_name)
        order = 1 if sort_order == "ascending" else 2
        pf.AutoSort(order, field_name)
    except:
        pass
```

## 🟣 `apply_pivot_layout(pivot_table, layout, grand_totals, repeat_labels)`

- Sets layout style (compact, outline, tabular)
- Enables/disables grand totals
- Repeats labels if needed
- ⚡ **Purpose**: Improves Pivot readability and formatting


```python
# =====================================================
# PIVOT LAYOUT CONTROL
# =====================================================
def apply_pivot_layout(
    pivot_table,
    layout="compact",
    grand_totals=True,
    repeat_labels=True
):

    try:
        layout_map = {
            "compact": 1,
            "outline": 2,
            "tabular": 3
        }

        try:
            pivot_table.RowAxisLayout(layout_map.get(layout, 1))
        except:
            pass

        try:
            pivot_table.RowGrand = grand_totals
            pivot_table.ColumnGrand = grand_totals
        except:
            pass

        try:
            if repeat_labels:
                pivot_table.RepeatAllLabels(True)
        except:
            pass

    except:
        pass

```

## 🟠 `build_pivot(wb, source_data, pivot_sheet_name, table_name)`

- Deletes existing Pivot sheet if present
- Creates new worksheet
- Initializes Pivot Table
- ⚡ **Purpose**: Prepares fresh Pivot workspace


```python
# =====================================================
# BUILD PIVOT SHEET
# =====================================================
def build_pivot(wb, source_data, pivot_sheet_name, table_name):

    delete_sheet_if_exists(wb, pivot_sheet_name)

    sheet = wb.Worksheets.Add()
    sheet.Name = pivot_sheet_name

    pivot_table = create_pivot(wb, source_data, sheet, table_name)

    return sheet, pivot_table
```

## 🔴 `configure_pivot(...)`

- Adds filters, rows, columns
- Applies multi-value aggregation
- Refreshes Pivot Table
- Applies sorting rules
- Applies layout settings
- ⚡ **Purpose**: Main configuration engine for Pivot setup


```python
# =====================================================
# CONFIGURE PIVOT
# =====================================================
def configure_pivot(
    pivot_table,
    filters,
    rows,
    columns,
    values,
    filter_values,
    sort_fields=None,
    value_formats=None,
    layout_config=None
):

    add_fields(pivot_table, filters, 3)
    add_fields(pivot_table, rows, 1)
    add_fields(pivot_table, columns, 2)

    apply_filter_values(pivot_table, filters, filter_values)

    #  MULTI AGG EXECUTION
    add_multi_values(pivot_table, values, value_formats)

    pivot_table.RefreshTable()

    # sorting
    for field, order in (sort_fields or {}).items():
        sort_pivot_field(pivot_table, field, order)

    # layout
    if layout_config:
        apply_pivot_layout(pivot_table, **layout_config)

    pivot_table.RowAxisLayout(1)

```

## 🟤 `finalize(wb, sheet)`

- Auto-fits all columns
- Saves workbook
- Closes Excel file
- ⚡ **Purpose**: Final output cleanup and save step


```python
# =====================================================
# FINALIZE
# =====================================================
def finalize(wb, sheet):
    sheet.Columns.AutoFit()
    wb.Save()
    wb.Close(True)
```

## ⚫ `cleanup(excel)`

- Quits Excel application
- Releases COM resources
- Uninitializes pythoncom
- ⚡ **Purpose**: Prevents memory leaks and hanging Excel processes


```python
# =====================================================
# CLEANUP
# =====================================================
def cleanup(excel):
    excel.Quit()
    del excel
    pythoncom.CoUninitialize()
```

## 🟢 `create_pivot_table_com(...)` 

- Initializes COM environment
- Opens Excel file and source sheet
- Cleans headers (removes spaces)
- Builds dynamic data range
- Creates Pivot sheet + Pivot table
- Applies full configuration (filters, rows, columns, values)
- Applies sorting, formatting, and layout
- Saves and closes workbook
- Cleans up Excel process
- ⚡ **Purpose**: End-to-end automation pipeline for Pivot Table creation


```python
# =====================================================
# MAIN FUNCTION
# =====================================================
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
    header_row=1,
    start_col=1,
    sort_fields=None,
    value_formats=None,
    layout_config=None
):

    pythoncom.CoInitialize()
    excel = None

    try:
        excel = init_excel()

        wb = excel.Workbooks.Open(file_path)
        sheet = wb.Worksheets(source_sheet_name)

        last_col = get_last_col(sheet, header_row, start_col)

        for c in range(1, last_col + 1):
            cell = sheet.Cells(header_row, c)
            if cell.Value:
                cell.Value = str(cell.Value).strip()

        source_data = build_source_range(sheet, header_row, start_col)

        pivot_sheet, pivot_table = build_pivot(
            wb,
            source_data,
            pivot_sheet_name,
            pivot_table_name
        )

        configure_pivot(
            pivot_table,
            filters or [],
            rows or [],
            columns or [],
            values or {},
            filter_values or {},
            sort_fields or {},
            value_formats or {},
            layout_config
        )

        finalize(wb, pivot_sheet)

        #print("Pivot table created successfully with multi-agg support.")

    finally:
        if excel:
            cleanup(excel)

```

## EXAMPLE


```python
# =====================================================
# EXAMPLE: CREATE PIVOT TABLE USING COM AUTOMATION
# =====================================================

create_pivot_table_com(
    # 📂 Excel file path (source dataset)
    file_path=r"D:\Advance_Data_Sets\FSD_MOCN\Cell_List\GSM.xlsx",
    # 📄 Source worksheet containing raw KPI data
    source_sheet_name="GSM_KPIs",
    # 📊 Name of the new sheet where Pivot Table will be created
    pivot_sheet_name="Pivot_CSSR",
    # 🧾 Name of the Pivot Table object inside Excel
    pivot_table_name="PT_CSSR",
    # 🔽 Fields used as Pivot Filters (top-level filtering control)
    filters=["Zone", "Date_Mapped"],
    # 🎯 Default filter selections (pre-applied filtering)
    filter_values={
        "Zone": ["Buffer_Tier_0", "Buffer_Tier_1"],  # multiple selection filter
        "Date_Mapped": ["Pre"]                        # single category filter
    },
    # 📌 Row fields (appears vertically in Pivot Table)
    rows=["Time"],
    # 📊 Column fields (appears horizontally in Pivot Table)
    columns=["Date"],
    # 📈 Values section with aggregation logic
    # Supports multiple KPIs with aggregation type per field
    values={
        "CSSR": ["average"],   # KPI 1: Average CSSR
        "DCR": ["average"]     # KPI 2: Average DCR
    },
    # 🔢 Sorting rules for Pivot fields
    sort_fields={
        "Date": "descending",  # sort dates latest to oldest
        "Time": "descending"   # sort time in reverse order
    },

    # 🎨 Number formatting rules for value fields
    # Example formats:
    # "0.000"   → 3 decimal places
    # "0.00%"   → percentage format
    # "#,##0"   → comma-separated numbers
    value_formats={
        "CSSR": "0.000"
    },

    # 🧱 Layout configuration for Pivot appearance
    layout_config={
        "layout": "tabular",        # options: compact / outline / tabular
        "grand_totals": True,       # enable totals row/column
        "repeat_labels": True       # repeat row labels for readability
    },

    # 📅 Header row position in source data
    header_row=2,

    # 📍 Starting column index of dataset (helps skip metadata columns)
    start_col=2
)
```


```python
# # =====================================================
# # EXAMPLE
# # =====================================================
create_pivot_table_com(

    file_path=r"D:\Advance_Data_Sets\FSD_MOCN\Cell_List\GSM_Test.xlsx",
    source_sheet_name="GSM_KPIs",
    pivot_sheet_name="Pivot_CSSR",
    pivot_table_name="PT_CSSR",
    filters=["GBSC"],
    #filter_values={"Zone": ["Buffer_Tier_0", "Buffer_Tier_1"],"Date_Mapped":["Pre"]},
    rows=["Date","Time"],
    columns=["Cell CI"],
    values={"CSSR_Non Blocking": ["average"]},
    #sort_fields={"Date": "descending","Time": "descending"},                                # ascending 
    value_formats={"CSSR_Non Blocking": "0.000"},                                                        #"0.00%", "#,##0", "Users": "0"
    layout_config={"layout": "tabular","grand_totals": True,"repeat_labels": True},         # compact, outline, tabular
    #header_row=2, start_col=2
)
# #"0.00%", "#,##0", "Users": "0"
```


```python

```
