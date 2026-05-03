# Reading and Analyzing Excel Chart Template (.CRTX) Using Python (win32com)

## Import Required Libraries for Excel Automation


```python
import os
import win32com.client as win32
```

## Define Path of Chart Template (.CRTX File)


```python
CRTX_PATH = r"C:\Users\AppData\Roaming\Microsoft\Templates\Charts\CS_Traffic_HQ.crtx"
```

## Function to Read and Extract Chart Template Properties Using Excel Engine


```python
def read_crtx_with_excel(crtx_path):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Add()

    # Apply chart template to a dummy chart
    sheet = wb.Sheets(1)

    # Create a dummy chart object
    chart_obj = sheet.Shapes.AddChart2(240, 51).Chart

    # Apply your template
    chart_obj.ApplyChartTemplate(crtx_path)

    result = {}

    # ------------------------
    # Title
    # ------------------------
    try:
        result["title"] = chart_obj.ChartTitle.Text
    except:
        result["title"] = None

    # ------------------------
    # Chart Type
    # ------------------------
    try:
        chart_type_map = {
            51: "column",
            4: "line",
            5: "pie",
            58: "bar",
            56: "area"
        }
        result["chart_type"] = chart_type_map.get(chart_obj.ChartType, "unknown")
    except:
        result["chart_type"] = "unknown"

    # ------------------------
    # Series
    # ------------------------
    series_list = []
    try:
        for s in chart_obj.SeriesCollection():
            series_list.append(s.Name)
    except:
        pass

    result["series"] = series_list

    # ------------------------
    # Secondary Axis Detection (REAL)
    # ------------------------
    has_secondary = False
    try:
        for s in chart_obj.SeriesCollection():
            if s.AxisGroup == 2:   # xlSecondary = 2
                has_secondary = True
                break
    except:
        pass

    result["has_secondary_axis"] = has_secondary

    wb.Close(False)
    excel.Quit()

    return result


if __name__ == "__main__":
    output = read_crtx_with_excel(CRTX_PATH)

    print("\n==============================")
    print("REAL CRTX ANALYSIS (EXCEL ENGINE)")
    print("==============================")
    print(output)
```

    
    ==============================
    REAL CRTX ANALYSIS (EXCEL ENGINE)
    ==============================
    {'title': None, 'chart_type': 'unknown', 'series': ['Series1', 'Series2', 'Series3', 'Series4'], 'has_secondary_axis': True}
    
