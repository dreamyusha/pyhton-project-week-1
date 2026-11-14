"""
Week 1 Data Analytics Assignment & Sales Performance Dashboard
Dataset: Global Superstore (global_superstore_2016.xlsx)

Tasks Automated:
1. Data Cleaning: Remove exact duplicates & impute missing values.
2. Pivot Table Summaries: Region-wise sales, Category revenue, Yearly trends, and Department/Segment revenue.
3. Descriptive Statistics: Mean, Median, Mode, Sample/Population Variance using live Excel formulas.
4. Conditional Formatting: Highlight Top 10 Sales transactions.
5. Dashboard Visualizations: Native Excel Bar & Line Charts for Category, Yearly, and Department revenue.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.chart import BarChart, LineChart, Reference


INPUT_FILE = "global_superstore_2016.xlsx"
OUTPUT_FILE = "Global_Superstore_Week1_Dashboard.xlsx"


def load_and_clean_data(filepath: str) -> tuple[pd.DataFrame, dict]:
    """Loads the Orders sheet, removes duplicates, and handles missing values."""
    df = pd.read_excel(filepath, sheet_name="Orders")
    initial_rows = len(df)
    missing_before = df.isnull().sum()
    missing_cols = missing_before[missing_before > 0].to_dict()

    # 1. Remove exact duplicate rows
    df_clean = df.drop_duplicates().copy()
    duplicates_removed = initial_rows - len(df_clean)

    # 2. Handle missing Postal Code (80.51% missing for non-US orders)
    df_clean["Postal Code"] = (
        df_clean["Postal Code"]
        .fillna("Not Available")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
    )

    # 3. Standardize date columns & extract Year for trend analysis
    df_clean["Order Date"] = pd.to_datetime(df_clean["Order Date"])
    df_clean["Ship Date"] = pd.to_datetime(df_clean["Ship Date"])
    df_clean["Year"] = df_clean["Order Date"].dt.year

    audit_log = {
        "initial_rows": initial_rows,
        "final_rows": len(df_clean),
        "duplicates_removed": duplicates_removed,
        "missing_before": missing_cols,
        "missing_after": int(df_clean.isnull().sum().sum()),
    }
    return df_clean, audit_log


def build_pivot_summaries(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Generates Pivot Table aggregations required for the assignment and dashboard."""
    # Region-wise Sales Pivot
    region_pivot = (
        df.groupby("Region", as_index=False)
        .agg(Total_Sales=("Sales", "sum"), Order_Items=("Row ID", "count"), Avg_Sale=("Sales", "mean"))
        .sort_values("Total_Sales", ascending=False)
    )

    # Category-wise Revenue Pivot
    category_pivot = (
        df.groupby("Category", as_index=False)
        .agg(Total_Revenue=("Sales", "sum"), Total_Profit=("Profit", "sum"), Order_Items=("Row ID", "count"))
        .sort_values("Total_Revenue", ascending=False)
    )

    # Yearly Sales Trends Pivot
    yearly_pivot = (
        df.groupby("Year", as_index=False)
        .agg(Total_Sales=("Sales", "sum"), Total_Profit=("Profit", "sum"), Order_Items=("Row ID", "count"))
        .sort_values("Year", ascending=True)
    )
    yearly_pivot["YoY_Growth_%"] = yearly_pivot["Total_Sales"].pct_change().fillna(0) * 100

    # Department-wise Revenue (Customer Segment & Product Sub-Category)
    segment_pivot = (
        df.groupby("Segment", as_index=False)
        .agg(Total_Revenue=("Sales", "sum"), Total_Profit=("Profit", "sum"))
        .sort_values("Total_Revenue", ascending=False)
    )

    subcat_pivot = (
        df.groupby("Sub-Category", as_index=False)
        .agg(Total_Revenue=("Sales", "sum"), Total_Profit=("Profit", "sum"))
        .sort_values("Total_Revenue", ascending=False)
    )

    return {
        "region": region_pivot,
        "category": category_pivot,
        "yearly": yearly_pivot,
        "segment": segment_pivot,
        "subcategory": subcat_pivot,
    }


def export_excel_dashboard(df_clean: pd.DataFrame, pivots: dict[str, pd.DataFrame], output_path: str) -> None:
    """Exports cleaned data, Excel formulas, conditional formatting, and charts to .xlsx."""
    wb = Workbook()

    # Styling definitions
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    kpi_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    bold_font = Font(bold=True, size=11)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # ---------------------------------------------------------
    # TAB 1: Cleaned_Orders (with Top 10 Conditional Formatting)
    # ---------------------------------------------------------
    ws_data = wb.active
    ws_data.title = "Cleaned_Orders"

    export_df = df_clean.drop(columns=["Year"]).copy()
    export_df["Order Date"] = export_df["Order Date"].dt.strftime("%Y-%m-%d")
    export_df["Ship Date"] = export_df["Ship Date"].dt.strftime("%Y-%m-%d")

    for r_idx, row in enumerate(dataframe_to_rows(export_df, index=False, header=True), start=1):
        ws_data.append(row)
        if r_idx == 1:
            for cell in ws_data[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

    last_row = len(export_df) + 1
    # Apply Conditional Formatting to Highlight Top 10 Sales (Column S)
    top10_style = DifferentialStyle(
        fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        font=Font(color="006100", bold=True),
    )
    top10_rule = Rule(type="top10", rank=10, dxf=top10_style)
    ws_data.conditional_formatting.add(f"S2:S{last_row}", top10_rule)

    # ---------------------------------------------------------
    # TAB 2: Descriptive_Stats (Live Excel Formulas)
    # ---------------------------------------------------------
    ws_stats = wb.create_sheet(title="Descriptive_Stats")
    ws_stats.append(["Statistic", "Sales", "Profit", "Quantity", "Discount", "Shipping Cost", "Excel Formula (Sales)"])
    for cell in ws_stats[1]:
        cell.fill = header_fill
        cell.font = header_font

    stats_rows = [
        ("Total (Sum)", f"=SUM(Cleaned_Orders!S2:S{last_row})", f"=SUM(Cleaned_Orders!V2:V{last_row})", f"=SUM(Cleaned_Orders!T2:T{last_row})", "N/A", f"=SUM(Cleaned_Orders!W2:W{last_row})", f"=SUM(S2:S{last_row})"),
        ("Mean (Average)", f"=AVERAGE(Cleaned_Orders!S2:S{last_row})", f"=AVERAGE(Cleaned_Orders!V2:V{last_row})", f"=AVERAGE(Cleaned_Orders!T2:T{last_row})", f"=AVERAGE(Cleaned_Orders!U2:U{last_row})", f"=AVERAGE(Cleaned_Orders!W2:W{last_row})", f"=AVERAGE(S2:S{last_row})"),
        ("Median", f"=MEDIAN(Cleaned_Orders!S2:S{last_row})", f"=MEDIAN(Cleaned_Orders!V2:V{last_row})", f"=MEDIAN(Cleaned_Orders!T2:T{last_row})", f"=MEDIAN(Cleaned_Orders!U2:U{last_row})", f"=MEDIAN(Cleaned_Orders!W2:W{last_row})", f"=MEDIAN(S2:S{last_row})"),
        ("Mode", f"=MODE.SNGL(Cleaned_Orders!S2:S{last_row})", f"=MODE.SNGL(Cleaned_Orders!V2:V{last_row})", f"=MODE.SNGL(Cleaned_Orders!T2:T{last_row})", f"=MODE.SNGL(Cleaned_Orders!U2:U{last_row})", f"=MODE.SNGL(Cleaned_Orders!W2:W{last_row})", f"=MODE.SNGL(S2:S{last_row})"),
        ("Sample Variance", f"=VAR.S(Cleaned_Orders!S2:S{last_row})", f"=VAR.S(Cleaned_Orders!V2:V{last_row})", f"=VAR.S(Cleaned_Orders!T2:T{last_row})", f"=VAR.S(Cleaned_Orders!U2:U{last_row})", f"=VAR.S(Cleaned_Orders!W2:W{last_row})", f"=VAR.S(S2:S{last_row})"),
        ("Population Variance", f"=VAR.P(Cleaned_Orders!S2:S{last_row})", f"=VAR.P(Cleaned_Orders!V2:V{last_row})", f"=VAR.P(Cleaned_Orders!T2:T{last_row})", f"=VAR.P(Cleaned_Orders!U2:U{last_row})", f"=VAR.P(Cleaned_Orders!W2:W{last_row})", f"=VAR.P(S2:S{last_row})"),
        ("Std. Deviation (Sample)", f"=STDEV.S(Cleaned_Orders!S2:S{last_row})", f"=STDEV.S(Cleaned_Orders!V2:V{last_row})", f"=STDEV.S(Cleaned_Orders!T2:T{last_row})", f"=STDEV.S(Cleaned_Orders!U2:U{last_row})", f"=STDEV.S(Cleaned_Orders!W2:W{last_row})", f"=STDEV.S(S2:S{last_row})"),
    ]
    for r in stats_rows:
        ws_stats.append(list(r))

    # ---------------------------------------------------------
    # TAB 3: Sales_Dashboard (KPIs, Pivot Summaries & Charts)
    # ---------------------------------------------------------
    ws_dash = wb.create_sheet(title="Sales_Dashboard")

    # KPI Banner
    ws_dash["A1"] = "GLOBAL SUPERSTORE — SALES PERFORMANCE DASHBOARD"
    ws_dash["A1"].font = Font(bold=True, size=14, color="1F4E78")

    ws_dash["A3"] = "Total Revenue"
    ws_dash["B3"] = f"=SUM(Cleaned_Orders!S2:S{last_row})"
    ws_dash["B3"].number_format = "$#,##0.00"
    ws_dash["D3"] = "Total Profit"
    ws_dash["E3"] = f"=SUM(Cleaned_Orders!V2:V{last_row})"
    ws_dash["E3"].number_format = "$#,##0.00"
    ws_dash["G3"] = "Total Line Items"
    ws_dash["H3"] = f"=COUNTA(Cleaned_Orders!A2:A{last_row})"
    ws_dash["H3"].number_format = "#,##0"

    for coord in ["A3", "B3", "D3", "E3", "G3", "H3"]:
        ws_dash[coord].fill = kpi_fill
        ws_dash[coord].font = bold_font
        ws_dash[coord].border = thin_border

    def write_table(ws, df_table, start_row, start_col, currency_cols):
        for c_idx, col_name in enumerate(df_table.columns, start=start_col):
            cell = ws.cell(row=start_row, column=c_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
        for r_offset, row in enumerate(df_table.itertuples(index=False), start=1):
            for c_offset, val in enumerate(row):
                cell = ws.cell(row=start_row + r_offset, column=start_col + c_offset, value=val)
                cell.border = thin_border
                if df_table.columns[c_offset] in currency_cols:
                    cell.number_format = "$#,##0.00"

    # Write Category Pivot (A5:D8)
    write_table(ws_dash, pivots["category"], 5, 1, ["Total_Revenue", "Total_Profit"])
    # Write Yearly Trends Pivot (A11:E15)
    write_table(ws_dash, pivots["yearly"], 11, 1, ["Total_Sales", "Total_Profit"])
    # Write Department/Segment Pivot (A18:C21)
    write_table(ws_dash, pivots["segment"], 18, 1, ["Total_Revenue", "Total_Profit"])
    # Write Region-Wise Pivot (A24:D47)
    write_table(ws_dash, pivots["region"], 24, 1, ["Total_Sales", "Avg_Sale"])

    # Chart 1: Category-Wise Revenue Bar Chart
    cat_chart = BarChart()
    cat_chart.type = "col"
    cat_chart.title = "Category-Wise Revenue ($)"
    cat_chart.y_axis.title = "Revenue ($)"
    cat_chart.x_axis.title = "Category"
    cat_data = Reference(ws_dash, min_col=2, min_row=5, max_row=8)
    cat_cats = Reference(ws_dash, min_col=1, min_row=6, max_row=8)
    cat_chart.add_data(cat_data, titles_from_data=True)
    cat_chart.set_categories(cat_cats)
    cat_chart.legend = None
    ws_dash.add_chart(cat_chart, "G5")

    # Chart 2: Yearly Sales Trends Line Chart
    yr_chart = LineChart()
    yr_chart.title = "Yearly Sales Trends (2012–2015)"
    yr_chart.y_axis.title = "Total Sales ($)"
    yr_chart.x_axis.title = "Year"
    yr_data = Reference(ws_dash, min_col=2, min_row=11, max_row=15)
    yr_cats = Reference(ws_dash, min_col=1, min_row=12, max_row=15)
    yr_chart.add_data(yr_data, titles_from_data=True)
    yr_chart.set_categories(yr_cats)
    yr_chart.legend = None
    ws_dash.add_chart(yr_chart, "G20")

    # Chart 3: Department (Segment)-Wise Revenue Bar Chart
    seg_chart = BarChart()
    seg_chart.type = "bar"
    seg_chart.title = "Department (Segment)-Wise Revenue ($)"
    seg_chart.x_axis.title = "Revenue ($)"
    seg_data = Reference(ws_dash, min_col=2, min_row=18, max_row=21)
    seg_cats = Reference(ws_dash, min_col=1, min_row=19, max_row=21)
    seg_chart.add_data(seg_data, titles_from_data=True)
    seg_chart.set_categories(seg_cats)
    seg_chart.legend = None
    ws_dash.add_chart(seg_chart, "G35")

    wb.save(output_path)


if __name__ == "__main__":
    cleaned_df, audit = load_and_clean_data(INPUT_FILE)
    pivot_tables = build_pivot_summaries(cleaned_df)
    export_excel_dashboard(cleaned_df, pivot_tables, OUTPUT_FILE)

    print("=== DATA CLEANING AUDIT ===")
    print(audit)
    print("\n=== DESCRIPTIVE STATISTICS (SALES) ===")
    print(f"Total Revenue : ${cleaned_df['Sales'].sum():,.2f}")
    print(f"Mean          : ${cleaned_df['Sales'].mean():,.2f}")
    print(f"Median        : ${cleaned_df['Sales'].median():,.2f}")
    print(f"Mode          : ${cleaned_df['Sales'].mode().iloc[0]:,.2f}")
    print(f"Variance (S)  : {cleaned_df['Sales'].var(ddof=1):,.2f}")
    print(f"\nExcel Dashboard saved successfully to: {OUTPUT_FILE}")