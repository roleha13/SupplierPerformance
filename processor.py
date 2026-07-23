"""
processor.py
Supplier Performance Report Tool
Part 1
---------------------------------
• Imports
• Validation
• Reading Excel files
• Cleaning data
• Merge Purchase Register & Receiving Report
• Delivery Days calculation
"""
from pathlib import Path
from io import BytesIO

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)

from openpyxl.utils import get_column_letter

from openpyxl.chart import (
    BarChart,
    PieChart,
    Reference
)

from openpyxl.chart.label import DataLabelList

from openpyxl.formatting.rule import (
    ColorScaleRule
)

from config import (

    REPORT_COLUMNS,

    REGISTER_REQUIRED_COLUMNS,

    RECEIVING_INPUT_COLUMNS,

    PURCHASE_REGISTER_COLUMNS,

    PURCHASE_RECEIVING_COLUMNS,

    EXCLUDED_SUPPLIERS,

    REPORT_TITLE,

    MASTER_SHEET,

    OUTPUT_FILE,

    HEADER_FILL,

    HEADER_FONT,

    TOTAL_FILL,

    FREEZE_PANES,

    HEADER_ROW_HEIGHT,

    DEFAULT_ROW_HEIGHT

)

# =============================================================================
# READ EXCEL FILE
# =============================================================================

def read_excel_file(file_path: str | Path) -> pd.DataFrame:
    """
    Reads Materials Control Excel exports (.xls, .xlsx, .xlsm)
    from the 'Data' worksheet.
    """

    file_path = Path(file_path)
    suffix = file_path.suffix.lower()

    if suffix == ".xls":

        df = pd.read_excel(
            file_path,
            sheet_name="Data",
            engine="xlrd"
        )

    elif suffix in [".xlsx", ".xlsm"]:

        df = pd.read_excel(
            file_path,
            sheet_name="Data",
            engine="openpyxl"
        )

    else:

        raise ValueError(
            f"Unsupported file type: {suffix}"
        )

    # Clean column names
    df.columns = (
        df.columns
          .astype(str)
          .str.strip()
          .str.replace(r"\s+", " ", regex=True)
    )

    return df

# =============================================================================
# VALIDATION
# =============================================================================

def validate_columns(df: pd.DataFrame, required: set, file_name: str):
    """
    Validate uploaded workbook columns.
    """

    missing = required - set(df.columns)

    if missing:
        raise ValueError(
            f"\n{file_name}\n\n"
            f"Missing columns:\n"
            f"{', '.join(sorted(missing))}"
        )



# =============================================================================
# READ PURCHASE REGISTER
# =============================================================================

def read_purchase_register(file_path: str | Path) -> pd.DataFrame:
    """
    Read Purchase Register workbook.
    """

    df = read_excel_file(file_path)

    validate_columns(
        df,
        PURCHASE_REGISTER_COLUMNS,
        "Purchase Register"
    )

    df = df[REGISTER_REQUIRED_COLUMNS].copy()

    df["Order No."] = (
        df["Order No."]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
    )

    df["Order Date"] = pd.to_datetime(
        df["Order Date"],
        errors="coerce"
    )

    return df


# =============================================================================
# READ RECEIVING REPORT
# =============================================================================

def read_receiving_report(file_path: str | Path) -> pd.DataFrame:
    """
    Read Purchase Receiving Deviation report.
    """

    df = read_excel_file(file_path)

    validate_columns(
        df,
        PURCHASE_RECEIVING_COLUMNS,
        "Purchase Receiving Deviation"
    )

    return df


# =============================================================================
# CLEAN RECEIVING DATA
# =============================================================================

def clean_receiving_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep required columns and convert datatypes.
    """

    df = df[RECEIVING_INPUT_COLUMNS].copy()

    df["Supplier"] = (
        df["Supplier"]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
    )

    df["Order No."] = (
        df["Order No."]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
    )

    df["Delivery Date"] = pd.to_datetime(
        df["Delivery Date"],
        errors="coerce"
    )

    numeric_columns = [

        "Ordered",
        "Booked QTY",
        "Variance QTY",
        "PO Price",
        "Booked Price",
        "Variance Price",
        "Variance Value"

    ]

    for col in numeric_columns:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)

    return df


# =============================================================================
# REMOVE EXCLUDED SUPPLIERS
# =============================================================================

def remove_excluded_suppliers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove suppliers defined in config.py.
    """

    return df[
        ~df["Supplier"].isin(EXCLUDED_SUPPLIERS)
    ].copy()


# =============================================================================
# MERGE ORDER DATES
# =============================================================================

def merge_order_dates(
    receiving_df: pd.DataFrame,
    register_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Merge Order Date from Purchase Register into the
    Purchase Receiving Deviation report.

    Prevents row multiplication by ensuring only one
    Order Date exists per Order No.

    Ignores placeholder Order Numbers such as
    'No PO defined', blank values, N/A, etc.
    """

    # ---------------------------------------------------------
    # Clean Order Numbers
    # ---------------------------------------------------------

    register_df = register_df.copy()

    register_df["Order No."] = (
        register_df["Order No."]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ---------------------------------------------------------
    # Ignore invalid Order Numbers
    # ---------------------------------------------------------

    invalid_orders = {
        "",
        "NO PO DEFINED",
        "N/A",
        "NONE"
    }

    valid_register = register_df[
        ~register_df["Order No."]
        .str.upper()
        .isin(invalid_orders)
    ].copy()

    # ---------------------------------------------------------
    # Check for conflicting Order Dates
    # ---------------------------------------------------------

    conflicting = (
        valid_register
        .groupby("Order No.")["Order Date"]
        .nunique()
    )

    conflicting = conflicting[conflicting > 1]

    if not conflicting.empty:

        conflicting_orders = ", ".join(conflicting.index.astype(str))

        raise ValueError(
            "Data quality issue detected.\n\n"
            "The following Order Numbers have multiple "
            f"Order Dates in the Purchase Register:\n\n"
            f"{conflicting_orders}\n\n"
            "Please verify the Purchase Register export."
        )

    # ---------------------------------------------------------
    # Keep one Order Date per Order Number
    # ---------------------------------------------------------

    order_lookup = (
        valid_register[["Order No.", "Order Date"]]
        .drop_duplicates(subset="Order No.")
    )

    # ---------------------------------------------------------
    # Merge
    # ---------------------------------------------------------

    merged = receiving_df.merge(
        order_lookup,
        how="left",
        on="Order No."
    )

    return merged

# =============================================================================
# CALCULATE DELIVERY DAYS
# =============================================================================

def calculate_delivery_days(df: pd.DataFrame) -> pd.DataFrame:
    """
    Delivery Days = Delivery Date - Order Date
    """

    df["Delivery Days"] = (
        df["Delivery Date"] -
        df["Order Date"]
    ).dt.days

    return df


# =============================================================================
# PREPARE REPORT DATASET
# =============================================================================

def prepare_report_data(
    purchase_register_file: str | Path,
    receiving_report_file: str | Path
) -> pd.DataFrame:
    """
    Complete preprocessing pipeline.
    """

    register = read_purchase_register(
        purchase_register_file
    )

    receiving = read_receiving_report(
        receiving_report_file
    )

    receiving = clean_receiving_data(receiving)

    receiving = remove_excluded_suppliers(receiving)

    merged = merge_order_dates(
        receiving,
        register
    )

    merged = calculate_delivery_days(merged)

    merged = merged[REPORT_COLUMNS]

    merged.sort_values(
        ["Supplier", "Order Date", "Order No."],
        inplace=True
    )

    merged.reset_index(
        drop=True,
        inplace=True
    )

    return merged  

# =============================================================================
# MASTER SUMMARY KPI CALCULATIONS
# =============================================================================

def create_master_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Create supplier performance summary.
    """

    summary = (
        df.groupby("Supplier", as_index=False)
        .agg(
            Orders=("Order No.", "nunique"),
            Ordered_Qty=("Ordered", "sum"),
            Received_Qty=("Booked QTY", "sum"),
            Qty_Variance=("Variance QTY", "sum"),
            Price_Variance=("Variance Value", "sum"),
            Average_Delivery_Days=("Delivery Days", "mean")
        )
    )

    summary["Order Fulfillment Rate %"] = (
        (
            summary["Received_Qty"] /
            summary["Ordered_Qty"]
        )
        .replace([float("inf")], 0)
        .fillna(0)
    )  

    summary["Order Fulfillment Rate %"] = (
        summary["Order Fulfillment Rate %"]
        .round(4)
    )

    summary.rename(
        columns={
            "Ordered_Qty": "Ordered Qty",
            "Received_Qty": "Received Qty",
            "Qty_Variance": "Qty Variance",
            "Price_Variance": "Price Variance",
            "Average_Delivery_Days": "Average Delivery Days",
        },
        inplace=True,
    )

    summary["Average Delivery Days"] = (
        summary["Average Delivery Days"]
        .round(1)
    )

    summary["Order Fulfillment Rate %"] = (
        summary["Order Fulfillment Rate %"]
        .round(4)
    )

    summary.sort_values(
        by=[
            "Order Fulfillment Rate %",
            "Average Delivery Days"
        ],
        ascending=[
            False,
            True
        ],
        inplace=True
    )

    return summary.reset_index(drop=True)


# =============================================================================
# EXECUTIVE SUMMARY
# =============================================================================

def create_executive_summary(df: pd.DataFrame) -> dict:
    """
    Dashboard KPI cards.
    """

    ordered = df["Ordered"].sum()
    received = df["Booked QTY"].sum()

    fill_rate = (
        (received / ordered) 
        if ordered else 0
    )

    return {

        "Total Suppliers":
            df["Supplier"].nunique(),

        "Total Orders":
            df["Order No."].nunique(),

        "Total Ordered Qty":
            ordered,

        "Total Received Qty":
            received,

        "Overall Order Fulfillment Rate %":
            round(fill_rate, 4),

        "Average Delivery Days":
            round(df["Delivery Days"].mean(), 1),

        "Total Price Variance":
            df["Variance Value"].sum(),

        "Total Quantity Variance":
            df["Variance QTY"].sum()

    }


# =============================================================================
# MASTER SUMMARY SHEET
# =============================================================================

def write_master_summary(workbook, summary_df):

    ws = workbook.create_sheet("Master Summary")

    ws.append(summary_df.columns.tolist())

    for row in summary_df.itertuples(index=False):

        ws.append(list(row))

    # ----------------------------------------------------
    # Format Order Fulfillment Rate column as %
    # ----------------------------------------------------

    fulfillment_col = None

    for col in range(1, ws.max_column + 1):

        if ws.cell(1, col).value == "Order Fulfillment Rate %":

            fulfillment_col = col
            break

    if fulfillment_col:

        for row in range(2, ws.max_row + 1):

            ws.cell(row, fulfillment_col).number_format = "0.00%"

    return ws


# =============================================================================
# SUPPLIER KPI PANEL
# =============================================================================

def supplier_kpis(df: pd.DataFrame):

    ordered = df["Ordered"].sum()
    received = df["Booked QTY"].sum()

    fill_rate = (
        (received / ordered)
        if ordered else 0
    )

    return [

        ("Orders", df["Order No."].nunique()),

        ("Ordered Qty", ordered),

        ("Received Qty", received),

        ("Order Fulfillment Rate %", round(fill_rate, 4)),

        ("Quantity Variance", df["Variance QTY"].sum()),

        ("Price Variance", df["Variance Value"].sum()),

        (
            "Average Delivery Days",
            round(df["Delivery Days"].mean(), 1)
        )

    ]


# =============================================================================
# SUPPLIER WORKSHEETS
# =============================================================================

def create_supplier_sheets(workbook, report_df):

    suppliers = sorted(
        report_df["Supplier"].unique()
    )

    for supplier in suppliers:

        sheet = workbook.create_sheet(
            supplier[:31]
        )

        supplier_df = (
            report_df[
                report_df["Supplier"] == supplier
            ]
            .sort_values(
                [
                    "Order Date",
                    "Order No."
                ]
            )
        )

        # -----------------------------
        # Transaction Table
        # -----------------------------

        sheet.append(
            supplier_df.columns.tolist()
        )

        for row in supplier_df.itertuples(index=False):

            sheet.append(list(row))

        # -----------------------------
        # KPI PANEL
        # -----------------------------

        start_row = sheet.max_row + 3

        sheet.cell(
            start_row,
            1,
            "Supplier KPI Summary"
        )

        kpis = supplier_kpis(
            supplier_df
        )

        for i, (kpi, value) in enumerate(
            kpis,
            start=start_row + 1
        ):

            sheet.cell(i, 1).value = kpi
            sheet.cell(i, 2).value = value

            if kpi == "Order Fulfillment Rate %":

                sheet.cell(i, 2).number_format = "0,00%"


# =============================================================================
# BUILD WORKBOOK
# =============================================================================

def build_workbook(report_df):

    wb = Workbook()

    wb.remove(wb.active)

    summary = create_master_summary(
        report_df
    )

    write_master_summary(
        wb,
        summary
    )

    create_supplier_sheets(
        wb,
        report_df
    )

    return wb 

###############################################################################
# EXCEL FORMATTING PART 3
###############################################################################

def format_worksheet(ws):
    """
    Apply professional formatting.
    """

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )

    header_font = Font(
        bold=True,
        color=HEADER_FONT
    )

    thin = Side(style="thin")

    border = Border(

        left=thin,

        right=thin,

        top=thin,

        bottom=thin

    )

    ws.freeze_panes = FREEZE_PANES

    ws.auto_filter.ref = ws.dimensions

    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT

    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(

            horizontal="center",

            vertical="center"

        )

        cell.border = border

    for row in ws.iter_rows(min_row=2):

        ws.row_dimensions[row[0].row].height = DEFAULT_ROW_HEIGHT

        for cell in row:

            cell.border = border

            if "Date" in str(ws.cell(1, cell.column).value):

                cell.number_format = "dd-mmm-yyyy"

            elif "%" in str(ws.cell(1, cell.column).value):

                cell.number_format = '0.00%'

            elif isinstance(cell.value, (int, float)):

                cell.number_format = '#,##0.00'

    for column in ws.columns:

        length = max(

            len(str(c.value))

            if c.value is not None else 0

            for c in column

        )

        ws.column_dimensions[

            get_column_letter(column[0].column)

        ].width = min(length + 3, 40)


###############################################################################
# CONDITIONAL FORMATTING
###############################################################################

def apply_conditional_formatting(ws):

    headers = {

        cell.value: cell.column

        for cell in ws[1]

    }

    if "Delivery Days" in headers:

        col = get_column_letter(

            headers["Delivery Days"]

        )

        ws.conditional_formatting.add(

            f"{col}2:{col}{ws.max_row}",

            ColorScaleRule(

                start_type="min",

                start_color="63BE7B",

                mid_type="percentile",

                mid_value=50,

                mid_color="FFEB84",

                end_type="max",

                end_color="F8696B"

            )

        )


###############################################################################
# CHARTS
###############################################################################

def add_supplier_chart(ws):

    headers = {

        cell.value: cell.column

        for cell in ws[1]

    }

    if not {

        "Ordered",

        "Booked QTY"

    }.issubset(headers):

        return

    ordered = headers["Ordered"]

    received = headers["Booked QTY"]

    chart = BarChart()

    chart.title = "Ordered vs Received"

    data = Reference(

        ws,

        min_col=ordered,

        max_col=received,

        min_row=1,

        max_row=min(

            ws.max_row,

            21

        )

    )

    cats = Reference(

        ws,

        min_col=headers["Article"],

        min_row=2,

        max_row=min(

            ws.max_row,

            21

        )

    )

    chart.add_data(

        data,

        titles_from_data=True

    )

    chart.set_categories(cats)

    chart.height = 7

    chart.width = 12

    chart.dLbls = DataLabelList()

    chart.dLbls.showVal = True

    ws.add_chart(

        chart,

        f"A{ws.max_row+3}"

    )


###############################################################################
# MASTER DASHBOARD
###############################################################################

def add_dashboard(master_ws, report_df):

    stats = create_executive_summary(

        report_df

    )

    row = 2

    master_ws.insert_rows(1, amount=10)

    master_ws["A1"] = REPORT_TITLE

    master_ws["A1"].font = Font(

        bold=True,

        size=16

    )

    for key, value in stats.items():

        master_ws.cell(

            row,

            1,

            key

        )

        master_ws.cell(

            row,

            2,

            value

        )

        row += 1


###############################################################################
# SAVE REPORT
###############################################################################
def save_workbook(workbook):
    """
    Save workbook to memory instead of disk.
    """

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return output


###############################################################################
# MAIN PROCESS
###############################################################################

def process_files(

    purchase_register,

    receiving_report,

):

    # ---------------------------------------------------------
    # Prepare report data
    # ---------------------------------------------------------

    report_df = prepare_report_data(

        purchase_register,

        receiving_report

    )

    # ---------------------------------------------------------
    # Determine report period for output filename
    # ---------------------------------------------------------

    report_period = (

        report_df["Order Date"]

        .dropna()

        .min()

        .strftime("%B_%Y")

    )

    # ---------------------------------------------------------
    # Build workbook
    # ---------------------------------------------------------

    workbook = build_workbook(

        report_df

    )

    # ---------------------------------------------------------
    # Add dashboard
    # ---------------------------------------------------------

    master = workbook[MASTER_SHEET]

    add_dashboard(

        master,

        report_df

    )

    # ---------------------------------------------------------
    # Format worksheets
    # ---------------------------------------------------------

    for sheet in workbook.worksheets:

        format_worksheet(sheet)

        apply_conditional_formatting(sheet)

        if sheet.title != MASTER_SHEET:

            add_supplier_chart(sheet)

    # ---------------------------------------------------------
    # Return workbook and report period
    # ---------------------------------------------------------

    return save_workbook(workbook), report_period
