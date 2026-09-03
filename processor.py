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
from openpyxl.worksheet.hyperlink import Hyperlink

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

    Business Rule
    -------------
    A Purchase Order may contain multiple line items and,
    therefore, may appear with more than one Order Date.

    When a Purchase Order has multiple Order Dates, the
    EARLIEST Order Date is used as the official PO Order Date.

    This ensures Delivery Days is calculated from the first
    date associated with the Purchase Order rather than from
    an arbitrary row.

    Invalid Order Numbers such as blank values, 'No PO defined',
    'N/A', and 'NONE' are ignored.
    """

    # ---------------------------------------------------------
    # Clean Register
    # ---------------------------------------------------------

    register_df = register_df.copy()

    register_df["Order No."] = (
        register_df["Order No."]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ---------------------------------------------------------
    # Clean Order Date
    # ---------------------------------------------------------

    register_df["Order Date"] = pd.to_datetime(
        register_df["Order Date"],
        errors="coerce"
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
        ~register_df["Order No."].isin(invalid_orders)
    ].copy()

    # ---------------------------------------------------------
    # Remove rows where Order Date is missing
    # ---------------------------------------------------------

    valid_register = valid_register[
        valid_register["Order Date"].notna()
    ].copy()

    # ---------------------------------------------------------
    # Identify POs with multiple Order Dates
    # ---------------------------------------------------------

    date_counts = (
        valid_register
        .groupby("Order No.")["Order Date"]
        .nunique()
    )

    conflicting_orders = date_counts[
        date_counts > 1
    ].index

    # ---------------------------------------------------------
    # Handle multiple Order Dates
    # ---------------------------------------------------------

    if len(conflicting_orders) > 0:

        print(
            "\n=================================================="
        )
        print(
            "MULTIPLE ORDER DATES DETECTED"
        )
        print(
            "=================================================="
        )

        for order_no in conflicting_orders:

            dates = (
                valid_register.loc[
                    valid_register["Order No."] == order_no,
                    "Order Date"
                ]
                .dropna()
                .sort_values()
                .dt.strftime("%d-%b-%Y")
                .unique()
            )

            selected_date = dates[0]

            print(
                f"PO: {order_no}"
            )

            print(
                f"Available Order Dates: "
                f"{', '.join(dates)}"
            )

            print(
                f"Selected Order Date: "
                f"{selected_date}"
            )

            print(
                "Rule Applied: Earliest Order Date"
            )

            print(
                "--------------------------------------------------"
            )

        print(
            "==================================================\n"
        )

    # ---------------------------------------------------------
    # Create one Order Date per Purchase Order
    #
    # IMPORTANT:
    # The earliest date is deliberately selected.
    # ---------------------------------------------------------

    order_lookup = (
        valid_register
        .groupby("Order No.", as_index=False)
        .agg(
            **{
                "Order Date": (
                    "Order Date",
                    "min"
                )
            }
        )
    )

    # ---------------------------------------------------------
    # Merge Order Date into Receiving Report
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
    Average Delivery Days is calculated per unique
    Purchase Order instead of per article line.
    """

    # -------------------------------------------------------------------------
    # Delivery Days (One record per Purchase Order)
    # -------------------------------------------------------------------------

    delivery_summary = (
        df[
            ["Supplier", "Order No.", "Delivery Days"]
        ]
        .drop_duplicates(
            subset=["Supplier", "Order No."]
        )
        .groupby(
            "Supplier",
            as_index=False
        )
        .agg(
            Average_Delivery_Days=("Delivery Days", "mean")
        )
    )

    # -------------------------------------------------------------------------
    # Main Supplier KPIs
    # -------------------------------------------------------------------------

    summary = (
        df.groupby("Supplier", as_index=False)
        .agg(
            Orders=("Order No.", "nunique"),
            Ordered_Qty=("Ordered", "sum"),
            Received_Qty=("Booked QTY", "sum"),
            Qty_Variance=("Variance QTY", "sum"),
            Price_Variance=("Variance Value", "sum")
        )
    )

    # -------------------------------------------------------------------------
    # Merge Average Delivery Days
    # -------------------------------------------------------------------------

    summary = summary.merge(
        delivery_summary,
        on="Supplier",
        how="left"
    )

    # -------------------------------------------------------------------------
    # Order Fulfillment Rate
    # -------------------------------------------------------------------------

    summary["Order Fulfillment Rate %"] = (
        (
            summary["Received_Qty"] /
            summary["Ordered_Qty"]
        )
        .replace([float("inf")], 0)
        .fillna(0)
        .round(4)
    )

    # -------------------------------------------------------------------------
    # Rename Columns
    # -------------------------------------------------------------------------

    summary.rename(
        columns={
            "Ordered_Qty": "Ordered Qty",
            "Received_Qty": "Received Qty",
            "Qty_Variance": "Qty Variance",
            "Price_Variance": "Price Variance",
            "Average_Delivery_Days": "Average Delivery Days"
        },
        inplace=True
    )

    # -------------------------------------------------------------------------
    # Round Values
    # -------------------------------------------------------------------------

    summary["Average Delivery Days"] = (
        summary["Average Delivery Days"]
        .round(1)
    )

    # -------------------------------------------------------------------------
    # Sort Suppliers
    # -------------------------------------------------------------------------

    summary.sort_values(
        by=[
            "Order Fulfillment Rate %",
            "Average Delivery Days",
            "Orders"
        ],
        ascending=[
            False,  # highest fulfillment first
            True,   # lowest delivery days first
            False   # most orders first
        ],
        inplace=True
    )

    summary.reset_index(
        drop=True,
        inplace=True
    )

    return summary
    
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

def write_master_summary(workbook, summary_df, supplier_sheet_map):

    ws = workbook.create_sheet("Master Summary")

    # -----------------------------------------------------
    # Reserve rows 1-10 for Dashboard
    # -----------------------------------------------------

    START_ROW = 11

    # -----------------------------------------------------
    # Styles
    # -----------------------------------------------------

    thin = Side(style="thin")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    # -----------------------------------------------------
    # Write Header Row
    # -----------------------------------------------------

    for col, header in enumerate(
        summary_df.columns,
        start=1
    ):

        cell = ws.cell(
            START_ROW,
            col
        )

        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # -----------------------------------------------------
    # Give Header Row Enough Vertical Space
    # -----------------------------------------------------

    ws.row_dimensions[START_ROW].height = 35

    # -----------------------------------------------------
    # Write Supplier Names / Base Rows
    # -----------------------------------------------------

    current_row = START_ROW + 1

    for record in summary_df.itertuples(
        index=False
    ):

        # -------------------------------------------------
        # Write initial values
        #
        # These values will be replaced with formulas
        # for the KPI columns below.
        # -------------------------------------------------

        for col, value in enumerate(
            record,
            start=1
        ):

            cell = ws.cell(
                current_row,
                col
            )

            cell.value = value
            cell.border = border

            if isinstance(
                value,
                (int, float)
            ):

                cell.number_format = "#,##0.00"

        # -----------------------------------------------------
        # Supplier Hyperlink
        # -----------------------------------------------------

        supplier_cell = ws.cell(
            current_row,
            1
        )

        supplier_name = str(
            supplier_cell.value
        )

        sheet_name = supplier_sheet_map.get(
            supplier_name,
            supplier_name[:31]
        )

        # -----------------------------------------------------
        # IMPORTANT:
        # Keep the existing working hyperlink approach.
        # Do NOT use insert_rows() anywhere here.
        # -----------------------------------------------------

        supplier_cell.hyperlink = (
            f"#'{sheet_name}'!A1"
        )

        # -----------------------------------------------------
        # Hyperlink appearance
        # -----------------------------------------------------

        supplier_cell.font = Font(
            color="0563C1",
            underline="single"
        )

        supplier_cell.border = border

        # -----------------------------------------------------
        # Move to next supplier
        # -----------------------------------------------------

        current_row += 1

    # =========================================================
    # LIVE FORMULAS
    # =========================================================

    # ---------------------------------------------------------
    # Identify Master Summary Columns
    # ---------------------------------------------------------

    headers = {
        cell.value: cell.column
        for cell in ws[START_ROW]
    }

    supplier_col = headers.get(
        "Supplier"
    )

    orders_col = headers.get(
        "Orders"
    )

    ordered_col = headers.get(
        "Ordered Qty"
    )

    received_col = headers.get(
        "Received Qty"
    )

    qty_variance_col = headers.get(
        "Qty Variance"
    )

    price_variance_col = headers.get(
        "Price Variance"
    )

    avg_days_col = headers.get(
        "Average Delivery Days"
    )

    fulfillment_col = headers.get(
        "Order Fulfillment Rate %"
    )

    # ---------------------------------------------------------
    # Loop through each supplier row
    # ---------------------------------------------------------

    for row in range(
        START_ROW + 1,
        ws.max_row + 1
    ):

        supplier_name = str(
            ws.cell(
                row,
                supplier_col
            ).value
        )

        sheet_name = supplier_sheet_map.get(
            supplier_name
        )

        if not sheet_name:
            continue

        supplier_sheet = workbook[sheet_name]

        # -----------------------------------------------------
        # Find Supplier KPI Panel
        # -----------------------------------------------------

        kpi_title_row = None

        for search_row in range(
            1,
            supplier_sheet.max_row + 1
        ):

            if (
                supplier_sheet.cell(
                    search_row,
                    1
                ).value
                == "Supplier KPI Summary"
            ):

                kpi_title_row = search_row
                break

        # -----------------------------------------------------
        # Safety Check
        # -----------------------------------------------------

        if kpi_title_row is None:
            continue

        # -----------------------------------------------------
        # Find KPI Rows by Label
        # -----------------------------------------------------

        kpi_rows = {}

        for search_row in range(
            kpi_title_row + 1,
            min(
                kpi_title_row + 8,
                supplier_sheet.max_row + 1
            )
        ):

            kpi_name = supplier_sheet.cell(
                search_row,
                1
            ).value

            if kpi_name:
                kpi_rows[str(kpi_name)] = search_row

        # =====================================================
        # ORDERS
        # =====================================================

        if (
            orders_col
            and "Orders" in kpi_rows
        ):

            source_row = kpi_rows["Orders"]

            ws.cell(
                row,
                orders_col
            ).value = (
                f"='{sheet_name}'!B{source_row}"
            )

            ws.cell(
                row,
                orders_col
            ).number_format = "0"

        # =====================================================
        # ORDERED QTY
        # =====================================================

        if (
            ordered_col
            and "Ordered Qty" in kpi_rows
        ):

            source_row = kpi_rows["Ordered Qty"]

            ws.cell(
                row,
                ordered_col
            ).value = (
                f"='{sheet_name}'!B{source_row}"
            )

            ws.cell(
                row,
                ordered_col
            ).number_format = "#,##0.00"

        # =====================================================
        # RECEIVED QTY
        # =====================================================

        if (
            received_col
            and "Received Qty" in kpi_rows
        ):

            source_row = kpi_rows["Received Qty"]

            ws.cell(
                row,
                received_col
            ).value = (
                f"='{sheet_name}'!B{source_row}"
            )

            ws.cell(
                row,
                received_col
            ).number_format = "#,##0.00"

        # =====================================================
        # QUANTITY VARIANCE
        # =====================================================

        if (
            qty_variance_col
            and "Quantity Variance" in kpi_rows
        ):

            source_row = kpi_rows[
                "Quantity Variance"
            ]

            ws.cell(
                row,
                qty_variance_col
            ).value = (
                f"='{sheet_name}'!B{source_row}"
            )

            ws.cell(
                row,
                qty_variance_col
            ).number_format = "#,##0.00"

        # =====================================================
        # PRICE VARIANCE
        # =====================================================

        if (
            price_variance_col
            and "Price Variance" in kpi_rows
        ):

            source_row = kpi_rows[
                "Price Variance"
            ]

            ws.cell(
                row,
                price_variance_col
            ).value = (
                f"='{sheet_name}'!B{source_row}"
            )

            ws.cell(
                row,
                price_variance_col
            ).number_format = "#,##0.00"

        # =====================================================
        # AVERAGE DELIVERY DAYS
        # =====================================================

        if (
            avg_days_col
            and "Average Delivery Days" in kpi_rows
        ):

            source_row = kpi_rows[
                "Average Delivery Days"
            ]

            ws.cell(
                row,
                avg_days_col
            ).value = (
                f"='{sheet_name}'!B{source_row}"
            )

            ws.cell(
                row,
                avg_days_col
            ).number_format = "0.0"

        # =====================================================
        # ORDER FULFILLMENT RATE
        # =====================================================

        if (
            fulfillment_col
            and ordered_col
            and received_col
        ):

            ordered_letter = get_column_letter(
                ordered_col
            )

            received_letter = get_column_letter(
                received_col
            )

            ws.cell(
                row,
                fulfillment_col
            ).value = (
                f"=IF("
                f"{ordered_letter}{row}=0,"
                f"0,"
                f"{received_letter}{row}/"
                f"{ordered_letter}{row}"
                f")"
            )

            ws.cell(
                row,
                fulfillment_col
            ).number_format = "0.00%"

    # =========================================================
    # GENERAL FORMATTING
    # =========================================================

    # ---------------------------------------------------------
    # Make sure all data cells retain borders
    # ---------------------------------------------------------

    for row in range(
        START_ROW + 1,
        ws.max_row + 1
    ):

        for col in range(
            1,
            ws.max_column + 1
        ):

            ws.cell(
                row,
                col
            ).border = border

    # ---------------------------------------------------------
    # Percentage Column
    # ---------------------------------------------------------

    if fulfillment_col:

        for row in range(
            START_ROW + 1,
            ws.max_row + 1
        ):

            ws.cell(
                row,
                fulfillment_col
            ).number_format = "0.00%"

    # ---------------------------------------------------------
    # Delivery Days
    # ---------------------------------------------------------

    if avg_days_col:

        for row in range(
            START_ROW + 1,
            ws.max_row + 1
        ):

            ws.cell(
                row,
                avg_days_col
            ).number_format = "0.0"

    # =========================================================
    # AUTO FILTER
    # =========================================================

    last_col = get_column_letter(
        ws.max_column
    )

    ws.auto_filter.ref = (
        f"A{START_ROW}:"
        f"{last_col}{ws.max_row}"
    )

    # =========================================================
    # FREEZE PANES
    # =========================================================

    ws.freeze_panes = (
        f"A{START_ROW + 1}"
    )

    # =========================================================
    # MASTER SUMMARY COLUMN WIDTHS
    # =========================================================

    format_master_summary_columns(ws)

    return ws
# =============================================================================
# MASTER SUMMARY COLUMN WIDTHS
# =============================================================================

def format_master_summary_columns(ws):
    """
    Set professional column widths for the Master Summary sheet.

    These widths accommodate both:
    - The Dashboard in rows 1-10
    - The Master Summary table beginning at row 11
    """

    column_widths = {

        "A": 40,   # Supplier / Dashboard labels
        "B": 18,   # Orders / Dashboard values
        "C": 18,   # Ordered Qty
        "D": 18,   # Received Qty
        "E": 18,   # Qty Variance
        "F": 20,   # Price Variance
        "G": 25,   # Average Delivery Days
        "H": 27    # Order Fulfillment Rate %
    }

    # ---------------------------------------------------------
    # Apply column widths
    # ---------------------------------------------------------

    for column, width in column_widths.items():

        ws.column_dimensions[column].width = width
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

    delivery_days = (
        df[
            ["Order No.", "Delivery Days"]
        ]
        .drop_duplicates(subset=["Order No."])
        ["Delivery Days"]
        .mean()
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
            round(delivery_days, 1)
        )

    ]


###############################################################################
# ARTICLE SUMMARY (PIVOT STYLE)
###############################################################################

def create_article_summary(sheet, supplier_df, start_row):
    """
    Creates a professional pivot-style Monthly Article Summary
    with expandable Order Number details.

    Returns
    -------
    summary_row : int
        First row of the article summary table.

    summary_rows : list[int]
        Worksheet row numbers containing ONLY the article summary rows.
        Used for chart creation.

    article_summary : pandas.DataFrame
        Aggregated article-level summary data.
    """

    # -------------------------------------------------------------------------
    # PROFESSIONAL TABLE STYLES
    # -------------------------------------------------------------------------

    from openpyxl.styles import (
        Font,
        PatternFill,
        Border,
        Side,
        Alignment
    )

    # Use the same colours already used elsewhere in your workbook
    title_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAD3"
    )

    detail_fill = PatternFill(
        fill_type="solid",
        fgColor="F7F7F7"
    )

    white_font = Font(
        color=HEADER_FONT,
        bold=True,
        size=12
    )

    header_font = Font(
        bold=True,
        size=10
    )

    article_font = Font(
        bold=False,
        size=10
    )

    detail_font = Font(
        italic=True,
        size=9
    )

    thin_side = Side(
        style="thin",
        color="B7B7B7"
    )

    table_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    # -------------------------------------------------------------------------
    # TITLE
    # -------------------------------------------------------------------------

    title_row = start_row

    # Merge title across the five article-summary columns
    sheet.merge_cells(
        start_row=title_row,
        start_column=1,
        end_row=title_row,
        end_column=5
    )

    title_cell = sheet.cell(title_row, 1)
    title_cell.value = "Monthly Article Summary"

    title_cell.fill = title_fill
    title_cell.font = white_font
    title_cell.alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    sheet.row_dimensions[title_row].height = 24

    # Apply title fill/borders across the full merged area
    for col in range(1, 6):

        cell = sheet.cell(title_row, col)

        cell.fill = title_fill
        cell.border = table_border

    # -------------------------------------------------------------------------
    # SPACE BETWEEN TITLE AND TABLE
    # -------------------------------------------------------------------------

    start_row += 2

    # -------------------------------------------------------------------------
    # HEADERS
    # -------------------------------------------------------------------------

    headers = [
        "Article",
        "Ordered Qty",
        "Delivered Qty",
        "Qty Variance",
        "No. of Orders"
    ]

    header_row = start_row

    for col, header in enumerate(headers, start=1):

        cell = sheet.cell(header_row, col)

        cell.value = header

        cell.fill = header_fill
        cell.font = header_font
        cell.border = table_border

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    sheet.row_dimensions[header_row].height = 21

    # -------------------------------------------------------------------------
    # FIRST ARTICLE DATA ROW
    # -------------------------------------------------------------------------

    summary_row = header_row + 1

    # -------------------------------------------------------------------------
    # KEEP TRACK OF ONLY ARTICLE ROWS
    # -------------------------------------------------------------------------

    summary_rows = []

    # -------------------------------------------------------------------------
    # MONTHLY ARTICLE TOTALS
    # -------------------------------------------------------------------------

    article_summary = (
        supplier_df
        .groupby("Article", as_index=False)
        .agg(
            Ordered=("Ordered", "sum"),
            Delivered=("Booked QTY", "sum"),
            Variance=("Variance QTY", "sum"),
            Order_Frequency=("Order No.", "nunique")
        )
        .sort_values("Article")
    )

    current_row = summary_row

    # -------------------------------------------------------------------------
    # WRITE EACH ARTICLE
    # -------------------------------------------------------------------------

    for _, article in article_summary.iterrows():

        article_row = current_row

        # ---------------------------------------------------------------------
        # ARTICLE SUMMARY ROW
        # ---------------------------------------------------------------------

        article_cell = sheet.cell(current_row, 1)
        article_cell.value = article["Article"]

        ordered_cell = sheet.cell(current_row, 2)
        ordered_cell.value = article["Ordered"]

        delivered_cell = sheet.cell(current_row, 3)
        delivered_cell.value = article["Delivered"]

        variance_cell = sheet.cell(current_row, 4)
        variance_cell.value = article["Variance"]

        frequency_cell = sheet.cell(current_row, 5)
        frequency_cell.value = article["Order_Frequency"]

        # ---------------------------------------------------------------------
        # ARTICLE ROW FORMATTING
        # ---------------------------------------------------------------------

        for col in range(1, 6):

            cell = sheet.cell(current_row, col)

            cell.border = table_border
            cell.font = article_font

            cell.alignment = Alignment(
                vertical="center"
            )

        # Article name left aligned
        article_cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        # Numbers right aligned
        for col in range(2, 6):

            sheet.cell(current_row, col).alignment = Alignment(
                horizontal="right",
                vertical="center"
            )

        # Number formatting
        ordered_cell.number_format = "#,##0.00"
        delivered_cell.number_format = "#,##0.00"
        variance_cell.number_format = "#,##0.00"
        frequency_cell.number_format = "0"

        sheet.row_dimensions[current_row].height = 20

        # Save ONLY article rows for charting
        summary_rows.append(current_row)

        current_row += 1

        # ---------------------------------------------------------------------
        # ORDER DETAIL ROWS
        # ---------------------------------------------------------------------

        detail = (
            supplier_df[
                supplier_df["Article"] == article["Article"]
            ]
            .sort_values(
                [
                    "Order Date",
                    "Order No."
                ]
            )
        )

        for _, order in detail.iterrows():

            # -------------------------------------------------------------
            # Order Number
            # -------------------------------------------------------------

            detail_article = sheet.cell(current_row, 1)

            detail_article.value = (
                "    " + str(order["Order No."])
            )

            # -------------------------------------------------------------
            # Quantities
            # -------------------------------------------------------------

            detail_ordered = sheet.cell(current_row, 2)
            detail_ordered.value = order["Ordered"]

            detail_delivered = sheet.cell(current_row, 3)
            detail_delivered.value = order["Booked QTY"]

            detail_variance = sheet.cell(current_row, 4)
            detail_variance.value = order["Variance QTY"]

            # No. of Orders column intentionally left blank
            sheet.cell(current_row, 5).value = None

            # -------------------------------------------------------------
            # Detail Row Formatting
            # -------------------------------------------------------------

            for col in range(1, 6):

                cell = sheet.cell(current_row, col)

                cell.fill = detail_fill
                cell.border = table_border
                cell.font = detail_font

                cell.alignment = Alignment(
                    vertical="center"
                )

            # Order number left aligned
            detail_article.alignment = Alignment(
                horizontal="left",
                vertical="center",
                indent=1
            )

            # Numeric values right aligned
            for col in range(2, 5):

                sheet.cell(current_row, col).alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            # Number formats
            detail_ordered.number_format = "#,##0.00"
            detail_delivered.number_format = "#,##0.00"
            detail_variance.number_format = "#,##0.00"

            # -------------------------------------------------------------
            # Make Order Rows Collapsible
            # -------------------------------------------------------------

            sheet.row_dimensions[current_row].outlineLevel = 1
            sheet.row_dimensions[current_row].hidden = True
            sheet.row_dimensions[current_row].height = 18

            current_row += 1

        # ---------------------------------------------------------------------
        # COLLAPSE DETAILS UNDER ARTICLE
        # ---------------------------------------------------------------------

        sheet.row_dimensions[article_row].collapsed = True

    # -------------------------------------------------------------------------
    # COLUMN WIDTHS
    # -------------------------------------------------------------------------

    # Set minimum professional widths.
    # format_worksheet() may later auto-fit them if your existing code does so.

    sheet.column_dimensions["A"].width = max(
        sheet.column_dimensions["A"].width or 0,
        28
    )

    sheet.column_dimensions["B"].width = max(
        sheet.column_dimensions["B"].width or 0,
        15
    )

    sheet.column_dimensions["C"].width = max(
        sheet.column_dimensions["C"].width or 0,
        16
    )

    sheet.column_dimensions["D"].width = max(
        sheet.column_dimensions["D"].width or 0,
        15
    )

    sheet.column_dimensions["E"].width = max(
        sheet.column_dimensions["E"].width or 0,
        15
    )

    # -------------------------------------------------------------------------
    # RETURN ONLY ARTICLE ROWS
    # -------------------------------------------------------------------------

    return (
        summary_row,
        summary_rows,
        article_summary
    )
###############################################################################
# HELPER TABLE
###############################################################################

def create_helper_table(sheet, supplier_df, start_row):
    """
    Creates a hidden helper table containing one row per valid Purchase Order.

    The helper table is used for:
        - Orders count
        - Average Delivery Days

    AA = Order No.
    AB = Delivery Days

    One row is created for each unique Purchase Order.
    """

    # -------------------------------------------------------------------------
    # Prepare helper data
    # -------------------------------------------------------------------------

    helper = (
        supplier_df[
            supplier_df["Order No."].notna()
            & ~supplier_df["Order No."].astype(str).str.strip().isin(
                ["", "NO PO DEFINED", "N/A", "NONE"]
            )
        ]
        .groupby("Order No.", as_index=False)
        .agg(
            Order_Date=("Order Date", "first"),
            Last_Delivery_Date=("Delivery Date", "max")
        )
    )

    # -------------------------------------------------------------------------
    # Calculate Delivery Days
    # -------------------------------------------------------------------------

    helper["Delivery Days"] = (
        helper["Last_Delivery_Date"]
        - helper["Order_Date"]
    ).dt.days

    # -------------------------------------------------------------------------
    # Write Helper Table Headers
    # -------------------------------------------------------------------------

    sheet.cell(start_row, 27).value = "Order No."
    sheet.cell(start_row, 28).value = "Delivery Days"

    # -------------------------------------------------------------------------
    # Write Helper Table Data
    # -------------------------------------------------------------------------

    row = start_row + 1

    for _, order in helper.iterrows():

        sheet.cell(row, 27).value = order["Order No."]
        sheet.cell(row, 28).value = order["Delivery Days"]

        row += 1

    # -------------------------------------------------------------------------
    # Hide Helper Columns
    # -------------------------------------------------------------------------

    # AA = column 27
    # AB = column 28

    sheet.column_dimensions["AA"].hidden = True
    sheet.column_dimensions["AB"].hidden = True

    # -------------------------------------------------------------------------
    # Return Last Helper Row
    # -------------------------------------------------------------------------

    return row - 1

###############################################################################
# SUPPLIER WORKSHEETS
###############################################################################

def create_supplier_sheets(workbook, report_df, worksheet_last_rows):

    suppliers = sorted(
        report_df["Supplier"].unique()
    )

    supplier_sheet_map = {}

    for supplier in suppliers:

        # =========================================================
        # CREATE A UNIQUE WORKSHEET NAME
        # =========================================================

        sheet_name = supplier[:31]

        count = 1

        while sheet_name in workbook.sheetnames:

            suffix = f"_{count}"

            sheet_name = (
                supplier[:31 - len(suffix)]
                + suffix
            )

            count += 1

        sheet = workbook.create_sheet(sheet_name)

        # Store the ACTUAL worksheet name
        supplier_sheet_map[supplier] = sheet_name

        # =========================================================
        # SUPPLIER DATA
        # =========================================================

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

        # =========================================================
        # TRANSACTION TABLE
        # =========================================================

        sheet.append(
            supplier_df.columns.tolist()
        )

        for row in supplier_df.itertuples(index=False):

            sheet.append(
                list(row)
            )

        # =========================================================
        # LAST TRANSACTION ROW
        # =========================================================

        last_data_row = sheet.max_row

        # IMPORTANT:
        # Keep worksheet_last_rows pointing ONLY to the transaction
        # table. This ensures format_worksheet() does not format
        # the TOTAL row, helper table, KPI panel, article summary,
        # or charts as transaction data.

        worksheet_last_rows[sheet.title] = last_data_row

        # =========================================================
        # TOTAL ROW
        # =========================================================

        total_row = last_data_row + 1

        sheet.cell(
            total_row,
            1
        ).value = "TOTAL"

        # =========================================================
        # LOCATE IMPORTANT TRANSACTION COLUMNS DYNAMICALLY
        # =========================================================

        transaction_headers = {
            str(sheet.cell(1, col).value).strip(): col
            for col in range(
                1,
                sheet.max_column + 1
            )
        }

        ordered_col = transaction_headers.get(
            "Ordered"
        )

        received_col = transaction_headers.get(
            "Booked QTY"
        )

        qty_variance_col = transaction_headers.get(
            "Variance QTY"
        )

        price_variance_col = transaction_headers.get(
            "Variance Value"
        )

        # =========================================================
        # TOTAL - ORDERED QTY
        # =========================================================

        if ordered_col:

            ordered_letter = get_column_letter(
                ordered_col
            )

            sheet.cell(
                total_row,
                ordered_col
            ).value = (
                f"=SUM("
                f"{ordered_letter}2:"
                f"{ordered_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL - RECEIVED QTY
        # =========================================================

        if received_col:

            received_letter = get_column_letter(
                received_col
            )

            sheet.cell(
                total_row,
                received_col
            ).value = (
                f"=SUM("
                f"{received_letter}2:"
                f"{received_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL - QUANTITY VARIANCE
        # =========================================================

        if qty_variance_col:

            qty_variance_letter = get_column_letter(
                qty_variance_col
            )

            sheet.cell(
                total_row,
                qty_variance_col
            ).value = (
                f"=SUM("
                f"{qty_variance_letter}2:"
                f"{qty_variance_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL - PRICE VARIANCE
        # =========================================================

        if price_variance_col:

            price_variance_letter = get_column_letter(
                price_variance_col
            )

            sheet.cell(
                total_row,
                price_variance_col
            ).value = (
                f"=SUM("
                f"{price_variance_letter}2:"
                f"{price_variance_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL ROW FORMATTING
        # =========================================================

        thin = Side(style="thin")

        total_border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        total_fill = PatternFill(
            fill_type="solid",
            fgColor=TOTAL_FILL
        )

        total_font = Font(
            bold=True
        )

        for col in range(
            1,
            sheet.max_column + 1
        ):

            cell = sheet.cell(
                total_row,
                col
            )

            cell.fill = total_fill
            cell.font = total_font
            cell.border = total_border

        # =========================================================
        # NUMBER FORMATTING FOR TOTAL VALUES
        # =========================================================

        if ordered_col:

            sheet.cell(
                total_row,
                ordered_col
            ).number_format = "#,##0.00"

        if received_col:

            sheet.cell(
                total_row,
                received_col
            ).number_format = "#,##0.00"

        if qty_variance_col:

            sheet.cell(
                total_row,
                qty_variance_col
            ).number_format = "#,##0.00"

        if price_variance_col:

            sheet.cell(
                total_row,
                price_variance_col
            ).number_format = "#,##0.00"

        # =========================================================
        # HELPER TABLE
        # =========================================================

        # Layout:
        #
        # Data
        # TOTAL
        # blank row
        # Helper Table

        helper_start = total_row + 2

        helper_end = create_helper_table(
            sheet,
            supplier_df,
            helper_start
        )

        # =========================================================
        # KPI PANEL
        # =========================================================

        start_row = sheet.max_row + 3

        # ---------------------------------------------------------
        # KPI PANEL STYLES
        # ---------------------------------------------------------

        kpi_border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        kpi_title_fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_FILL
        )

        kpi_title_font = Font(
            bold=True,
            color=HEADER_FONT,
            size=12
        )

        kpi_header_fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAD3"
        )

        kpi_header_font = Font(
            bold=True
        )

        # =========================================================
        # KPI PANEL TITLE
        # =========================================================

        title_cell = sheet.cell(
            start_row,
            1
        )

        title_cell.value = "Supplier KPI Summary"

        title_cell.fill = kpi_title_fill
        title_cell.font = kpi_title_font
        title_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        title_cell.border = kpi_border

        # Apply title formatting to column B BEFORE merging
        # so the highlighted title area covers the full A:B section.

        title_value_cell = sheet.cell(
            start_row,
            2
        )

        title_value_cell.fill = kpi_title_fill
        title_value_cell.border = kpi_border

        # Merge title across two columns

        sheet.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row,
            end_column=2
        )

        sheet.row_dimensions[start_row].height = 24

        # =========================================================
        # KPI TABLE HEADERS
        # =========================================================

        kpi_header_row = start_row + 1

        sheet.cell(
            kpi_header_row,
            1
        ).value = "KPI"

        sheet.cell(
            kpi_header_row,
            2
        ).value = "Value"

        for col in range(1, 3):

            cell = sheet.cell(
                kpi_header_row,
                col
            )

            cell.fill = kpi_header_fill
            cell.font = kpi_header_font
            cell.border = kpi_border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # =========================================================
        # KPI LABELS
        # =========================================================

        kpi_labels = [
            "Orders",
            "Ordered Qty",
            "Received Qty",
            "Order Fulfillment Rate %",
            "Quantity Variance",
            "Price Variance",
            "Average Delivery Days"
        ]

        # =========================================================
        # WRITE KPI PANEL
        # =========================================================

        for offset, kpi in enumerate(
            kpi_labels,
            start=2
        ):

            row = start_row + offset

            # -----------------------------------------------------
            # KPI NAME
            # -----------------------------------------------------

            label_cell = sheet.cell(
                row,
                1
            )

            label_cell.value = kpi

            label_cell.border = kpi_border

            label_cell.alignment = Alignment(
                horizontal="left",
                vertical="center"
            )

            # -----------------------------------------------------
            # KPI VALUE
            # -----------------------------------------------------

            value_cell = sheet.cell(
                row,
                2
            )

            value_cell.border = kpi_border

            value_cell.alignment = Alignment(
                horizontal="right",
                vertical="center"
            )

            # =====================================================
            # ORDERS
            # =====================================================

            if kpi == "Orders":

                if helper_end >= helper_start + 1:

                    value_cell.value = (
                        f"=COUNTA("
                        f"AA{helper_start + 1}:"
                        f"AA{helper_end}"
                        f")"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "0"

            # =====================================================
            # ORDERED QTY
            # =====================================================

            elif kpi == "Ordered Qty":

                if ordered_col:

                    ordered_letter = get_column_letter(
                        ordered_col
                    )

                    value_cell.value = (
                        f"={ordered_letter}{total_row}"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "#,##0.00"

            # =====================================================
            # RECEIVED QTY
            # =====================================================

            elif kpi == "Received Qty":

                if received_col:

                    received_letter = get_column_letter(
                        received_col
                    )

                    value_cell.value = (
                        f"={received_letter}{total_row}"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "#,##0.00"

            # =====================================================
            # ORDER FULFILLMENT RATE
            # =====================================================

            elif kpi == "Order Fulfillment Rate %":

                # -------------------------------------------------
                # IMPORTANT:
                #
                # Because the KPI HEADER ROW was added:
                #
                # start_row + 1 = KPI / Value header
                # start_row + 2 = Orders
                # start_row + 3 = Ordered Qty
                # start_row + 4 = Received Qty
                # start_row + 5 = Order Fulfillment Rate %
                #
                # Therefore:
                #
                # Ordered Qty = start_row + 3
                # Received Qty = start_row + 4
                #
                # Formula:
                #
                # Received Qty / Ordered Qty
                # -------------------------------------------------

                ordered_kpi_row = start_row + 3

                received_kpi_row = start_row + 4

                value_cell.value = (
                    f"=IF("
                    f"B{ordered_kpi_row}=0,"
                    f"0,"
                    f"B{received_kpi_row}/"
                    f"B{ordered_kpi_row}"
                    f")"
                )

                value_cell.number_format = "0.00%"

            # =====================================================
            # QUANTITY VARIANCE
            # =====================================================

            elif kpi == "Quantity Variance":

                if qty_variance_col:

                    qty_variance_letter = get_column_letter(
                        qty_variance_col
                    )

                    value_cell.value = (
                        f"={qty_variance_letter}{total_row}"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "#,##0.00"

            # =====================================================
            # PRICE VARIANCE
            # =====================================================

            elif kpi == "Price Variance":

                if price_variance_col:

                    price_variance_letter = get_column_letter(
                        price_variance_col
                    )

                    value_cell.value = (
                        f"={price_variance_letter}{total_row}"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "#,##0.00"

            # =====================================================
            # AVERAGE DELIVERY DAYS
            # =====================================================

            elif kpi == "Average Delivery Days":

                if helper_end >= helper_start + 1:

                    value_cell.value = (
                        f"=IFERROR("
                        f"AVERAGE("
                        f"AB{helper_start + 1}:"
                        f"AB{helper_end}"
                        f"),"
                        f"0"
                        f")"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "0.0"

        # =========================================================
        # KPI COLUMN WIDTHS
        # =========================================================

        sheet.column_dimensions["A"].width = max(
            sheet.column_dimensions["A"].width or 0,
            30
        )

        sheet.column_dimensions["B"].width = max(
            sheet.column_dimensions["B"].width or 0,
            18
        )

        # =========================================================
        # MONTHLY ARTICLE SUMMARY
        # =========================================================

        # IMPORTANT:
        # Added one extra KPI header row.
        #
        # Therefore the article summary starts one row lower.

        summary_start = (
            start_row
            + len(kpi_labels)
            + 5
        )

        article_start, summary_rows, article_summary = (
            create_article_summary(
                sheet,
                supplier_df,
                summary_start
            )
        )

        # =========================================================
        # CHARTS
        # =========================================================

        add_supplier_chart(
            sheet,
            article_summary
        )

    return supplier_sheet_map

# =============================================================================
# BUILD WORKBOOK
# =============================================================================

def build_workbook(report_df):

    wb = Workbook()

    wb.remove(wb.active)

    worksheet_last_rows = {}

    # Create supplier sheets FIRST
    supplier_sheet_map = create_supplier_sheets(
        wb,
        report_df,
        worksheet_last_rows
    )

    # Create the summary dataframe
    summary = create_master_summary(
        report_df
    )

    # Now create the Master Summary
    write_master_summary(
        wb,
        summary,
        supplier_sheet_map
    )

    # Move Master Summary to the first sheet
    master = wb["Master Summary"]

    wb._sheets.remove(master)
    wb._sheets.insert(0, master)

    # Make it the active sheet
    wb.active = 0

    return wb, worksheet_last_rows

###############################################################################
# EXCEL FORMATTING
###############################################################################

def format_worksheet(ws, last_data_row):
    """
    Apply professional worksheet formatting.

    Only the transaction table (rows 1 to last_data_row)
    receives column-specific formatting.
    """

    # -------------------------------------------------------------------------
    # Styles
    # -------------------------------------------------------------------------

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

    # -------------------------------------------------------------------------
    # Worksheet settings
    # -------------------------------------------------------------------------

    ws.freeze_panes = FREEZE_PANES
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{last_data_row}"

    # -------------------------------------------------------------------------
    # Transaction Table Header
    # -------------------------------------------------------------------------

    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT

    for cell in ws[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    # -------------------------------------------------------------------------
    # Store transaction headers
    # -------------------------------------------------------------------------

    headers = {
        cell.column: str(cell.value)
        for cell in ws[1]
    }

    # -------------------------------------------------------------------------
    # Format ONLY the transaction table
    # -------------------------------------------------------------------------

    for row in ws.iter_rows(
        min_row=2,
        max_row=last_data_row
    ):

        ws.row_dimensions[row[0].row].height = DEFAULT_ROW_HEIGHT

        for cell in row:

            cell.border = border

            header = headers.get(cell.column, "")

            # Dates
            if "Date" in header:

                cell.number_format = "dd-mmm-yyyy"

            # Percentages
            elif "%" in header:

                cell.number_format = "0.00%"

            # Numbers
            elif isinstance(cell.value, (int, float)):

                cell.number_format = "#,##0.00"

    # -------------------------------------------------------------------------
    # Auto-fit all columns
    # -------------------------------------------------------------------------

    for column in ws.columns:

        max_length = 0

        column_letter = get_column_letter(column[0].column)

        for cell in column:

            try:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            except Exception:

                pass

        ws.column_dimensions[column_letter].width = min(
            max_length + 3,
            40
        )

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

def add_supplier_chart(
    ws,
    article_summary
):
    """
    Creates an Ordered vs Delivered chart using the
    article_summary DataFrame.

    A visible Chart Summary table is written first,
    followed immediately by the chart.
    """

    # ---------------------------------------------------------
    # CHART SUMMARY TITLE
    # ---------------------------------------------------------

    chart_title_row = ws.max_row + 3

    ws.cell(
        chart_title_row,
        1
    ).value = "Ordered vs Delivered Summary"

    # ---------------------------------------------------------
    # HEADERS
    # ---------------------------------------------------------

    header_row = chart_title_row + 1

    headers = [
        "Article",
        "Ordered Qty",
        "Delivered Qty"
    ]

    for col, header in enumerate(headers, start=1):

        cell = ws.cell(
            header_row,
            col
        )

        cell.value = header

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_FILL
        )

        cell.font = Font(
            bold=True,
            color=HEADER_FONT
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # ---------------------------------------------------------
    # WRITE SUMMARY DATA
    # ---------------------------------------------------------

    data_start = header_row + 1

    current_row = data_start

    for _, row in article_summary.iterrows():

        ws.cell(
            current_row,
            1
        ).value = row["Article"]

        ws.cell(
            current_row,
            2
        ).value = row["Ordered"]

        ws.cell(
            current_row,
            3
        ).value = row["Delivered"]

        current_row += 1

    # ---------------------------------------------------------
    # TOTAL ROW
    # ---------------------------------------------------------

    ws.cell(current_row, 1).value = "TOTAL"

    ws.cell(
        current_row,
        2
    ).value = (
        f"=SUM(B{data_start}:B{current_row-1})"
    )

    ws.cell(
        current_row,
        3
    ).value = (
        f"=SUM(C{data_start}:C{current_row-1})"
    )

    total_row = current_row

    # ---------------------------------------------------------
    # CREATE CHART
    # ---------------------------------------------------------

    chart = BarChart()

    chart.type = "col"

    chart.style = 10

    chart.title = "Ordered vs Delivered by Article"

    chart.y_axis.title = "Quantity"

    chart.x_axis.title = "Article"

    chart.height = 8

    chart.width = 16

    chart.dLbls = DataLabelList()

    chart.dLbls.showVal = True

    data = Reference(
        ws,
        min_col=2,
        max_col=3,
        min_row=header_row,
        max_row=total_row - 1
    )

    categories = Reference(
        ws,
        min_col=1,
        min_row=data_start,
        max_row=total_row - 1
    )

    chart.add_data(
        data,
        titles_from_data=True
    )

    chart.set_categories(
        categories
    )

    # ---------------------------------------------------------
    # POSITION CHART
    # ---------------------------------------------------------

    chart_row = total_row + 2

    ws.add_chart(
        chart,
        f"A{chart_row}"
    )

###############################################################################
# MASTER DASHBOARD
###############################################################################

def add_dashboard(master_ws, report_df):
    """
    Creates a live Executive Dashboard using Excel formulas
    linked to the Master Summary table.
    """

    # ---------------------------------------------------------
    # Dashboard Title
    # ---------------------------------------------------------

    
    master_ws["A1"] = REPORT_TITLE

    master_ws["A1"].font = Font(
        bold=True,
        size=16
    )

    master_ws["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )

    master_ws["A1"].font = Font(
        bold=True,
        size=16,
        color=HEADER_FONT
    )

    master_ws.merge_cells("A1:B1")

    master_ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # ---------------------------------------------------------
    # Dashboard Styles
    # ---------------------------------------------------------

    thin = Side(style="thin")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    label_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAD3"
    )

    # ---------------------------------------------------------
    # Locate Master Summary columns
    # ---------------------------------------------------------

    headers = {
        cell.value: cell.column
        for cell in master_ws[11]
    }

    summary_last_row = master_ws.max_row

    supplier_col = get_column_letter(headers["Supplier"])
    orders_col = get_column_letter(headers["Orders"])
    ordered_col = get_column_letter(headers["Ordered Qty"])
    received_col = get_column_letter(headers["Received Qty"])
    qty_var_col = get_column_letter(headers["Qty Variance"])
    price_var_col = get_column_letter(headers["Price Variance"])
    avg_days_col = get_column_letter(headers["Average Delivery Days"])

    # ---------------------------------------------------------
    # Dashboard Labels & Formulas
    # ---------------------------------------------------------

    dashboard = [

        (
            "Total Suppliers",
            f"=COUNTA({supplier_col}12:{supplier_col}{summary_last_row})"
        ),

        (
            "Total Orders",
            f"=SUM({orders_col}12:{orders_col}{summary_last_row})"
        ),

        (
            "Total Ordered Qty",
            f"=SUM({ordered_col}12:{ordered_col}{summary_last_row})"
        ),

        (
            "Total Received Qty",
            f"=SUM({received_col}12:{received_col}{summary_last_row})"
        ),

        (
            "Overall Order Fulfillment Rate",
            f"=IF(SUM({ordered_col}12:{ordered_col}{summary_last_row})=0,"
            f"0,"
            f"SUM({received_col}12:{received_col}{summary_last_row})/"
            f"SUM({ordered_col}12:{ordered_col}{summary_last_row}))"
        ),

        (
            "Average Delivery Days",
            f"=AVERAGE({avg_days_col}12:{avg_days_col}{summary_last_row})"
        ),

        (
            "Total Price Variance",
            f"=SUM({price_var_col}12:{price_var_col}{summary_last_row})"
        ),

        (
            "Total Quantity Variance",
            f"=SUM({qty_var_col}12:{qty_var_col}{summary_last_row})"
        )

    ]

    # ---------------------------------------------------------
    # Write Dashboard
    # ---------------------------------------------------------

    start_row = 2

    for label, formula in dashboard:

        # Label Cell
        label_cell = master_ws.cell(start_row, 1)
        label_cell.value = label
        label_cell.font = Font(bold=True)
        label_cell.fill = label_fill
        label_cell.border = border

        # Value Cell
        value_cell = master_ws.cell(start_row, 2)
        value_cell.value = formula
        value_cell.border = border

        # Number Formatting
        if "Fulfillment Rate" in label:

            value_cell.number_format = "0.00%"

        elif "Average Delivery Days" in label:

            value_cell.number_format = "0.0"

        else:

            value_cell.number_format = "#,##0.00"

        start_row += 1

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

        report_df["Delivery Date"]

        .dropna()

        .max()

        .strftime("%B_%Y")

    )

    # ---------------------------------------------------------
    # Build workbook
    # ---------------------------------------------------------

    workbook, worksheet_last_rows = build_workbook(report_df)
    
    # ---------------------------------------------------------
    # Add dashboard
    # ---------------------------------------------------------

    master = workbook[MASTER_SHEET]

    add_dashboard(master, report_df)

    # ---------------------------------------------------------
    # Format worksheets
    # ---------------------------------------------------------

    for sheet in workbook.worksheets:

        if sheet.title in worksheet_last_rows:

            format_worksheet(

                sheet,

                worksheet_last_rows[sheet.title]

            )

        apply_conditional_formatting(sheet)
    # ---------------------------------------------------------
    # Return workbook and report period
    # ---------------------------------------------------------

    return save_workbook(workbook), report_period
