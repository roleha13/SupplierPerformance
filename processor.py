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

    KPIs:
        - Orders
        - Ordered Qty
        - Received Qty
        - Qty Variance
        - Variance Value
        - Average Delivery Days
        - Order Fulfillment Rate %

    Average Delivery Days is calculated once per unique
    Purchase Order instead of once per article line.
    """

    # -------------------------------------------------------------------------
    # Delivery Days
    # One record per Purchase Order
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
        df.groupby(
            "Supplier",
            as_index=False
        )
        .agg(
            Orders=("Order No.", "nunique"),

            Ordered_Qty=(
                "Ordered",
                "sum"
            ),

            Received_Qty=(
                "Booked QTY",
                "sum"
            ),

            Qty_Variance=(
                "Variance QTY",
                "sum"
            ),

            # IMPORTANT:
            # This is the monetary value of the quantity variance.
            # It is NOT unit price variance.
            Variance_Value=(
                "Variance Value",
                "sum"
            )
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
            summary["Received_Qty"]
            /
            summary["Ordered_Qty"]
        )
        .replace(
            [float("inf")],
            0
        )
        .fillna(0)
        .round(4)
    )

    # -------------------------------------------------------------------------
    # Rename Columns
    # -------------------------------------------------------------------------

    summary.rename(
        columns={
            "Ordered_Qty":
                "Ordered Qty",

            "Received_Qty":
                "Received Qty",

            "Qty_Variance":
                "Qty Variance",

            "Variance_Value":
                "Variance Value",

            "Average_Delivery_Days":
                "Average Delivery Days"
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
            False,  # Highest fulfillment first
            True,   # Lowest delivery days first
            False   # Most orders first
        ],
        inplace=True
    )

    # -------------------------------------------------------------------------
    # Reset Index
    # -------------------------------------------------------------------------

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
    Create executive dashboard KPI values.

    Variance Value represents the monetary value of the
    quantity variance and is calculated from the transaction-level
    'Variance Value' column.
    """

    # -------------------------------------------------------------------------
    # Total Ordered Quantity
    # -------------------------------------------------------------------------

    ordered = df["Ordered"].sum()

    # -------------------------------------------------------------------------
    # Total Received Quantity
    # -------------------------------------------------------------------------

    received = df["Booked QTY"].sum()

    # -------------------------------------------------------------------------
    # Overall Order Fulfillment Rate
    # -------------------------------------------------------------------------

    fill_rate = (
        received / ordered
        if ordered
        else 0
    )

    # -------------------------------------------------------------------------
    # Executive KPI Dictionary
    # -------------------------------------------------------------------------

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
            round(
                fill_rate,
                4
            ),

        "Average Delivery Days":
            round(
                df["Delivery Days"].mean(),
                1
            ),

        # IMPORTANT:
        # This is Variance Value, not Price Variance.
        "Total Variance Value":
            df["Variance Value"].sum(),

        "Total Quantity Variance":
            df["Variance QTY"].sum()

    }
# =============================================================================
# MASTER SUMMARY SHEET
# =============================================================================

def write_master_summary(
    workbook,
    summary_df,
    supplier_sheet_map
):

    ws = workbook.create_sheet(
        "Master Summary"
    )

    # -------------------------------------------------------------------------
    # Reserve rows 1-10 for Dashboard
    # -------------------------------------------------------------------------

    START_ROW = 11

    # -------------------------------------------------------------------------
    # Styles
    # -------------------------------------------------------------------------

    thin = Side(
        style="thin"
    )

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

    # -------------------------------------------------------------------------
    # Write Header Row
    # -------------------------------------------------------------------------

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

    # -------------------------------------------------------------------------
    # Header Row Height
    # -------------------------------------------------------------------------

    ws.row_dimensions[
        START_ROW
    ].height = 35

    # -------------------------------------------------------------------------
    # Write Supplier Names / Base Rows
    # -------------------------------------------------------------------------

    current_row = START_ROW + 1

    for record in summary_df.itertuples(
        index=False
    ):

        # ---------------------------------------------------------------------
        # Write Initial Values
        #
        # These values will be replaced with live formulas for KPI columns.
        # ---------------------------------------------------------------------

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

        # ---------------------------------------------------------------------
        # Supplier Hyperlink
        # ---------------------------------------------------------------------

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

        # IMPORTANT:
        # Do not use insert_rows().
        # This preserves the working hyperlinks.
        supplier_cell.hyperlink = (
            f"#'{sheet_name}'!A1"
        )

        # ---------------------------------------------------------------------
        # Hyperlink Appearance
        # ---------------------------------------------------------------------

        supplier_cell.font = Font(
            color="0563C1",
            underline="single"
        )

        supplier_cell.border = border

        current_row += 1

    # =========================================================================
    # LIVE FORMULAS
    # =========================================================================

    # -------------------------------------------------------------------------
    # Identify Master Summary Columns
    # -------------------------------------------------------------------------

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

    # IMPORTANT:
    # Correct name is Variance Value.
    variance_value_col = headers.get(
        "Variance Value"
    )

    avg_days_col = headers.get(
        "Average Delivery Days"
    )

    fulfillment_col = headers.get(
        "Order Fulfillment Rate %"
    )

    # -------------------------------------------------------------------------
    # Loop through Each Supplier
    # -------------------------------------------------------------------------

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

        supplier_sheet = workbook[
            sheet_name
        ]

        # ---------------------------------------------------------------------
        # Find Supplier KPI Panel
        # ---------------------------------------------------------------------

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

        # ---------------------------------------------------------------------
        # Safety Check
        # ---------------------------------------------------------------------

        if kpi_title_row is None:
            continue

        # ---------------------------------------------------------------------
        # Find KPI Rows by Label
        # ---------------------------------------------------------------------

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

                kpi_rows[
                    str(kpi_name)
                ] = search_row

        # =====================================================================
        # ORDERS
        # =====================================================================

        if (
            orders_col
            and "Orders" in kpi_rows
        ):

            source_row = kpi_rows[
                "Orders"
            ]

            ws.cell(
                row,
                orders_col
            ).value = (
                f"='{sheet_name}'!"
                f"B{source_row}"
            )

            ws.cell(
                row,
                orders_col
            ).number_format = "0"

        # =====================================================================
        # ORDERED QTY
        # =====================================================================

        if (
            ordered_col
            and "Ordered Qty" in kpi_rows
        ):

            source_row = kpi_rows[
                "Ordered Qty"
            ]

            ws.cell(
                row,
                ordered_col
            ).value = (
                f"='{sheet_name}'!"
                f"B{source_row}"
            )

            ws.cell(
                row,
                ordered_col
            ).number_format = "#,##0.00"

        # =====================================================================
        # RECEIVED QTY
        # =====================================================================

        if (
            received_col
            and "Received Qty" in kpi_rows
        ):

            source_row = kpi_rows[
                "Received Qty"
            ]

            ws.cell(
                row,
                received_col
            ).value = (
                f"='{sheet_name}'!"
                f"B{source_row}"
            )

            ws.cell(
                row,
                received_col
            ).number_format = "#,##0.00"

        # =====================================================================
        # QUANTITY VARIANCE
        # =====================================================================

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
                f"='{sheet_name}'!"
                f"B{source_row}"
            )

            ws.cell(
                row,
                qty_variance_col
            ).number_format = "#,##0.00"

        # =====================================================================
        # VARIANCE VALUE
        # =====================================================================

        if (
            variance_value_col
            and "Variance Value" in kpi_rows
        ):

            source_row = kpi_rows[
                "Variance Value"
            ]

            ws.cell(
                row,
                variance_value_col
            ).value = (
                f"='{sheet_name}'!"
                f"B{source_row}"
            )

            # Monetary value
            ws.cell(
                row,
                variance_value_col
            ).number_format = "#,##0.00"

        # =====================================================================
        # AVERAGE DELIVERY DAYS
        # =====================================================================

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
                f"='{sheet_name}'!"
                f"B{source_row}"
            )

            ws.cell(
                row,
                avg_days_col
            ).number_format = "0.0"

        # =====================================================================
        # ORDER FULFILLMENT RATE
        # =====================================================================

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

    # =========================================================================
    # GENERAL FORMATTING
    # =========================================================================

    # -------------------------------------------------------------------------
    # Ensure All Data Cells Have Borders
    # -------------------------------------------------------------------------

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

    # -------------------------------------------------------------------------
    # Percentage Column
    # -------------------------------------------------------------------------

    if fulfillment_col:

        for row in range(
            START_ROW + 1,
            ws.max_row + 1
        ):

            ws.cell(
                row,
                fulfillment_col
            ).number_format = "0.00%"

    # -------------------------------------------------------------------------
    # Delivery Days
    # -------------------------------------------------------------------------

    if avg_days_col:

        for row in range(
            START_ROW + 1,
            ws.max_row + 1
        ):

            ws.cell(
                row,
                avg_days_col
            ).number_format = "0.0"

    # =========================================================================
    # AUTO FILTER
    # =========================================================================

    last_col = get_column_letter(
        ws.max_column
    )

    ws.auto_filter.ref = (
        f"A{START_ROW}:"
        f"{last_col}{ws.max_row}"
    )


    # =========================================================================
    # FREEZE PANES
    # =========================================================================

    ws.freeze_panes = (
        f"A{START_ROW + 1}"
    )

    # =========================================================================
    # MASTER SUMMARY COLUMN WIDTHS
    # =========================================================================

    format_master_summary_columns(
        ws
    )

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

        "F": 20,   # Variance Value

        "G": 25,   # Average Delivery Days

        "H": 27    # Order Fulfillment Rate %

    }

    # -------------------------------------------------------------------------
    # Apply Column Widths
    # -------------------------------------------------------------------------

    for column, width in column_widths.items():

        ws.column_dimensions[
            column
        ].width = width
# =============================================================================
# SUPPLIER KPI PANEL
# =============================================================================

def supplier_kpis(df: pd.DataFrame):
    """
    Calculate supplier-level KPI values.

    KPI definitions:
        Orders
        Ordered Qty
        Received Qty
        Order Fulfillment Rate %
        Quantity Variance
        Variance Value
        Average Delivery Days

    IMPORTANT:
    Variance Value is the monetary value associated with the
    quantity variance. It should not be labelled Price Variance.
    """

    # -------------------------------------------------------------------------
    # Ordered Quantity
    # -------------------------------------------------------------------------

    ordered = df["Ordered"].sum()

    # -------------------------------------------------------------------------
    # Received Quantity
    # -------------------------------------------------------------------------

    received = df["Booked QTY"].sum()

    # -------------------------------------------------------------------------
    # Order Fulfillment Rate
    # -------------------------------------------------------------------------

    fill_rate = (
        received / ordered
        if ordered
        else 0
    )

    # -------------------------------------------------------------------------
    # Average Delivery Days
    #
    # Calculate once per unique Purchase Order rather than
    # once per article line.
    # -------------------------------------------------------------------------

    delivery_days = (
        df[
            [
                "Order No.",
                "Delivery Days"
            ]
        ]
        .drop_duplicates(
            subset=["Order No."]
        )[
            "Delivery Days"
        ]
        .mean()
    )

    # -------------------------------------------------------------------------
    # Supplier KPI Results
    # -------------------------------------------------------------------------

    return [

        # Number of unique Purchase Orders
        (
            "Orders",
            df["Order No."].nunique()
        ),

        # Total quantity ordered
        (
            "Ordered Qty",
            ordered
        ),

        # Total quantity received/booked
        (
            "Received Qty",
            received
        ),

        # Received Qty / Ordered Qty
        (
            "Order Fulfillment Rate %",
            round(
                fill_rate,
                4
            )
        ),

        # Quantity shortfall / excess
        (
            "Quantity Variance",
            df["Variance QTY"].sum()
        ),

        # Monetary value of quantity variance
        #
        # IMPORTANT:
        # This is NOT unit price variance.
        (
            "Variance Value",
            df["Variance Value"].sum()
        ),

        # Average delivery time
        (
            "Average Delivery Days",
            round(
                delivery_days,
                1
            )
        )

    ]


# =============================================================================
# MONTHLY ARTICLE SUMMARY
# =============================================================================

def create_article_summary(sheet, supplier_df, start_row):
    """
    Creates the Monthly Article Summary section.

    Structure:

        Article | Ordered Qty | Delivered Qty | Qty Variance | No. of Orders

    Each article summary row is followed by hidden PO detail rows.

    Article totals are calculated using live Excel formulas that reference
    the hidden PO detail rows.

    Column mapping in the hidden detail rows:

        A = Order No.
        B = Ordered
        C = Booked QTY
        D = Variance QTY
        E = Blank

    The detail rows are hidden and grouped underneath each article.
    """

    from openpyxl.styles import (
        Font,
        PatternFill,
        Alignment,
        Border,
        Side
    )

    from openpyxl.utils import get_column_letter

    # =========================================================================
    # 1. FIND TRANSACTION TABLE HEADER ROW
    # =========================================================================

    required_headers = [
        "Article",
        "Order No.",
        "Ordered",
        "Booked QTY",
        "Variance QTY"
    ]

    transaction_header_row = None

    search_limit = min(sheet.max_row, 100)

    for row in range(1, search_limit + 1):

        row_values = [
            sheet.cell(
                row=row,
                column=col
            ).value
            for col in range(1, sheet.max_column + 1)
        ]

        if all(
            header in row_values
            for header in required_headers
        ):
            transaction_header_row = row
            break

    if transaction_header_row is None:

        raise ValueError(
            f"Could not find the transaction table headers "
            f"in supplier sheet '{sheet.title}'. "
            f"Required headers: {required_headers}"
        )

    transaction_first_data_row = transaction_header_row + 1

    # =========================================================================
    # 2. FIND TRANSACTION COLUMN NUMBERS
    # =========================================================================

    header_columns = {}

    for col in range(1, sheet.max_column + 1):

        value = sheet.cell(
            row=transaction_header_row,
            column=col
        ).value

        if value in required_headers:

            header_columns[value] = col

    # -------------------------------------------------------------------------
    # Validate that all required columns were found
    # -------------------------------------------------------------------------

    missing_columns = [
        header
        for header in required_headers
        if header not in header_columns
    ]

    if missing_columns:

        raise ValueError(
            f"Missing transaction columns in supplier sheet "
            f"'{sheet.title}': {missing_columns}"
        )

    # =========================================================================
    # 3. CONVERT TRANSACTION COLUMNS TO EXCEL LETTERS
    # =========================================================================

    transaction_article_letter = get_column_letter(
        header_columns["Article"]
    )

    transaction_order_letter = get_column_letter(
        header_columns["Order No."]
    )

    transaction_ordered_letter = get_column_letter(
        header_columns["Ordered"]
    )

    transaction_booked_letter = get_column_letter(
        header_columns["Booked QTY"]
    )

    transaction_variance_letter = get_column_letter(
        header_columns["Variance QTY"]
    )

    # =========================================================================
    # 4. CREATE ARTICLE SUMMARY DATA
    # =========================================================================

    if supplier_df.empty:

        article_summary = supplier_df.copy()

        article_summary = article_summary.assign(
            Ordered=pd.Series(dtype="float64"),
            Delivered=pd.Series(dtype="float64"),
            Variance=pd.Series(dtype="float64"),
            Order_Frequency=pd.Series(dtype="int64")
        )

    else:

        article_summary = (
            supplier_df
            .groupby("Article", as_index=False)
            .agg(
                Ordered=("Ordered", "sum"),
                Delivered=("Booked QTY", "sum"),
                Variance=("Variance QTY", "sum"),
                Order_Frequency=("Order No.", "nunique")
            )
            .sort_values(
                "Article",
                na_position="last"
            )
        )

    # =========================================================================
    # 5. CREATE ARTICLE SUMMARY TITLE
    # =========================================================================

    summary_row = start_row

    sheet.cell(
        row=summary_row,
        column=1,
        value="Monthly Article Summary"
    )

    sheet.merge_cells(
        start_row=summary_row,
        start_column=1,
        end_row=summary_row,
        end_column=5
    )

    title_cell = sheet.cell(
        row=summary_row,
        column=1
    )

    title_cell.font = Font(
        bold=True,
        size=12
    )

    title_cell.fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )
        

    title_cell.alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    # =========================================================================
    # 6. CREATE SUMMARY HEADERS
    # =========================================================================

    header_row = summary_row + 1

    summary_headers = [
        "Article",
        "Ordered Qty",
        "Delivered Qty",
        "Qty Variance",
        "No. of Orders"
    ]

    for col, header in enumerate(
        summary_headers,
        start=1
    ):

        cell = sheet.cell(
            row=header_row,
            column=col,
            value=header
        )

        cell.font = Font(
            bold=True
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAD3"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # =========================================================================
    # 7. CREATE FORMATTING STYLES
    # =========================================================================

    thin_side = Side(
        style="thin",
        color="B7B7B7"
    )

    thin_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    detail_fill = PatternFill(
        fill_type="solid",
        fgColor="F7F7F7"
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )
        

    # =========================================================================
    # 8. MAP DATAFRAME INDEX TO TRANSACTION SHEET ROW
    # =========================================================================
    #
    # This allows the hidden article detail rows to reference the correct
    # transaction rows in the supplier sheet.
    #
    # We intentionally preserve the original supplier_df index.
    # =========================================================================

    transaction_row_map = {
        index: transaction_first_data_row + position
        for position, index in enumerate(supplier_df.index)
    }

    if len(transaction_row_map) != len(supplier_df):

        raise ValueError(
            "Supplier dataframe index contains duplicate values. "
            "Cannot reliably map article detail rows to transaction rows."
        )

    # =========================================================================
    # 9. WRITE ARTICLE SUMMARY + HIDDEN PO DETAIL ROWS
    # =========================================================================

    current_row = header_row + 1

    summary_rows = []

    for _, article in article_summary.iterrows():

        article_name = article["Article"]

        # ---------------------------------------------------------------------
        # Get all transactions belonging to this article
        # ---------------------------------------------------------------------

        article_rows = supplier_df[
            supplier_df["Article"] == article_name
        ].copy()

        # ---------------------------------------------------------------------
        # Sort PO details
        # ---------------------------------------------------------------------

        sort_columns = []

        if "Order Date" in article_rows.columns:
            sort_columns.append("Order Date")

        if "Order No." in article_rows.columns:
            sort_columns.append("Order No.")

        if sort_columns:

            article_rows = article_rows.sort_values(
                by=sort_columns,
                na_position="last"
            )

        # ---------------------------------------------------------------------
        # ARTICLE SUMMARY ROW
        # ---------------------------------------------------------------------

        article_row = current_row

        summary_rows.append(article_row)

        sheet.cell(
            row=article_row,
            column=1,
            value=article_name
        )

        # ---------------------------------------------------------------------
        # Determine hidden PO detail row range
        # ---------------------------------------------------------------------

        detail_start_row = article_row + 1

        detail_end_row = (
            detail_start_row
            + len(article_rows)
            - 1
        )

        # ---------------------------------------------------------------------
        # ARTICLE ORDERED QTY
        #
        # Calculated from hidden PO detail rows.
        # ---------------------------------------------------------------------

        if len(article_rows) > 0:

            sheet.cell(
                row=article_row,
                column=2,
                value=(
                    f"=SUM("
                    f"B{detail_start_row}:"
                    f"B{detail_end_row}"
                    f")"
                )
            )

        else:

            sheet.cell(
                row=article_row,
                column=2,
                value=0
            )

        # ---------------------------------------------------------------------
        # ARTICLE DELIVERED QTY
        # ---------------------------------------------------------------------

        if len(article_rows) > 0:

            sheet.cell(
                row=article_row,
                column=3,
                value=(
                    f"=SUM("
                    f"C{detail_start_row}:"
                    f"C{detail_end_row}"
                    f")"
                )
            )

        else:

            sheet.cell(
                row=article_row,
                column=3,
                value=0
            )

        # ---------------------------------------------------------------------
        # ARTICLE QUANTITY VARIANCE
        #
        # Ordered Qty - Delivered Qty
        # ---------------------------------------------------------------------

        sheet.cell(
            row=article_row,
            column=4,
            value=(
                f"=B{article_row}-C{article_row}"
            )
        )

        # ---------------------------------------------------------------------
        # NUMBER OF ORDERS
        #
        # Counts unique PO numbers from the hidden detail rows.
        # ---------------------------------------------------------------------

        if len(article_rows) > 0:

            sheet.cell(
                row=article_row,
                column=5,
                value=(
                    f'=SUMPRODUCT(('
                    f'A{detail_start_row}:A{detail_end_row}<>""'
                    f')/COUNTIF('
                    f'A{detail_start_row}:A{detail_end_row},'
                    f'A{detail_start_row}:A{detail_end_row}'
                    f'))'
                )
            )

        else:

            sheet.cell(
                row=article_row,
                column=5,
                value=0
            )

        # ---------------------------------------------------------------------
        # FORMAT ARTICLE SUMMARY ROW
        # ---------------------------------------------------------------------

        for col in range(1, 6):

            cell = sheet.cell(
                row=article_row,
                column=col
            )

            cell.font = Font(
                bold=True
            )

            cell.border = thin_border

            cell.alignment = Alignment(
                horizontal=(
                    "left"
                    if col == 1
                    else "right"
                ),
                vertical="center"
            )

        # =========================================================================
        # WRITE HIDDEN PO DETAIL ROWS
        # =========================================================================

        for _, order in article_rows.iterrows():

            detail_row = current_row + 1

            # -----------------------------------------------------------------
            # Find corresponding transaction table row
            # -----------------------------------------------------------------

            if order.name not in transaction_row_map:

                raise ValueError(
                    f"Could not map dataframe row '{order.name}' "
                    f"to a transaction row in supplier sheet "
                    f"'{sheet.title}'."
                )

            transaction_row = transaction_row_map[
                order.name
            ]

            # -----------------------------------------------------------------
            # COLUMN A
            #
            # Live formula pointing to Order No.
            # -----------------------------------------------------------------

            detail_article = sheet.cell(
                row=detail_row,
                column=1
            )

            detail_article.value = (
                f"={transaction_order_letter}"
                f"{transaction_row}"
            )

            # -----------------------------------------------------------------
            # COLUMN B
            #
            # Live formula pointing to Ordered Qty.
            # -----------------------------------------------------------------

            detail_ordered = sheet.cell(
                row=detail_row,
                column=2
            )

            detail_ordered.value = (
                f"={transaction_ordered_letter}"
                f"{transaction_row}"
            )

            # -----------------------------------------------------------------
            # COLUMN C
            #
            # Live formula pointing to Booked QTY.
            # -----------------------------------------------------------------

            detail_delivered = sheet.cell(
                row=detail_row,
                column=3
            )

            detail_delivered.value = (
                f"={transaction_booked_letter}"
                f"{transaction_row}"
            )

            # -----------------------------------------------------------------
            # COLUMN D
            #
            # Live formula pointing to Variance QTY.
            # -----------------------------------------------------------------

            detail_variance = sheet.cell(
                row=detail_row,
                column=4
            )

            detail_variance.value = (
                f"={transaction_variance_letter}"
                f"{transaction_row}"
            )

            # -----------------------------------------------------------------
            # COLUMN E
            #
            # Intentionally blank for PO detail rows.
            # -----------------------------------------------------------------

            sheet.cell(
                row=detail_row,
                column=5,
                value=None
            )

            # -----------------------------------------------------------------
            # FORMAT DETAIL ROW
            # -----------------------------------------------------------------

            for col in range(1, 6):

                cell = sheet.cell(
                    row=detail_row,
                    column=col
                )

                cell.fill = detail_fill

                cell.border = thin_border

                cell.font = Font(
                    size=9
                )

                cell.alignment = Alignment(
                    horizontal=(
                        "left"
                        if col == 1
                        else "right"
                    ),
                    vertical="center"
                )

            # -----------------------------------------------------------------
            # INDENT PO NUMBER
            # -----------------------------------------------------------------

            detail_article.alignment = Alignment(
                horizontal="left",
                vertical="center",
                indent=1
            )

            # -----------------------------------------------------------------
            # HIDE / GROUP PO DETAIL ROW
            # -----------------------------------------------------------------

            sheet.row_dimensions[
                detail_row
            ].outlineLevel = 1

            sheet.row_dimensions[
                detail_row
            ].hidden = True

            current_row = detail_row

        # ---------------------------------------------------------------------
        # COLLAPSE PO DETAIL ROWS UNDER ARTICLE
        # ---------------------------------------------------------------------

        if len(article_rows) > 0:

            sheet.row_dimensions[
                article_row
            ].collapsed = True

    # =========================================================================
    # 10. FORMAT SUMMARY COLUMNS
    # =========================================================================

    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 15
    sheet.column_dimensions["C"].width = 16
    sheet.column_dimensions["D"].width = 15
    sheet.column_dimensions["E"].width = 15

    # =========================================================================
    # 11. NUMBER FORMATTING
    # =========================================================================

    for row in range(
        header_row + 1,
        current_row + 1
    ):

        sheet.cell(
            row=row,
            column=2
        ).number_format = "#,##0.00"

        sheet.cell(
            row=row,
            column=3
        ).number_format = "#,##0.00"

        sheet.cell(
            row=row,
            column=4
        ).number_format = "#,##0.00"

        sheet.cell(
            row=row,
            column=5
        ).number_format = "0"

    # =========================================================================
    # 12. ENABLE EXCEL OUTLINE / GROUPING
    # =========================================================================

    sheet.sheet_properties.outlinePr.summaryBelow = True

    # =========================================================================
    # 13. RETURN RESULTS
    # =========================================================================

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

        # =========================================================
        # IDENTIFY TRANSACTION COLUMNS DYNAMICALLY
        # =========================================================

        transaction_headers = {
            str(sheet.cell(1, col).value).strip(): col
            for col in range(
                1,
                sheet.max_column + 1
            )
        }

        # ---------------------------------------------------------
        # GET IMPORTANT TRANSACTION COLUMNS
        # ---------------------------------------------------------

        ordered_col = transaction_headers.get(
            "Ordered"
        )

        booked_qty_col = transaction_headers.get(
            "Booked QTY"
        )

        variance_qty_col = transaction_headers.get(
            "Variance QTY"
        )

        po_price_col = transaction_headers.get(
            "PO Price"
        )

        booked_price_col = transaction_headers.get(
            "Booked Price"
        )

        variance_price_col = transaction_headers.get(
            "Variance Price"
        )

        variance_value_col = transaction_headers.get(
            "Variance Value"
        )

        # =========================================================
        # GET COLUMN LETTERS
        # =========================================================

        ordered_letter = (
            get_column_letter(ordered_col)
            if ordered_col
            else None
        )

        booked_qty_letter = (
            get_column_letter(booked_qty_col)
            if booked_qty_col
            else None
        )

        variance_qty_letter = (
            get_column_letter(variance_qty_col)
            if variance_qty_col
            else None
        )

        po_price_letter = (
            get_column_letter(po_price_col)
            if po_price_col
            else None
        )

        booked_price_letter = (
            get_column_letter(booked_price_col)
            if booked_price_col
            else None
        )

        variance_price_letter = (
            get_column_letter(variance_price_col)
            if variance_price_col
            else None
        )

        variance_value_letter = (
            get_column_letter(variance_value_col)
            if variance_value_col
            else None
        )

        # =========================================================
        # LIVE TRANSACTION FORMULAS
        #
        # These calculations are performed directly in Excel.
        #
        # Variance QTY
        #     = Ordered - Booked QTY
        #
        # Variance Price
        #     = Booked Price - PO Price
        #
        # Variance Value
        #     = Variance QTY × PO Price
        #
        # IMPORTANT:
        # "Variance Value" is the monetary variance used by the
        # TOTAL row and Supplier KPI Summary.
        # =========================================================

        for row_num in range(
            2,
            last_data_row + 1
        ):

            # -----------------------------------------------------
            # QUANTITY VARIANCE
            # Ordered - Booked QTY
            # -----------------------------------------------------

            if (
                variance_qty_col
                and ordered_letter
                and booked_qty_letter
            ):

                sheet.cell(
                    row_num,
                    variance_qty_col
                ).value = (
                    f"={ordered_letter}{row_num}"
                    f"-{booked_qty_letter}{row_num}"
                )

                sheet.cell(
                    row_num,
                    variance_qty_col
                ).number_format = "#,##0.00"

            # -----------------------------------------------------
            # PRICE VARIANCE
            # Booked Price - PO Price
            #
            # This remains a separate unit-price calculation.
            # It is NOT used as the KPI monetary variance.
            # -----------------------------------------------------

            if (
                variance_price_col
                and booked_price_letter
                and po_price_letter
            ):

                sheet.cell(
                    row_num,
                    variance_price_col
                ).value = (
                    f"={booked_price_letter}{row_num}"
                    f"-{po_price_letter}{row_num}"
                )

                sheet.cell(
                    row_num,
                    variance_price_col
                ).number_format = "#,##0.00"

            # -----------------------------------------------------
            # VARIANCE VALUE
            #
            # Quantity Variance × PO Price
            #
            # THIS is the monetary variance used by the KPI.
            # -----------------------------------------------------

            if (
                variance_value_col
                and variance_qty_letter
                and po_price_letter
            ):

                sheet.cell(
                    row_num,
                    variance_value_col
                ).value = (
                    f"={variance_qty_letter}{row_num}"
                    f"*{po_price_letter}{row_num}"
                )

                sheet.cell(
                    row_num,
                    variance_value_col
                ).number_format = "#,##0.00"

        # =========================================================
        # IMPORTANT:
        # Keep worksheet_last_rows pointing ONLY to the transaction
        # table.
        #
        # This ensures format_worksheet() does not format the TOTAL
        # row, helper table, KPI panel, article summary, or charts
        # as transaction data.
        # =========================================================

        worksheet_last_rows[
            sheet.title
        ] = last_data_row

        # =========================================================
        # TOTAL ROW
        # =========================================================

        total_row = last_data_row + 1

        sheet.cell(
            total_row,
            1
        ).value = "TOTAL"

        # =========================================================
        # TOTAL - ORDERED QTY
        # =========================================================

        if ordered_col:

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

        if booked_qty_col:

            sheet.cell(
                total_row,
                booked_qty_col
            ).value = (
                f"=SUM("
                f"{booked_qty_letter}2:"
                f"{booked_qty_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL - QUANTITY VARIANCE
        # =========================================================

        if variance_qty_col:

            sheet.cell(
                total_row,
                variance_qty_col
            ).value = (
                f"=SUM("
                f"{variance_qty_letter}2:"
                f"{variance_qty_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL - VARIANCE VALUE
        #
        # IMPORTANT:
        # This is the monetary variance used by the KPI and
        # Master Summary.
        # =========================================================

        if variance_value_col:

            sheet.cell(
                total_row,
                variance_value_col
            ).value = (
                f"=SUM("
                f"{variance_value_letter}2:"
                f"{variance_value_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL - VARIANCE PRICE
        #
        # This is retained as a separate informational total.
        # It does NOT replace Variance Value.
        # =========================================================

        if variance_price_col:

            sheet.cell(
                total_row,
                variance_price_col
            ).value = (
                f"=SUM("
                f"{variance_price_letter}2:"
                f"{variance_price_letter}{last_data_row}"
                f")"
            )

        # =========================================================
        # TOTAL ROW FORMATTING
        # =========================================================

        thin = Side(
            style="thin"
        )

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
        # TOTAL NUMBER FORMATTING
        # =========================================================

        if ordered_col:

            sheet.cell(
                total_row,
                ordered_col
            ).number_format = "#,##0.00"

        if booked_qty_col:

            sheet.cell(
                total_row,
                booked_qty_col
            ).number_format = "#,##0.00"

        if variance_qty_col:

            sheet.cell(
                total_row,
                variance_qty_col
            ).number_format = "#,##0.00"

        if variance_price_col:

            sheet.cell(
                total_row,
                variance_price_col
            ).number_format = "#,##0.00"

        if variance_value_col:

            sheet.cell(
                total_row,
                variance_value_col
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

        title_cell.value = (
            "Supplier KPI Summary"
        )

        title_cell.fill = kpi_title_fill
        title_cell.font = kpi_title_font

        title_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        title_cell.border = kpi_border

        # ---------------------------------------------------------
        # APPLY TITLE FORMATTING TO COLUMN B BEFORE MERGING
        # ---------------------------------------------------------

        title_value_cell = sheet.cell(
            start_row,
            2
        )

        title_value_cell.fill = kpi_title_fill
        title_value_cell.border = kpi_border

        # ---------------------------------------------------------
        # MERGE TITLE ACROSS A:B
        # ---------------------------------------------------------

        sheet.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row,
            end_column=2
        )

        sheet.row_dimensions[
            start_row
        ].height = 24

        # =========================================================
        # BACK TO MASTER SUMMARY LINK
        # =========================================================

        navigation_cell = sheet.cell(
            start_row,
            4
        )

        navigation_cell.value = (
            "← Back to Master Summary"
        )

        navigation_cell.hyperlink = (
            "#'Master Summary'!A1"
        )

        navigation_cell.font = Font(
            bold=True,
            color="0563C1",
            underline="single"
        )

        navigation_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

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
            "Variance Value",
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

                if booked_qty_col:

                    value_cell.value = (
                        f"={booked_qty_letter}{total_row}"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "#,##0.00"

            # =====================================================
            # ORDER FULFILLMENT RATE
            # =====================================================

            elif kpi == "Order Fulfillment Rate %":

                # Because the KPI header row was added:
                #
                # start_row + 1 = KPI / Value header
                # start_row + 2 = Orders
                # start_row + 3 = Ordered Qty
                # start_row + 4 = Received Qty
                # start_row + 5 = Order Fulfillment Rate %

                ordered_kpi_row = (
                    start_row + 3
                )

                received_kpi_row = (
                    start_row + 4
                )

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

                if variance_qty_col:

                    value_cell.value = (
                        f"={variance_qty_letter}{total_row}"
                    )

                else:

                    value_cell.value = 0

                value_cell.number_format = "#,##0.00"

            # =====================================================
            # VARIANCE VALUE
            #
            # IMPORTANT:
            # The KPI is explicitly called "Variance Value".
            #
            # It references the TOTAL of the "Variance Value"
            # transaction column.
            # =====================================================

            elif kpi == "Variance Value":

                if variance_value_col:

                    value_cell.value = (
                        f"={variance_value_letter}{total_row}"
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

        sheet.column_dimensions["D"].width = max(
            sheet.column_dimensions["D"].width or 0,
            28
        )

        # =========================================================
        # MONTHLY ARTICLE SUMMARY
        # =========================================================

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

# =============================================================================
# MASTER DASHBOARD
# =============================================================================

def add_dashboard(master_ws, report_df):
    """
    Creates a live Executive Dashboard using Excel formulas
    linked to the Master Summary table.

    IMPORTANT:
    The Master Summary contains "Variance Value", not "Price Variance".

    "Variance Value" represents the monetary value of the
    quantity variance.

    "Variance Price" is a separate transaction-level unit
    price variance and is not part of the Master Summary KPI.
    """

    # ---------------------------------------------------------
    # Dashboard Title
    # ---------------------------------------------------------

    master_ws["A1"] = REPORT_TITLE

    master_ws["A1"].font = Font(
        bold=True,
        size=16,
        color=HEADER_FONT
    )

    master_ws["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )

    master_ws.merge_cells("A1:B1")

    master_ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # ---------------------------------------------------------
    # Dashboard Styles
    # ---------------------------------------------------------

    thin = Side(
        style="thin"
    )

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
    #
    # Master Summary header is on row 11.
    #
    # IMPORTANT:
    # The correct monetary variance column is
    # "Variance Value".
    #
    # There is NO "Price Variance" column in the
    # Master Summary.
    # ---------------------------------------------------------

    headers = {
        str(cell.value).strip(): cell.column
        for cell in master_ws[11]
        if cell.value is not None
    }

    # ---------------------------------------------------------
    # Validate Required Columns
    # ---------------------------------------------------------

    required_headers = [
        "Supplier",
        "Orders",
        "Ordered Qty",
        "Received Qty",
        "Qty Variance",
        "Variance Value",
        "Average Delivery Days"
    ]

    missing_headers = [
        header
        for header in required_headers
        if header not in headers
    ]

    if missing_headers:

        raise ValueError(
            "Master Summary is missing the following required "
            "dashboard columns:\n\n"
            + "\n".join(
                f"- {header}"
                for header in missing_headers
            )
        )

    # ---------------------------------------------------------
    # Convert Column Numbers to Excel Letters
    # ---------------------------------------------------------

    summary_last_row = master_ws.max_row

    supplier_col = get_column_letter(
        headers["Supplier"]
    )

    orders_col = get_column_letter(
        headers["Orders"]
    )

    ordered_col = get_column_letter(
        headers["Ordered Qty"]
    )

    received_col = get_column_letter(
        headers["Received Qty"]
    )

    qty_var_col = get_column_letter(
        headers["Qty Variance"]
    )

    # ---------------------------------------------------------
    # IMPORTANT FIX
    #
    # Use "Variance Value" instead of "Price Variance".
    #
    # This matches create_master_summary() and the
    # Supplier KPI Summary.
    # ---------------------------------------------------------

    variance_value_col = get_column_letter(
        headers["Variance Value"]
    )

    avg_days_col = get_column_letter(
        headers["Average Delivery Days"]
    )

    # ---------------------------------------------------------
    # Dashboard Labels & Formulas
    # ---------------------------------------------------------

    dashboard = [

        # =====================================================
        # TOTAL SUPPLIERS
        # =====================================================

        (
            "Total Suppliers",

            f"=COUNTA("
            f"{supplier_col}12:"
            f"{supplier_col}{summary_last_row}"
            f")"
        ),

        # =====================================================
        # TOTAL ORDERS
        # =====================================================

        (
            "Total Orders",

            f"=SUM("
            f"{orders_col}12:"
            f"{orders_col}{summary_last_row}"
            f")"
        ),

        # =====================================================
        # TOTAL ORDERED QUANTITY
        # =====================================================

        (
            "Total Ordered Qty",

            f"=SUM("
            f"{ordered_col}12:"
            f"{ordered_col}{summary_last_row}"
            f")"
        ),

        # =====================================================
        # TOTAL RECEIVED QUANTITY
        # =====================================================

        (
            "Total Received Qty",

            f"=SUM("
            f"{received_col}12:"
            f"{received_col}{summary_last_row}"
            f")"
        ),

        # =====================================================
        # OVERALL ORDER FULFILLMENT RATE
        # =====================================================

        (
            "Overall Order Fulfillment Rate",

            f"=IF("
            f"SUM("
            f"{ordered_col}12:"
            f"{ordered_col}{summary_last_row}"
            f")=0,"
            f"0,"
            f"SUM("
            f"{received_col}12:"
            f"{received_col}{summary_last_row}"
            f")/"
            f"SUM("
            f"{ordered_col}12:"
            f"{ordered_col}{summary_last_row}"
            f")"
            f")"
        ),

        # =====================================================
        # AVERAGE DELIVERY DAYS
        # =====================================================

        (
            "Average Delivery Days",

            f"=IFERROR("
            f"AVERAGE("
            f"{avg_days_col}12:"
            f"{avg_days_col}{summary_last_row}"
            f"),"
            f"0"
            f")"
        ),

        # =====================================================
        # TOTAL VARIANCE VALUE
        #
        # IMPORTANT:
        # This replaces the incorrect "Total Price Variance".
        #
        # Variance Value = Quantity Variance × PO Price
        # at transaction level.
        # =====================================================

        (
            "Total Variance Value",

            f"=SUM("
            f"{variance_value_col}12:"
            f"{variance_value_col}{summary_last_row}"
            f")"
        ),

        # =====================================================
        # TOTAL QUANTITY VARIANCE
        # =====================================================

        (
            "Total Quantity Variance",

            f"=SUM("
            f"{qty_var_col}12:"
            f"{qty_var_col}{summary_last_row}"
            f")"
        )

    ]

    # ---------------------------------------------------------
    # Write Dashboard
    # ---------------------------------------------------------

    start_row = 2

    for label, formula in dashboard:

        # -----------------------------------------------------
        # Label Cell
        # -----------------------------------------------------

        label_cell = master_ws.cell(
            start_row,
            1
        )

        label_cell.value = label

        label_cell.font = Font(
            bold=True
        )

        label_cell.fill = label_fill

        label_cell.border = border

        label_cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        # -----------------------------------------------------
        # Value Cell
        # -----------------------------------------------------

        value_cell = master_ws.cell(
            start_row,
            2
        )

        value_cell.value = formula

        value_cell.border = border

        value_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        # -----------------------------------------------------
        # Number Formatting
        # -----------------------------------------------------

        if "Fulfillment Rate" in label:

            value_cell.number_format = "0.00%"

        elif "Average Delivery Days" in label:

            value_cell.number_format = "0.0"

        else:

            value_cell.number_format = "#,##0.00"

        # -----------------------------------------------------
        # Row Height
        # -----------------------------------------------------

        master_ws.row_dimensions[
            start_row
        ].height = 22

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
