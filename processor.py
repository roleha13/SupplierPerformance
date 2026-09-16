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


###############################################################################
# ARTICLE SUMMARY (PIVOT STYLE)
###############################################################################

def create_article_summary(sheet, supplier_df, start_row):
    """
    Creates a professional pivot-style Monthly Article Summary with
    expandable Order Number details.

    Structure:

        Transaction Table
              ↓
        Hidden PO Detail Rows
              ↓
        Monthly Article Summary

    The hidden PO detail rows use LIVE EXCEL FORMULAS referencing the
    supplier transaction table.

    The visible Article Summary rows then use LIVE EXCEL FORMULAS
    referencing the hidden PO detail rows.

    Monthly Article Summary columns:

        A = Article / Order Number
        B = Ordered Qty
        C = Delivered Qty
        D = Qty Variance
        E = No. of Orders

    Hidden PO detail rows:

        A = Order Number
        B = Ordered Qty
        C = Delivered Qty
        D = Qty Variance
        E = Blank

    No. of Orders counts UNIQUE Order Numbers from Column A of the
    hidden PO detail rows.

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
    # FIND TRANSACTION TABLE
    # -------------------------------------------------------------------------
    #
    # The hidden PO detail rows will reference the actual supplier transaction
    # table instead of writing static pandas values.
    #
    # We identify the transaction table columns by their header names.
    #
    # -------------------------------------------------------------------------

    transaction_header_row = None

    required_headers = [
        "Article",
        "Order No.",
        "Ordered",
        "Booked QTY",
        "Variance QTY"
    ]

    header_positions = {}

    # Search the worksheet for the transaction table headers.
    #
    # We search the first 100 rows because the supplier transaction table
    # normally appears near the top of the supplier sheet.

    for row in range(1, min(sheet.max_row, 100) + 1):

        row_values = {}

        for col in range(1, sheet.max_column + 1):
            value = sheet.cell(row, col).value

            if value is not None:
                row_values[str(value).strip()] = col

        if all(header in row_values for header in required_headers):

            transaction_header_row = row
            header_positions = row_values
            break

    if transaction_header_row is None:
        raise ValueError(
            "Could not find the supplier transaction table headers "
            "required for the Monthly Article Summary."
        )

    # -------------------------------------------------------------------------
    # TRANSACTION TABLE COLUMN NUMBERS
    # -------------------------------------------------------------------------

    transaction_article_col = header_positions["Article"]
    transaction_order_col = header_positions["Order No."]
    transaction_ordered_col = header_positions["Ordered"]
    transaction_booked_col = header_positions["Booked QTY"]
    transaction_variance_col = header_positions["Variance QTY"]

    # -------------------------------------------------------------------------
    # FIND TRANSACTION TABLE LAST ROW
    # -------------------------------------------------------------------------
    #
    # We use supplier_df length because it represents the transaction records
    # that were written to the supplier sheet.
    #
    # -------------------------------------------------------------------------

    transaction_first_data_row = transaction_header_row + 1

    transaction_last_data_row = (
        transaction_first_data_row + len(supplier_df) - 1
    )

    if transaction_last_data_row < transaction_first_data_row:
        transaction_last_data_row = transaction_first_data_row

    # -------------------------------------------------------------------------
    # HELPER TO CREATE EXCEL COLUMN LETTER
    # -------------------------------------------------------------------------

    from openpyxl.utils import get_column_letter

    transaction_article_letter = get_column_letter(
        transaction_article_col
    )

    transaction_order_letter = get_column_letter(
        transaction_order_col
    )

    transaction_ordered_letter = get_column_letter(
        transaction_ordered_col
    )

    transaction_booked_letter = get_column_letter(
        transaction_booked_col
    )

    transaction_variance_letter = get_column_letter(
        transaction_variance_col
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

        cell = sheet.cell(
            header_row,
            col
        )

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
    #
    # The DataFrame is used only to determine:
    #
    # - Which articles exist
    # - Which transaction records belong to each article
    #
    # The visible Excel totals are NOT written from pandas.
    #
    # The hidden PO detail rows are Excel formulas.
    #
    # The visible Article Summary rows are also Excel formulas.
    #
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

        article_name = article["Article"]

        # ---------------------------------------------------------------------
        # ARTICLE SUMMARY ROW
        # ---------------------------------------------------------------------

        article_cell = sheet.cell(
            current_row,
            1
        )

        article_cell.value = article_name

        ordered_cell = sheet.cell(
            current_row,
            2
        )

        delivered_cell = sheet.cell(
            current_row,
            3
        )

        variance_cell = sheet.cell(
            current_row,
            4
        )

        frequency_cell = sheet.cell(
            current_row,
            5
        )

        # ---------------------------------------------------------------------
        # ARTICLE ROW FORMATTING
        # ---------------------------------------------------------------------

        for col in range(1, 6):

            cell = sheet.cell(
                current_row,
                col
            )

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

            sheet.cell(
                current_row,
                col
            ).alignment = Alignment(
                horizontal="right",
                vertical="center"
            )

        # Number formats
        ordered_cell.number_format = "#,##0.00"

        delivered_cell.number_format = "#,##0.00"

        variance_cell.number_format = "#,##0.00"

        frequency_cell.number_format = "0"

        sheet.row_dimensions[current_row].height = 20

        # Save ONLY article rows for charting
        summary_rows.append(current_row)

        current_row += 1

        # ---------------------------------------------------------------------
        # GET TRANSACTION ROWS FOR THIS ARTICLE
        # ---------------------------------------------------------------------
        #
        # We use pandas only to determine WHICH transaction records belong
        # to this article.
        #
        # The actual values in the hidden rows will come from Excel formulas
        # referencing the transaction table.
        #
        # ---------------------------------------------------------------------

        detail = (
            supplier_df[
                supplier_df["Article"] == article_name
            ]
            .sort_values(
                [
                    "Order Date",
                    "Order No."
                ]
            )
        )

        # ---------------------------------------------------------------------
        # HIDDEN PO DETAIL START ROW
        # ---------------------------------------------------------------------

        detail_start_row = current_row

        # ---------------------------------------------------------------------
        # WRITE HIDDEN PO DETAIL ROWS
        # ---------------------------------------------------------------------

        for _, order in detail.iterrows():

            # -------------------------------------------------------------
            # ORIGINAL TRANSACTION ROW
            # -------------------------------------------------------------
            #
            # Because the supplier_df order matches the transaction table
            # order, find the corresponding transaction row by locating the
            # matching record.
            #
            # We use the Order No. and Article to identify the transaction
            # record.
            #
            # -------------------------------------------------------------

            matching_rows = supplier_df[
                (supplier_df["Article"] == order["Article"]) &
                (supplier_df["Order No."] == order["Order No."])
            ]

            if matching_rows.empty:

                raise ValueError(
                    f"Could not find transaction row for "
                    f"Article '{article_name}' and "
                    f"Order No. '{order['Order No.']}'."
                )

            # -------------------------------------------------------------
            # FIND EXACT TRANSACTION ROW
            # -------------------------------------------------------------
            #
            # If the same Article + Order No. occurs more than once,
            # use occurrence matching so that each hidden detail row
            # points to the corresponding transaction record.
            #
            # -------------------------------------------------------------

            occurrence_index = (
                supplier_df[
                    (supplier_df["Article"] == order["Article"]) &
                    (supplier_df["Order No."] == order["Order No."])
                ]
                .index
                .tolist()
            )

            original_index = order.name

            if original_index not in occurrence_index:

                raise ValueError(
                    f"Could not match transaction record for "
                    f"Article '{article_name}' and "
                    f"Order No. '{order['Order No.']}'."
                )

            # Convert pandas index to worksheet row.
            #
            # supplier_df is assumed to have been written to the transaction
            # table in its current row order.

            try:

                dataframe_position = (
                    supplier_df.index.get_loc(original_index)
                )

                transaction_row = (
                    transaction_first_data_row
                    + dataframe_position
                )

            except Exception:

                raise ValueError(
                    f"Could not determine transaction worksheet row "
                    f"for Article '{article_name}' and "
                    f"Order No. '{order['Order No.']}'."
                )

            # -------------------------------------------------------------
            # ORDER NUMBER
            # -------------------------------------------------------------

            detail_article = sheet.cell(
                current_row,
                1
            )

            detail_article.value = (
                f"={transaction_order_letter}{transaction_row}"
            )

            # -------------------------------------------------------------
            # ORDERED QTY
            # -------------------------------------------------------------
            #
            # LIVE FORMULA REFERENCING TRANSACTION TABLE
            # -------------------------------------------------------------

            detail_ordered = sheet.cell(
                current_row,
                2
            )

            detail_ordered.value = (
                f"={transaction_ordered_letter}{transaction_row}"
            )

            # -------------------------------------------------------------
            # DELIVERED QTY
            # -------------------------------------------------------------
            #
            # LIVE FORMULA REFERENCING TRANSACTION TABLE
            # -------------------------------------------------------------

            detail_delivered = sheet.cell(
                current_row,
                3
            )

            detail_delivered.value = (
                f"={transaction_booked_letter}{transaction_row}"
            )

            # -------------------------------------------------------------
            # QTY VARIANCE
            # -------------------------------------------------------------
            #
            # LIVE FORMULA REFERENCING TRANSACTION TABLE
            # -------------------------------------------------------------

            detail_variance = sheet.cell(
                current_row,
                4
            )

            detail_variance.value = (
                f"={transaction_variance_letter}{transaction_row}"
            )

            # No. of Orders column intentionally left blank
            sheet.cell(
                current_row,
                5
            ).value = None

            # -------------------------------------------------------------
            # DETAIL ROW FORMATTING
            # -------------------------------------------------------------

            for col in range(1, 6):

                cell = sheet.cell(
                    current_row,
                    col
                )

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

                sheet.cell(
                    current_row,
                    col
                ).alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            # Number formats
            detail_ordered.number_format = "#,##0.00"

            detail_delivered.number_format = "#,##0.00"

            detail_variance.number_format = "#,##0.00"

            # -------------------------------------------------------------
            # MAKE ORDER ROWS COLLAPSIBLE
            # -------------------------------------------------------------

            sheet.row_dimensions[
                current_row
            ].outlineLevel = 1

            sheet.row_dimensions[
                current_row
            ].hidden = True

            sheet.row_dimensions[
                current_row
            ].height = 18

            current_row += 1

        # ---------------------------------------------------------------------
        # REMEMBER FINAL HIDDEN PO DETAIL ROW
        # ---------------------------------------------------------------------

        detail_end_row = current_row - 1

        # ---------------------------------------------------------------------
        # LIVE FORMULAS FOR ARTICLE SUMMARY
        # ---------------------------------------------------------------------
        #
        # Monthly Article Summary:
        #
        # A = Article / Order Number
        # B = Ordered Qty
        # C = Delivered Qty
        # D = Qty Variance
        # E = No. of Orders
        #
        # Hidden detail rows:
        #
        # A = Order Number
        # B = Ordered Qty
        # C = Delivered Qty
        # D = Qty Variance
        #
        # ---------------------------------------------------------------------

        if detail_end_row >= detail_start_row:

            # -----------------------------------------------------------------
            # ORDERED QUANTITY
            # -----------------------------------------------------------------
            #
            # Sum Column B of hidden detail rows.
            #
            # -----------------------------------------------------------------

            ordered_cell.value = (
                f"=SUM("
                f"B{detail_start_row}:B{detail_end_row}"
                f")"
            )

            # -----------------------------------------------------------------
            # DELIVERED QUANTITY
            # -----------------------------------------------------------------
            #
            # Sum Column C of hidden detail rows.
            #
            # -----------------------------------------------------------------

            delivered_cell.value = (
                f"=SUM("
                f"C{detail_start_row}:C{detail_end_row}"
                f")"
            )

            # -----------------------------------------------------------------
            # QUANTITY VARIANCE
            # -----------------------------------------------------------------
            #
            # Ordered - Delivered
            #
            # Short delivery → Positive
            # Full delivery  → Zero
            # Over-delivery  → Negative
            #
            # -----------------------------------------------------------------

            variance_cell.value = (
                f"=B{article_row}-C{article_row}"
            )

            # -----------------------------------------------------------------
            # NO. OF ORDERS
            # -----------------------------------------------------------------
            #
            # Count UNIQUE Order Numbers from Column A of the hidden
            # PO detail rows.
            #
            # Same PO appearing multiple times for the same article
            # is counted only once.
            #
            # Example:
            #
            # TML202607-07092
            # TML202607-07491
            # TML202607-07491
            #
            # Result = 2
            #
            # -----------------------------------------------------------------

            frequency_cell.value = (
                f'=SUMPRODUCT(('
                f'A{detail_start_row}:A{detail_end_row}<>"")/'
                f'COUNTIF('
                f'A{detail_start_row}:A{detail_end_row},'
                f'A{detail_start_row}:A{detail_end_row}&""'
                f'))'
            )

        else:

            # -----------------------------------------------------------------
            # SAFETY FALLBACK
            # -----------------------------------------------------------------

            ordered_cell.value = "=0"

            delivered_cell.value = "=0"

            variance_cell.value = (
                f"=B{article_row}-C{article_row}"
            )

            frequency_cell.value = "=0"

        # ---------------------------------------------------------------------
        # COLLAPSE DETAILS UNDER ARTICLE
        # ---------------------------------------------------------------------

        sheet.row_dimensions[
            article_row
        ].collapsed = True

    # -------------------------------------------------------------------------
    # COLUMN WIDTHS
    # -------------------------------------------------------------------------

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
# ORDER NO. SUMMARY
###############################################################################

def create_order_summary(
    sheet,
    supplier_df,
    start_row,
    transaction_first_data_row,
    transaction_last_data_row
):
    """
    Creates a visible Order No. Summary with expandable article details.

    Structure:

        Order No. | Ordered Qty | Delivered Qty | Qty Variance | Delivery Days
            Article detail
            Article detail

        Order No. | Ordered Qty | Delivered Qty | Qty Variance | Delivery Days
            Article detail
            Article detail

    Main Order No. rows use LIVE EXCEL FORMULAS referencing
    the supplier transaction table.

    Article detail rows also use LIVE EXCEL FORMULAS referencing
    the supplier transaction table.

    Returns
    -------
    title_row : int
        Row containing the Order No. Summary title.

    header_row : int
        Row containing the Order No. Summary headers.

    order_summary_first_data_row : int
        First main Order No. data row.

    order_summary_last_data_row : int
        Last main Order No. data row.
    """

    from openpyxl.styles import (
        Font,
        PatternFill,
        Border,
        Side,
        Alignment
    )

    from openpyxl.utils import get_column_letter

    # =========================================================================
    # STYLES
    # =========================================================================

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

    title_font = Font(
        color=HEADER_FONT,
        bold=True,
        size=12
    )

    header_font = Font(
        bold=True,
        size=10
    )

    main_font = Font(
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

    # =========================================================================
    # FIND TRANSACTION COLUMNS
    # =========================================================================

    transaction_headers = {
        str(sheet.cell(1, col).value).strip(): col
        for col in range(
            1,
            transaction_last_data_row * 0 + sheet.max_column + 1
        )
    }

    transaction_order_col = transaction_headers.get(
        "Order No."
    )

    transaction_article_col = transaction_headers.get(
        "Article"
    )

    transaction_ordered_col = transaction_headers.get(
        "Ordered"
    )

    transaction_booked_col = transaction_headers.get(
        "Booked QTY"
    )

    transaction_variance_col = transaction_headers.get(
        "Variance QTY"
    )

    transaction_order_date_col = transaction_headers.get(
        "Order Date"
    )

    transaction_delivery_date_col = transaction_headers.get(
        "Delivery Date"
    )

    transaction_delivery_days_col = transaction_headers.get(
        "Delivery Days"
    )

    required_columns = [
        transaction_order_col,
        transaction_article_col,
        transaction_ordered_col,
        transaction_booked_col,
        transaction_variance_col,
        transaction_order_date_col,
        transaction_delivery_date_col,
        transaction_delivery_days_col
    ]

    if any(
        column is None
        for column in required_columns
    ):
        raise ValueError(
            "Could not find all required transaction columns "
            "for the Order No. Summary."
        )

    # =========================================================================
    # TRANSACTION COLUMN LETTERS
    # =========================================================================

    transaction_order_letter = get_column_letter(
        transaction_order_col
    )

    transaction_article_letter = get_column_letter(
        transaction_article_col
    )

    transaction_ordered_letter = get_column_letter(
        transaction_ordered_col
    )

    transaction_booked_letter = get_column_letter(
        transaction_booked_col
    )

    transaction_variance_letter = get_column_letter(
        transaction_variance_col
    )

    transaction_order_date_letter = get_column_letter(
        transaction_order_date_col
    )

    transaction_delivery_date_letter = get_column_letter(
        transaction_delivery_date_col
    )

    transaction_delivery_days_letter = get_column_letter(
         transaction_delivery_days_col
    )
        

    # =========================================================================
    # VALID PURCHASE ORDERS
    # =========================================================================
    #
    # IMPORTANT:
    # Convert the GroupBy object to an explicit list.
    #
    # This makes the number of unique order groups predictable and allows
    # len(valid_orders) to be used safely later.
    #
    # =========================================================================

    valid_order_df = (
        supplier_df[
            supplier_df["Order No."].notna()
            & ~supplier_df["Order No."].astype(str).str.strip().isin(
                [
                    "",
                    "NO PO DEFINED",
                    "N/A",
                    "NONE"
                ]
            )
        ]
    )

    valid_orders = list(
        valid_order_df.groupby(
            "Order No.",
            sort=True
        )
    )

    # =========================================================================
    # TITLE
    # =========================================================================

    title_row = start_row

    sheet.merge_cells(
        start_row=title_row,
        start_column=1,
        end_row=title_row,
        end_column=5
    )

    title_cell = sheet.cell(
        title_row,
        1
    )

    title_cell.value = "Order No. Summary"
    title_cell.fill = title_fill
    title_cell.font = title_font

    title_cell.alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    title_cell.border = table_border

    for col in range(1, 6):

        cell = sheet.cell(
            title_row,
            col
        )

        cell.fill = title_fill
        cell.border = table_border

    sheet.row_dimensions[
        title_row
    ].height = 24

    # =========================================================================
    # SPACE BETWEEN TITLE AND HEADER
    # =========================================================================

    header_row = title_row + 2

    # =========================================================================
    # HEADERS
    # =========================================================================

    headers = [
        "Order No.",
        "Ordered Qty",
        "Delivered Qty",
        "Qty Variance",
        "Delivery Days"
    ]

    for col, header in enumerate(
        headers,
        start=1
    ):

        cell = sheet.cell(
            header_row,
            col
        )

        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = table_border

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    sheet.row_dimensions[
        header_row
    ].height = 21

    # =========================================================================
    # FIRST MAIN ORDER SUMMARY ROW
    # =========================================================================

    current_row = header_row + 1

    order_summary_first_data_row = current_row

    # =========================================================================
    # WRITE ONE MAIN ROW PER UNIQUE ORDER
    # =========================================================================

    for order_no, order_group in valid_orders:

        # ---------------------------------------------------------------------
        # MAIN ORDER ROW
        # ---------------------------------------------------------------------

        order_row = current_row

        # ---------------------------------------------------------------------
        # ORDER NUMBER
        # ---------------------------------------------------------------------

        order_cell = sheet.cell(
            order_row,
            1
        )

        order_cell.value = order_no
        order_cell.border = table_border
        order_cell.font = main_font

        order_cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        # ---------------------------------------------------------------------
        # ORDERED QTY
        #
        # LIVE FORMULA:
        # Sum all Ordered quantities for this Purchase Order.
        # ---------------------------------------------------------------------

        ordered_cell = sheet.cell(
            order_row,
            2
        )

        ordered_cell.value = (
            f'=SUMIF('
            f'${transaction_order_letter}${transaction_first_data_row}:'
            f'${transaction_order_letter}${transaction_last_data_row},'
            f'A{order_row},'
            f'${transaction_ordered_letter}${transaction_first_data_row}:'
            f'${transaction_ordered_letter}${transaction_last_data_row}'
            f')'
        )

        ordered_cell.number_format = "#,##0.00"
        ordered_cell.border = table_border
        ordered_cell.font = main_font

        ordered_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        # ---------------------------------------------------------------------
        # DELIVERED QTY
        #
        # LIVE FORMULA:
        # Sum all Booked QTY values for this Purchase Order.
        # ---------------------------------------------------------------------

        delivered_cell = sheet.cell(
            order_row,
            3
        )

        delivered_cell.value = (
            f'=SUMIF('
            f'${transaction_order_letter}${transaction_first_data_row}:'
            f'${transaction_order_letter}${transaction_last_data_row},'
            f'A{order_row},'
            f'${transaction_booked_letter}${transaction_first_data_row}:'
            f'${transaction_booked_letter}${transaction_last_data_row}'
            f')'
        )

        delivered_cell.number_format = "#,##0.00"
        delivered_cell.border = table_border
        delivered_cell.font = main_font

        delivered_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        # ---------------------------------------------------------------------
        # QTY VARIANCE
        #
        # Ordered - Delivered
        #
        # Positive = short delivery
        # Zero     = full delivery
        # Negative = over-delivery
        # ---------------------------------------------------------------------

        variance_cell = sheet.cell(
            order_row,
            4
        )

        variance_cell.value = (
            f"=B{order_row}-C{order_row}"
        )

        variance_cell.number_format = "#,##0.00"
        variance_cell.border = table_border
        variance_cell.font = main_font

        variance_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        # ---------------------------------------------------------------------
        # DELIVERY DAYS
        #
        # LIVE FORMULA:
        #
        # Latest Delivery Date
        # minus
        # Earliest Order Date
        #
        # for this Purchase Order.
        # ---------------------------------------------------------------------

        delivery_days_cell = sheet.cell(
            order_row,
            5
        )

        delivery_days_cell.value = (
            f'=IFERROR('
            f'MAXIFS('
            f'${transaction_delivery_days_letter}${transaction_first_data_row}:'
            f'${transaction_delivery_days_letter}${transaction_last_data_row},'
            f'${transaction_order_letter}${transaction_first_data_row}:'
            f'${transaction_order_letter}${transaction_last_data_row},'
            f'A{order_row}'
            f'),'
            f'0'
            f')'
        )

        delivery_days_cell.number_format = "0.0"
        delivery_days_cell.border = table_border
        delivery_days_cell.font = main_font

        delivery_days_cell.alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        # ---------------------------------------------------------------------
        # MAIN ORDER ROW HEIGHT
        # ---------------------------------------------------------------------

        sheet.row_dimensions[
            order_row
        ].height = 20

        # ---------------------------------------------------------------------
        # ARTICLE DETAIL ROWS
        # ---------------------------------------------------------------------
        #
        # Every transaction belonging to this Purchase Order gets one
        # expandable detail row underneath the main Order No. row.
        #
        # These rows reference the transaction table directly.
        #
        # ---------------------------------------------------------------------

        for _, transaction in order_group.iterrows():

            # -----------------------------------------------------------------
            # FIND ORIGINAL TRANSACTION ROW
            # -----------------------------------------------------------------

            original_index = transaction.name

            try:

                dataframe_position = (
                    supplier_df.index.get_loc(
                        original_index
                    )
                )

            except Exception as exc:

                raise ValueError(
                    f"Could not determine transaction worksheet row "
                    f"for Order No. '{order_no}'."
                ) from exc

            transaction_row = (
                transaction_first_data_row
                + dataframe_position
            )

            # -----------------------------------------------------------------
            # MOVE TO NEXT DETAIL ROW
            # -----------------------------------------------------------------

            detail_row = current_row + 1

            # -----------------------------------------------------------------
            # ARTICLE
            # -----------------------------------------------------------------

            detail_article = sheet.cell(
                detail_row,
                1
            )

            detail_article.value = (
                f"={transaction_article_letter}{transaction_row}"
            )

            # -----------------------------------------------------------------
            # ORDERED QTY
            # -----------------------------------------------------------------

            detail_ordered = sheet.cell(
                detail_row,
                2
            )

            detail_ordered.value = (
                f"={transaction_ordered_letter}{transaction_row}"
            )

            # -----------------------------------------------------------------
            # DELIVERED QTY
            # -----------------------------------------------------------------

            detail_delivered = sheet.cell(
                detail_row,
                3
            )

            detail_delivered.value = (
                f"={transaction_booked_letter}{transaction_row}"
            )

            # -----------------------------------------------------------------
            # QTY VARIANCE
            # -----------------------------------------------------------------

            detail_variance = sheet.cell(
                detail_row,
                4
            )

            detail_variance.value = (
                f"={transaction_variance_letter}{transaction_row}"
            )

            # -----------------------------------------------------------------
            # DELIVERY DAYS
            #
            # Delivery Days is shown only on the main Order No. row.
            # -----------------------------------------------------------------

            detail_delivery = sheet.cell(
                detail_row,
                5
            )

            detail_delivery.value = None

            # -----------------------------------------------------------------
            # DETAIL ROW FORMATTING
            # -----------------------------------------------------------------

            for col in range(1, 6):

                cell = sheet.cell(
                    detail_row,
                    col
                )

                cell.fill = detail_fill
                cell.border = table_border
                cell.font = detail_font

                cell.alignment = Alignment(
                    vertical="center"
                )

            # -----------------------------------------------------------------
            # ARTICLE ALIGNMENT
            # -----------------------------------------------------------------

            detail_article.alignment = Alignment(
                horizontal="left",
                vertical="center",
                indent=1
            )

            # -----------------------------------------------------------------
            # NUMERIC ALIGNMENT
            # -----------------------------------------------------------------

            for col in range(2, 5):

                sheet.cell(
                    detail_row,
                    col
                ).alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            # -----------------------------------------------------------------
            # NUMBER FORMATS
            # -----------------------------------------------------------------

            detail_ordered.number_format = "#,##0.00"

            detail_delivered.number_format = "#,##0.00"

            detail_variance.number_format = "#,##0.00"

            # -----------------------------------------------------------------
            # DETAIL ROW HEIGHT
            # -----------------------------------------------------------------

            sheet.row_dimensions[
                detail_row
            ].height = 18

            # -----------------------------------------------------------------
            # MAKE DETAIL ROW COLLAPSIBLE
            # -----------------------------------------------------------------

            sheet.row_dimensions[
                detail_row
            ].outlineLevel = 1

            sheet.row_dimensions[
                detail_row
            ].hidden = True

            # -----------------------------------------------------------------
            # UPDATE CURRENT ROW
            # -----------------------------------------------------------------

            current_row = detail_row

        # ---------------------------------------------------------------------
        # COLLAPSE ALL DETAILS UNDER THIS ORDER
        # ---------------------------------------------------------------------

        sheet.row_dimensions[
            order_row
        ].collapsed = True

        # ---------------------------------------------------------------------
        # MOVE TO NEXT MAIN ORDER ROW
        # ---------------------------------------------------------------------

        current_row += 1

    # =========================================================================
    # LAST MAIN ORDER SUMMARY ROW
    # =========================================================================

    if valid_orders:

        order_summary_last_data_row = (
            order_summary_first_data_row
            + len(valid_orders)
            - 1
        )

    else:

        order_summary_last_data_row = (
            order_summary_first_data_row - 1
        )

    # =========================================================================
    # COLUMN WIDTHS
    # =========================================================================

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

    # =========================================================================
    # RETURN
    # =========================================================================

    return (
        title_row,
        header_row,
        order_summary_first_data_row,
        order_summary_last_data_row
    )
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
        # =========================================================
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
        # row, KPI panel, Order No. Summary, article summary, or
        # charts as transaction data.
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
        # THREE BLANK ROWS AFTER TOTAL
        # =========================================================
        #
        # Layout:
        #
        # Transaction Data
        # TOTAL
        # blank
        # blank
        # blank
        # Supplier KPI Summary
        #
        # =========================================================

        start_row = total_row + 4

        # =========================================================
        # KPI PANEL STYLES
        # =========================================================

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
            # ORDERED QTY
            # =====================================================

            if kpi == "Ordered Qty":

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
            # The KPI explicitly uses the monetary
            # "Variance Value" column.
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
            # ORDERS / AVERAGE DELIVERY DAYS
            #
            # These are deliberately left blank for now.
            #
            # They will be populated AFTER create_order_summary()
            # because the Order No. Summary is created below the
            # KPI panel.
            # =====================================================

            elif kpi == "Orders":

                value_cell.value = None
                value_cell.number_format = "0"

            elif kpi == "Average Delivery Days":

                value_cell.value = None
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
        # ORDER NO. SUMMARY
        # =========================================================
        #
        # Layout:
        #
        # Supplier KPI Summary
        #
        # [3 blank rows]
        #
        # Order No. Summary
        # Order No. | Ordered Qty | Delivered Qty |
        # Qty Variance | Delivery Days
        #
        # PO 1
        #     article detail
        #     article detail
        #
        # PO 2
        #     article detail
        #
        # =========================================================

        order_summary_start = (
            start_row
            + len(kpi_labels)
            + 5
        )

        (
            order_summary_title_row,
            order_summary_header_row,
            order_summary_first_data_row,
            order_summary_last_data_row
        ) = create_order_summary(
            sheet,
            supplier_df,
            order_summary_start,
            2,
            last_data_row
        )

        # =========================================================
        # UPDATE KPI FORMULAS USING ORDER NO. SUMMARY
        # =========================================================
        #
        # The Order No. Summary contains:
        #
        # Main Order rows = visible
        # Article detail rows = hidden
        #
        # We therefore use SUBTOTAL so Excel counts/averages only
        # the visible main Order No. rows and ignores the hidden
        # article detail rows.
        # =========================================================

        # ---------------------------------------------------------
        # ORDERS KPI
        # ---------------------------------------------------------

        orders_kpi_row = start_row + 2

        if order_summary_last_data_row >= order_summary_first_data_row:

            sheet.cell(
                orders_kpi_row,
                2
            ).value = (
                f"=COUNTA("
                f"A{order_summary_first_data_row}:"
                f"A{order_summary_last_data_row}"
                f")"
            )

        else:

            sheet.cell(
                orders_kpi_row,
                2
            ).value = 0

        sheet.cell(
            orders_kpi_row,
            2
        ).number_format = "0"

        # ---------------------------------------------------------
        # AVERAGE DELIVERY DAYS KPI
        # ---------------------------------------------------------

        average_delivery_kpi_row = start_row + 8

        if order_summary_last_data_row >= order_summary_first_data_row:

            sheet.cell(
                average_delivery_kpi_row,
                2
            ).value = (
                f"=IFERROR("
                f"AVERAGE("
                f"E{order_summary_first_data_row}:"
                f"E{order_summary_last_data_row}"
                f"),"
                f"0"
                f")"
            )

        else:

            sheet.cell(
                average_delivery_kpi_row,
                2
            ).value = 0

        sheet.cell(
            average_delivery_kpi_row,
            2
        ).number_format = "0.0"

        # =========================================================
        # MONTHLY ARTICLE SUMMARY
        # =========================================================
        #
        # The Order No. Summary contains both main Order rows and
        # hidden article detail rows.
        #
        # Therefore we use sheet.max_row AFTER create_order_summary()
        # so the Monthly Article Summary starts below the entire
        # Order No. Summary.
        #
        # Three blank rows are left between the sections.
        # =========================================================

        summary_start = sheet.max_row + 4

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
    Apply professional formatting to the supplier worksheet.

    IMPORTANT
    ---------
    Only the transaction table is treated as the transaction table.

    Transaction table:
        Row 1                  = headers
        Rows 2:last_data_row   = transaction data
        Columns A:O            = transaction columns

    The following sections are NOT included in transaction formatting:
        - TOTAL row
        - Supplier KPI Summary
        - Order No. Summary
        - Monthly Article Summary
        - Charts

    This prevents the lower report sections from being affected by
    transaction-table filters, number formatting, or row formatting.
    """

    # =========================================================================
    # STYLES
    # =========================================================================

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=HEADER_FILL
    )

    header_font = Font(
        bold=True,
        color=HEADER_FONT
    )

    thin = Side(
        style="thin"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # =========================================================================
    # TRANSACTION TABLE SETTINGS
    # =========================================================================

    # Keep the existing configured freeze pane.
    # This freezes the transaction header row.
    ws.freeze_panes = FREEZE_PANES

    # IMPORTANT:
    # The transaction table is A:O.
    #
    # Do NOT use ws.max_column here because the worksheet now contains
    # additional report sections below the transaction table.
    ws.auto_filter.ref = (
        f"A1:O{last_data_row}"
    )

    # =========================================================================
    # TRANSACTION TABLE HEADER
    # =========================================================================

    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT

    for col in range(1, 16):

        cell = ws.cell(
            1,
            col
        )

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    # =========================================================================
    # STORE TRANSACTION HEADERS
    # =========================================================================

    headers = {
        cell.column: str(cell.value)
        for cell in ws[1]
        if cell.column <= 15
    }

    # =========================================================================
    # FORMAT ONLY TRANSACTION DATA
    # =========================================================================

    for row in ws.iter_rows(
        min_row=2,
        max_row=last_data_row,
        min_col=1,
        max_col=15
    ):

        ws.row_dimensions[
            row[0].row
        ].height = DEFAULT_ROW_HEIGHT

        for cell in row:

            cell.border = border

            header = headers.get(
                cell.column,
                ""
            )

            # -----------------------------------------------------------------
            # DATES
            # -----------------------------------------------------------------

            if "Date" in header:

                cell.number_format = "dd-mmm-yyyy"

            # -----------------------------------------------------------------
            # PERCENTAGES
            # -----------------------------------------------------------------

            elif "%" in header:

                cell.number_format = "0.00%"

            # -----------------------------------------------------------------
            # NUMBERS
            # -----------------------------------------------------------------

            elif isinstance(
                cell.value,
                (int, float)
            ):

                cell.number_format = "#,##0.00"

    # =========================================================================
    # AUTO-FIT TRANSACTION COLUMNS ONLY
    # =========================================================================
    #
    # IMPORTANT:
    # Do NOT use:
    #
    #     for column in ws.columns:
    #
    # because ws.columns now includes the Order No. Summary and other
    # sections.
    #
    # We only auto-fit A:O using the transaction table.
    #
    # Existing widths created by the summary functions are preserved if
    # they are already wider.
    # =========================================================================

    for col_num in range(1, 16):

        column_letter = get_column_letter(
            col_num
        )

        max_length = 0

        # Header length
        header_value = ws.cell(
            1,
            col_num
        ).value

        if header_value is not None:

            max_length = len(
                str(header_value)
            )

        # Transaction data only
        for row_num in range(
            2,
            last_data_row + 1
        ):

            value = ws.cell(
                row_num,
                col_num
            ).value

            if value is not None:

                max_length = max(
                    max_length,
                    len(str(value))
                )

        calculated_width = min(
            max_length + 3,
            40
        )

        # Never reduce a width already established by
        # create_order_summary() or create_article_summary().
        existing_width = (
            ws.column_dimensions[
                column_letter
            ].width
            or 0
        )

        ws.column_dimensions[
            column_letter
        ].width = max(
            existing_width,
            calculated_width
        )


###############################################################################
# CONDITIONAL FORMATTING
###############################################################################

def apply_conditional_formatting(ws, last_data_row):
    """
    Apply conditional formatting ONLY to the transaction table.

    Delivery Days colour scale is restricted to:
        Row 2:last_data_row

    This prevents the conditional formatting from affecting:
        - TOTAL row
        - Supplier KPI Summary
        - Order No. Summary
        - Monthly Article Summary
        - Chart Summary
    """

    # =========================================================================
    # FIND TRANSACTION HEADERS
    # =========================================================================

    headers = {
        cell.value: cell.column
        for cell in ws[1]
        if cell.column <= 15
    }

    # =========================================================================
    # DELIVERY DAYS
    # =========================================================================

    if "Delivery Days" in headers:

        col = get_column_letter(
            headers["Delivery Days"]
        )

        # ---------------------------------------------------------------------
        # Apply colour scale ONLY to transaction data.
        # ---------------------------------------------------------------------

        if last_data_row >= 2:

            ws.conditional_formatting.add(

                f"{col}2:{col}{last_data_row}",

                ColorScaleRule(

                    start_type="min",
                    start_color="63BE7B",

                    mid_type="percentile",
                    mid_value=50,
                    mid_color="FFEB84",

                    end_type="max",
                    end_color="F8696A"
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

            last_data_row = (
                worksheet_last_rows[sheet.title]
            )
                
            format_worksheet(
                sheet,
                worksheet_last_rows[sheet.title]
            )

            apply_conditional_formatting(sheet,last_data_row)
    # ---------------------------------------------------------
    # Return workbook and report period
    # ---------------------------------------------------------

    return save_workbook(workbook), report_period
