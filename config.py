"""
config.py
----------
Configuration settings for the Supplier Performance Report Tool.
Modify this file whenever business rules change.
"""

# =============================================================================
# REPORT SETTINGS
# =============================================================================

REPORT_TITLE = "Supplier Performance Report"
OUTPUT_FILE = "Supplier_Performance_Report.xlsx"
MASTER_SHEET = "Master Summary"

# =============================================================================
# EXCEL SETTINGS
# =============================================================================

FREEZE_PANES = "A2"
HEADER_ROW_HEIGHT = 22
DEFAULT_ROW_HEIGHT = 18

# =============================================================================
# REPORT FORMATTING
# =============================================================================

HEADER_FILL = "1F4E78"      # Dark Blue
HEADER_FONT = "FFFFFF"      # White
TOTAL_FILL = "D9EAD3"       # Light Green

# =============================================================================
# REQUIRED COLUMNS - PURCHASE REGISTER
# =============================================================================

REGISTER_REQUIRED_COLUMNS = [
    "Order No.",
    "Order Date"
]

# =============================================================================
# REQUIRED COLUMNS - PURCHASE RECEIVING DEVIATION
# (Columns expected in the uploaded report)
# =============================================================================

RECEIVING_INPUT_COLUMNS = [

    "Supplier",
    "Article",
    "Ordered",
    "Order Unit",
    "Booked QTY",
    "Variance QTY",
    "PO Price",
    "Booked Price",
    "Variance Price",
    "Variance Value",
    "Order No.",
    "Delivery Date"

]

# =============================================================================
# FINAL REPORT COLUMNS
# (Columns that will appear in every supplier worksheet)
# =============================================================================

REPORT_COLUMNS = [

    "Supplier",
    "Article",
    "Order No.",
    "Order Date",
    "Delivery Date",
    "Delivery Days",

    "Ordered",
    "Order Unit",
    "Booked QTY",
    "Variance QTY",

    "PO Price",
    "Booked Price",
    "Variance Price",
    "Variance Value"

]

# =============================================================================
# MASTER SUMMARY COLUMNS
# =============================================================================

SUMMARY_COLUMNS = [

    "Supplier",
    "Orders",
    "Ordered Qty",
    "Received Qty",
    "Fill Rate %",
    "Qty Variance",
    "Price Variance",
    "Average Delivery Days"

]

# =============================================================================
# SUPPLIERS TO EXCLUDE
# =============================================================================

EXCLUDED_SUPPLIERS = {

    "BEYOND FRUITS LIMITED",
    "BIGCOLD KENYA LTD / (SIMPLIFINE)",
    "NAIVAS LIMITED",
    "GHANSHYAM FRUITS AND VEGETABLES LTD",
    "MGANDINI FRUITS & GREEN GROCERY",
    "BLESSINGS FRESH JUICES LIMITED",
    "RAHMA LISHERZ KAYANDA",
    "SUNDRY CREDITORS - MOMBASA AREA",
    "CHANDARANA SUPERMARKET",
    "CARREFOUR (MAJID AL FUTTAIM HYPERMARKET LIMITED) DAMA CHARO",
    "FENNY TEMBO",
    "URBAGDA FRUITS AND VEGETABLE ENTERPRISES",
    "REDCREST WINES LIMITED",
    "MORANI LIMITED",
    "MEGA WINES AND SPIRITS LTD",
    "LENFLO HOLDINGS LIMITED",
    "STABEX INTERNATIONAL LIMITED",
    "ALICE GIKANDI",
    "KEN COMP EXPERTS LIMITED",
    "FAHA FRESH WATER LTD",
    "BLUECOLLAR SUPPLY SOLUTIONS LIMITED",
    "JAMBO WOOD WORKS L.T.D"

}

# =============================================================================
# REQUIRED COLUMNS FOR VALIDATION
# =============================================================================

PURCHASE_REGISTER_COLUMNS = {

    "Order No.",
    "Order Date"

}

PURCHASE_RECEIVING_COLUMNS = {

    "Supplier",
    "Article",
    "Ordered",
    "Order Unit",
    "Booked QTY",
    "Variance QTY",
    "PO Price",
    "Booked Price",
    "Variance Price",
    "Variance Value",
    "Order No.",
    "Delivery Date"

}
