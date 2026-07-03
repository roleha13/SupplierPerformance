"""
processor.py
Section 1
-------------------------------
Imports
Validation
Reading Excel Files
Cleaning Data
Merging Data
Calculating Delivery Days
"""

import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

from config import (
    REPORT_COLUMNS,
    REGISTER_REQUIRED_COLUMNS,
    PURCHASE_REGISTER_COLUMNS,
    PURCHASE_RECEIVING_COLUMNS,
    EXCLUDED_SUPPLIERS,
    SUMMARY_COLUMNS,
    REPORT_TITLE,
    MASTER_SHEET,
    OUTPUT_FILE,
    HEADER_FILL,
    HEADER_FONT,
    TOTAL_FILL,
    FREEZE_PANES,
    HEADER_ROW_HEIGHT,
    DEFAULT_ROW_HEIGHT,
)

###############################################################################
# Helper Functions
###############################################################################

def validate_columns(df: pd.DataFrame, required_columns: set, file_name: str):
    """
    Validate uploaded workbook columns.
    Raises ValueError if any required column is missing.
    """

    missing = required_columns.difference(df.columns)

    if missing:
        raise ValueError(
            f"\n{file_name}\n"
            f"Missing columns:\n"
            f"{', '.join(sorted(missing))}"
        )


###############################################################################
# Read Excel Files
###############################################################################

def read_purchase_register(file_path):
    """
    Read Purchase Register workbook.
    """

    df = pd.read_excel(file_path)

    validate_columns(
        df,
        PURCHASE_REGISTER_COLUMNS,
        "Purchase Register"
    )

    df = df[REGISTER_REQUIRED_COLUMNS].copy()

    df["Order No."] = (
        df["Order No."]
        .astype(str)
        .str.strip()
    )

    df["Order Date"] = pd.to_datetime(
        df["Order Date"],
        errors="coerce"
    )

    return df


def read_receiving_report(file_path):
    """
    Read Purchase Receiving Deviation workbook.
    """

    df = pd.read_excel(file_path)

    validate_columns(
        df,
        PURCHASE_RECEIVING_COLUMNS,
        "Purchase Receiving Deviation"
    )

    return df


###############################################################################
# Cleaning
###############################################################################

def clean_receiving_data(df):
    """
    Keep required columns only.
    """

    keep = [

        "Supplier",
        "Article",
        "Order No.",
        "Ordered",
        "Order Unit",
        "Booked QTY",
        "Variance QTY",
        "PO Price",
        "Booked Price",
        "Variance Price",
        "Variance Value",
        "Delivery Date"

    ]

    df = df[keep].copy()

    df["Supplier"] = (
        df["Supplier"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Order No."] = (
        df["Order No."]
        .astype(str)
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


###############################################################################
# Remove Excluded Suppliers
###############################################################################

def remove_excluded_suppliers(df):
    """
    Remove suppliers defined in config.py
    """

    df["Supplier"] = (
        df["Supplier"]
        .str.upper()
        .str.strip()
    )

    return df[
        ~df["Supplier"].isin(EXCLUDED_SUPPLIERS)
    ].copy()


###############################################################################
# Merge Purchase Register
###############################################################################

def merge_order_dates(receiving_df, register_df):
    """
    Bring Order Date into receiving report
    using Order Number.
    """

    merged = receiving_df.merge(

        register_df,

        on="Order No.",

        how="left"

    )

    return merged


###############################################################################
# Delivery Days KPI
###############################################################################

def calculate_delivery_days(df):
    """
    Delivery Days =
    Delivery Date - Order Date
    """

    df["Delivery Days"] = (

        df["Delivery Date"] -

        df["Order Date"]

    ).dt.days

    return df


###############################################################################
# Final Report Dataset
###############################################################################

def prepare_report_data(register_file, receiving_file):
    """
    Complete preprocessing pipeline.
    """

    register = read_purchase_register(register_file)

    receiving = read_receiving_report(receiving_file)

    receiving = clean_receiving_data(receiving)

    receiving = remove_excluded_suppliers(receiving)

    merged = merge_order_dates(

        receiving,

        register

    )

    merged = calculate_delivery_days(merged)

    merged = merged[REPORT_COLUMNS]

    merged.sort_values(

        by=[

            "Supplier",

            "Order Date"

        ],

        inplace=True

    )

    merged.reset_index(

        drop=True,

        inplace=True

    )

    return merged

###############################################################################
# KPI CALCULATIONS
###############################################################################

def create_master_summary(df):
    """
    Generate supplier KPI summary.
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

    # Fill Rate %

    summary["Fill Rate %"] = (
        summary["Received_Qty"] /
        summary["Ordered_Qty"]
    ).replace([float("inf")], 0).fillna(0) * 100

    summary["Average_Delivery_Days"] = (
        summary["Average_Delivery_Days"]
        .round(1)
    )

    summary["Fill Rate %"] = (
        summary["Fill Rate %"]
        .round(2)
    )

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

    summary.sort_values(

        by=[

            "Fill Rate %",
            "Average Delivery Days"

        ],

        ascending=[False, True],

        inplace=True

    )

    summary.reset_index(

        drop=True,

        inplace=True

    )

    return summary


###############################################################################
# MASTER SUMMARY WORKSHEET
###############################################################################

def write_master_summary(workbook, summary_df):
    """
    Create the dashboard worksheet.
    """

    ws = workbook.create_sheet(MASTER_SHEET)

    ws.append(summary_df.columns.tolist())

    for row in summary_df.itertuples(index=False):

        ws.append(list(row))

    return ws


###############################################################################
# SUPPLIER WORKSHEETS
###############################################################################

def create_supplier_sheets(workbook, df):
    """
    One worksheet per supplier.
    """

    supplier_sheets = {}

    suppliers = sorted(

        df["Supplier"].unique()

    )

    for supplier in suppliers:

        sheet_name = supplier[:31]

        ws = workbook.create_sheet(sheet_name)

        supplier_df = (

            df[df["Supplier"] == supplier]

            .sort_values(

                by=[

                    "Order Date",

                    "Order No."

                ]

            )

        )

        ws.append(

            supplier_df.columns.tolist()

        )

        for row in supplier_df.itertuples(index=False):

            ws.append(list(row))

        supplier_sheets[supplier] = ws

    return supplier_sheets


###############################################################################
# BUILD REPORT WORKBOOK
###############################################################################

def build_workbook(report_df):
    """
    Create workbook and populate
    all worksheets.
    """

    wb = Workbook()

    default_sheet = wb.active

    wb.remove(default_sheet)

    summary = create_master_summary(report_df)

    write_master_summary(

        wb,

        summary

    )

    create_supplier_sheets(

        wb,

        report_df

    )

    return wb
