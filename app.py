"""
app.py
---------------------------------------------------------
Supplier Performance Report Tool
Streamlit Application
---------------------------------------------------------
"""

from pathlib import Path
import tempfile

import streamlit as st

from processor import process_files


# ==========================================================
# PAGE CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="Supplier Performance Report",
    page_icon="📊",
    layout="wide"
)


# ==========================================================
# SIDEBAR
# ==========================================================

with st.sidebar:

    logo_path = Path("assets/logo.png")

    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)

    st.markdown("## Supplier Performance Report")

    st.caption("Version 1.0")

    st.divider()

    st.markdown(
        """
        ### Required Files

        • Purchase Register.xlsx

        • Purchase Receiving Deviation per Supplier.xlsx
        """
    )

    st.divider()

    st.info(
        """
        Suppliers listed in **config.py**
        are automatically excluded
        from the report.
        """
    )


# ==========================================================
# HEADER
# ==========================================================

st.title("📊 Supplier Performance Report")

st.write(
    """
Generate a professional supplier performance report
from Purchase Register and Purchase Receiving Deviation
workbooks.

The report includes:

- Master Summary Dashboard
- Supplier KPI Sheets
- Delivery Days
- Fill Rate
- Quantity Variance
- Price Variance
- Charts
"""
)

st.divider()


# ==========================================================
# FILE UPLOADS
# ==========================================================

col1, col2 = st.columns(2)

with col1:

    purchase_register = st.file_uploader(

        "Purchase Register",

        type=["xlsx"]

    )

with col2:

    receiving_report = st.file_uploader(

        "Purchase Receiving Deviation per Supplier",

        type=["xlsx"]

    )


st.divider()


# ==========================================================
# GENERATE BUTTON
# ==========================================================

generate = st.button(

    "Generate Report",

    type="primary",

    use_container_width=True

)


# ==========================================================
# PROCESS REPORT
# ==========================================================

if generate:

    if purchase_register is None:

        st.error(
            "Please upload Purchase Register.xlsx"
        )

        st.stop()

    if receiving_report is None:

        st.error(
            "Please upload Purchase Receiving Deviation.xlsx"
        )

        st.stop()

    progress = st.progress(0)

    status = st.empty()

    try:

        with tempfile.TemporaryDirectory() as temp_dir:

            register_path = Path(temp_dir) / purchase_register.name

            receiving_path = Path(temp_dir) / receiving_report.name

            register_path.write_bytes(

                purchase_register.getbuffer()

            )

            receiving_path.write_bytes(

                receiving_report.getbuffer()

            )

            status.info("Reading Excel files...")

            progress.progress(15)

            status.info("Cleaning data...")

            progress.progress(30)

            status.info("Calculating KPIs...")

            progress.progress(50)

            status.info("Generating supplier worksheets...")

            progress.progress(70)

            status.info("Building dashboard...")

            progress.progress(85)

            excel_file = process_files(

                register_path,

                receiving_path,
            )

            progress.progress(100)

            status.success(

                "Report generated successfully."

            )

            with open(output_file, "rb") as file:

                st.download_button(

                    label="📥 Download Supplier Performance Report",

                    data=excel_file,

                    file_name="Supplier_Performance_Report.xlsx",

                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                    use_container_width=True

                )

    except Exception as e:

        st.error("Report generation failed.")

        st.exception(e)


# ==========================================================
# FOOTER
# ==========================================================

st.divider()

st.caption(
    """
Supplier Performance Report Tool | Version 1.0

Developed by Data Analytics & Innovation
"""
)
