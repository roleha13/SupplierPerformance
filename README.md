# Supplier Performance Report Tool

## Overview

The Supplier Performance Report Tool is a web-based application developed using Streamlit and Python to automate supplier performance analysis.

The application combines data from:

- Purchase Register
- Purchase Receiving Deviation per Supplier

into a single professionally formatted Excel report containing supplier performance KPIs, charts, and management summaries.

---

## Features

- Supports Microsoft Excel 97–2003 (.xls)
- Supports Microsoft Excel Workbook (.xlsx)
- Supports Microsoft Excel Macro-Enabled Workbook (.xlsm)
- Automatically maps Order Date using Order Number
- Calculates Delivery Days
- Calculates Fill Rate (%)
- Calculates Quantity Variance
- Calculates Price Variance
- Generates one worksheet per supplier
- Generates a Master Summary dashboard
- Automatically excludes configured suppliers
- Creates professional charts
- Applies professional Excel formatting
- Runs entirely through a web browser using Streamlit

---

## Required Input Files

The application requires two reports exported from Materials Control.

### 1. Purchase Register

Required columns:

- Order No.
- Order Date

### 2. Purchase Receiving Deviation per Supplier

Contains supplier delivery information including:

- Supplier
- Article
- Ordered Quantity
- Booked Quantity
- Quantity Variance
- Price Variance
- Delivery Date
- Order Number

---

## Supported File Formats

The application automatically detects and reads:

- .xls
- .xlsx
- .xlsm

No user configuration is required.

---

## Output

The application generates:

Supplier_Performance_Report.xlsx

The report includes:

- Master Summary Dashboard
- One worksheet for every supplier
- Delivery Days KPI
- Fill Rate %
- Quantity Variance
- Price Variance
- KPI Summary Panel
- Charts for each supplier
- Professional formatting

---

## Business Rules

- Order Date is retrieved from the Purchase Register using Order Number.
- Delivery Days = Delivery Date − Order Date.
- Fill Rate = Received Quantity ÷ Ordered Quantity × 100.
- Excluded suppliers are maintained in config.py.
- One worksheet is generated for every included supplier.

---

## Project Structure

```
SupplierPerformanceTool/

│

├── app.py

├── processor.py

├── config.py

├── requirements.txt

├── README.md

│

└── assets/

    └── logo.png
```

---

## Installation

Clone the repository.

Install the required packages.

```
pip install -r requirements.txt
```

Run the application.

```
streamlit run app.py
```

---

## Technologies Used

- Python
- Streamlit
- Pandas
- OpenPyXL
- xlrd

---

## Future Enhancements

Potential improvements include:

- PDF Report Export
- Email Report Distribution
- Multi-file Processing
- Report Generation History
- Supplier Trend Analysis
- User Authentication
- Company Branding Enhancements

---

## Version

Version 1.0

---

## Developed By

Data Analytics & Innovation

Tamarind Group
