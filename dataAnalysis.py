# ==========================================================
# SALES DATA ANALYSIS PROJECT
# ==========================================================
#
# Reads sales_data.xlsx, performs analysis, exports:
#   - sales_summary.xlsx
#   - sales_report.pdf
#   - chart PNG files
#
# Required:
#   pip install pandas openpyxl matplotlib reportlab
#
# ==========================================================
# Import libraries

import pandas as pd # used for data analysis and manipulation
import matplotlib.pyplot as plt # pyplot module used for creating graphs and charts
from datetime import datetime # imports the datetime class from Python's built-in datetime module

# used to create PDF documents programmatically
# imports classes and functions from the reportlab library
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    PageBreak,
    Table,
    TableStyle,
    Image
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# ----------------------------------------------------------
# LOAD DATA
# ----------------------------------------------------------
# uses Pandas to load data from an Excel file into a DataFrame.
df = pd.read_excel("sales_data.xlsx")

# Create calculated column
df["TotalSale"] = df["Quantity"] * df["UnitPrice"]

# Convert date column
# Converts the values in the "Date" column of a Pandas DataFrame into actual datetime objects
# Tells Pandas to interpret dates as day/month/year
df["Date"] = pd.to_datetime(df["Date"], dayfirst=True)

# ----------------------------------------------------------
# DATA ANALYSIS
# ----------------------------------------------------------

total_sales = df["TotalSale"].sum()

# returns a Series
sales_by_product = df.groupby("Product")["TotalSale"].sum()# returns a Series (1-dimensional-1 column)
sales_by_person = df.groupby("Salesperson")["TotalSale"].sum()# returns a Series
sales_by_region = df.groupby("Region")["TotalSale"].sum()# returns a Series

# finds the row with the highest value in the TotalSale column. idmax() = returns the index label of max TotalSale. df.loc[number] = return full row 
highest_sale = df.loc[df["TotalSale"].idxmax()]# returns a Series

daily_sales = df.groupby("Date")["TotalSale"].sum()# returns a Series
average_daily_sales = daily_sales.mean()# returns a single number

best_product = sales_by_product.idxmax() # finds the index label corresponding to the largest value in sales_by_product
best_salesperson = sales_by_person.idxmax()
best_region = sales_by_region.idxmax()

# creates a new DataFrame containing only the rows where TotalSale is greater than 1000
large_sales = df[df["TotalSale"] > 1000]

# ----------------------------------------------------------
# EXPORT EXCEL REPORT SUMMARY FILE
# ----------------------------------------------------------
# using Pandas ExcelWriter to export multiple tables into a single Excel workbook with multiple sheets.
with pd.ExcelWriter("sales_summary.xlsx") as writer:

# Writes the full DataFrame df into Excel (looks like original data sheet)
    df.to_excel(
        writer,
        sheet_name="Raw Data",
        index=False # removes the Pandas row numbers
    )

    # sales_by_product is a Series (single column of data with labels (an index))
    # Converts it into a DataFrame
    # Writes it to Excel sheet
    sales_by_product.to_frame("TotalSales").to_excel(
        writer,
        sheet_name="Product Summary"
    )

    sales_by_person.to_frame("TotalSales").to_excel(
        writer,
        sheet_name="Salesperson Summary"
    )

    sales_by_region.to_frame("TotalSales").to_excel(
        writer,
        sheet_name="Region Summary"
    )

# ----------------------------------------------------------
# CREATE CHARTS
# ----------------------------------------------------------

sales_by_product.plot(kind="bar")
plt.title("Sales by Product")
plt.xlabel("Product")
plt.ylabel("Sales ($)")
plt.tight_layout()
plt.savefig("sales_by_product.png")
plt.close()

sales_by_region.plot(
    kind="pie",
    autopct="%1.1f%%" # 1 decimal place float, literal % symbol (escaped)
)
plt.title("Sales by Region")
plt.ylabel("")
plt.tight_layout()
plt.savefig("sales_by_region.png")
plt.close()

daily_sales.plot(
    kind="line",
    marker="o"
)
plt.title("Daily Sales Trend")
plt.xlabel("Date")
plt.ylabel("Sales ($)")
plt.grid(True)
plt.tight_layout()
plt.savefig("daily_sales.png")
plt.close()

# ----------------------------------------------------------
# CREATE PDF REPORT
# ----------------------------------------------------------
#
# The file will be saved as sales_report.pdf in the
# current working directory.
#
pdf = SimpleDocTemplate("sales_report.pdf")

# ----------------------------------------------------------
# Load ReportLab's built-in collection of text styles.
#
# These predefined styles control the appearance of text
# in the PDF, including font, size, spacing, and formatting.
#
# Examples:
#   Title
#   Heading1
#   Heading2
#   Normal
#
styles = getSampleStyleSheet()

# ----------------------------------------------------------
# Create an empty list that will contain every object
# added to the report.
#
# Examples of objects:
#   Paragraphs
#   Tables
#   Images
#   Page breaks
#
content = []

content.append(
    # content.append() adds the Paragraph object to the
    # content list so it will be included when the PDF is built.
    Paragraph("Sales Analysis Report", styles["Title"]) # Creates a PDF text object.
)

# Add the report generation date and time to the PDF.
# datetime.now() retrieves the current system date and time.
content.append(
    Paragraph(
        f"Generated: {datetime.now():%d/%m/%Y %H:%M}",
        styles["Normal"]
    )
)

# Add vertical spacing after the previous PDF element.
#
# Spacer(width, height) creates an empty area in the PDF.
#
content.append(Spacer(1, 20))

# ----------------------------------------------------------
# EXECUTIVE SUMMARY SECTION
# ----------------------------------------------------------
#
# This section provides a high-level overview of the
# business performance.
#
# It is intended for managers who want a quick snapshot
# without reviewing detailed tables.
#
# Statistics included:
#   - Total Sales
#   - Average Daily Sales
#   - Best Product
#   - Best Salesperson
#   - Best Region
#
# ----------------------------------------------------------
# Add a main section heading to the PDF report.
content.append(
    Paragraph("Executive Summary", styles["Heading1"])
)

# Create a list containing key summary statistics that will
# appear in the Executive Summary section of the PDF.
#
# f-strings are used to insert calculated values into the
# text, making the report dynamic and automatically updated
# whenever the data changes.
#
# :,.2f formats numbers as currency with:
#   , = thousands separator
#   .2f = 2 decimal places
#
# Example:
#   Total Sales: $12,345.67
#
summary_items = [
    f"Total Sales: ${total_sales:,.2f}",
    f"Average Daily Sales: ${average_daily_sales:,.2f}",
    f"Best Product: {best_product}",
    f"Best Salesperson: {best_salesperson}",
    f"Best Region: {best_region}"
]

for item in summary_items:
    content.append(Paragraph(item, styles["Normal"]))

content.append(Spacer(1, 20))

# Highest sale
# Add a heading to identify the Highest Sale section in
# the PDF report.
content.append(Paragraph("Highest Sale", styles["Heading1"]))
# Loop through every field in the highest sale record and
# add it to the PDF as a paragraph.
#
# highest_sale is a pandas Series containing the complete
# row of data for the largest sale transaction.
#
# Example:
#   Date         2026-01-07
#   Product      Laptop
#   Quantity     3
#   UnitPrice    1200
#   TotalSale    3600
#
# .items() returns each column name (key) and its
# corresponding value.
#
# The f-string combines the field name and value into a
# readable format:
#
#   Product: Laptop
#   Quantity: 3
#   TotalSale: 3600
#
# Each line is added to the PDF using the Normal text style.
#
for key, value in highest_sale.items():
    content.append(
        Paragraph(f"{key}: {value}", styles["Normal"])
    )

# PageBreak() forces all subsequent content to start on
# a new page rather than continuing on the current page.
#
content.append(PageBreak())

# ----------------------------------------------------------
# SALES BY PRODUCT TABLE
# ----------------------------------------------------------
#
# Create a table showing total sales for each product.
#
# The first row is used as a header row.
# Remaining rows are generated automatically from the
# sales_by_product summary calculated earlier.
#
# Example:
#
# Product     Sales ($)
# Laptop      7200.00
# Monitor     1800.00
#
# ----------------------------------------------------------
# Add a section heading to identify the Sales By Product
# table in the PDF report.

content.append(
    Paragraph("Sales By Product", styles["Heading1"])
)

# Create the table data structure.
#
# The first row is the table header and will appear at
# the top of the table.
#
# Example:
#
# Product      Sales ($)
# list containing nested list
table_data = [["Product", "Sales ($)"]]

# Loop through the sales_by_product summary and add each
# product and its total sales amount as a new row in the
# table.
#
# Example rows added:
#
# Laptop       $7,200.00
# Monitor      $1,800.00
# Mouse         $575.00
#
# :,.2f formats the sales amount with:
#   , = thousands separator
#   .2f = 2 decimal places
#

for product, sales in sales_by_product.items():
    table_data.append([product, f"${sales:,.2f}"])

# Create a ReportLab Table object using the completed
# table data.
#
# The table will contain:
#   - Header row
#   - One row for each product
#
table = Table(table_data)

# Apply formatting to improve the appearance of the table.
#
# GRID:
#   Draws borders around every cell.
#
# BACKGROUND:
#   Applies a light grey background colour to the header
#   row to distinguish it from the data rows.
#
# FONTNAME:
#   Makes the header row bold for improved readability.
#
table.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")
]))

# Add the completed table to the PDF content list.
#
# The table will be included in the report when
# pdf.build(content) is executed.
#
content.append(table)

content.append(Spacer(1, 20))

# Salesperson table
content.append(
    Paragraph("Sales By Salesperson", styles["Heading1"])
)

# Initializes the table as a list of rows
# First row is the column headers
table_data = [["Salesperson", "Sales ($)"]]

# Fill table with data
# Adds each row to the table
# Formats sales as currency:
# :,.2f means:
# comma separators → 3,400.00
# 2 decimal places
# ["Alice", "$1,200.50"]
for person, sales in sales_by_person.items():
    table_data.append([person, f"${sales:,.2f}"])

# Converts the list into a PDF table object
table = Table(table_data)
# style the table
table.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 1, colors.black), # draws black borders around all cells
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")
]))
# Inserts the table into the final PDF document
content.append(table)

content.append(Spacer(1, 20))

# Region table
content.append(
    Paragraph("Sales By Region", styles["Heading1"])
)

table_data = [["Region", "Sales ($)"]]

for region, sales in sales_by_region.items():
    table_data.append([region, f"${sales:,.2f}"])

table = Table(table_data)
table.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")
]))
content.append(table)

content.append(PageBreak())

# Large sales
content.append(
    Paragraph("Sales Over $1000", styles["Heading1"])
)

# defines the column names
table_data = [["Date", "Product", "Amount"]]

# Loop through filtered DataFrame
# large_sales is a pandas DataFrame
# .iterrows() loops through each row
# _ means you’re ignoring the index
for _, row in large_sales.iterrows():
    # Add each row to the table
    table_data.append([
        str(row["Date"].date()), # Converts datetime → just date
        row["Product"], # Directly inserts product name
        f"${row['TotalSale']:,.2f}" # e.g $1,234.50
    ])

# Converts list → PDF table object
table = Table(table_data)
table.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")
]))
content.append(table)

content.append(PageBreak())

# ----------------------------------------------------------
# EMBED CHARTS INTO PDF
# ----------------------------------------------------------
#
# Earlier in the program, matplotlib charts were saved as
# PNG image files.
#
# These images are now inserted into the PDF so users can
# visually analyse trends and patterns.
#
# Charts included:
#
# 1. Sales by Product
#    Compares revenue generated by each product.
#
# 2. Sales by Region
#    Shows each region's percentage contribution.
#
# 3. Daily Sales Trend
#    Displays sales performance over time.
#
# PageBreak() is used so each chart appears on a separate
# page for improved readability.
#
# ----------------------------------------------------------
content.append(
    Paragraph("Sales By Product Chart", styles["Heading1"])
)
content.append(Image("sales_by_product.png", width=450, height=300))

content.append(PageBreak())

content.append(
    Paragraph("Sales By Region Chart", styles["Heading1"])
)
content.append(Image("sales_by_region.png", width=450, height=300))

content.append(PageBreak())

content.append(
    Paragraph("Daily Sales Trend", styles["Heading1"])
)
content.append(Image("daily_sales.png", width=450, height=300))

# ----------------------------------------------------------
# BUILD PDF DOCUMENT
# ----------------------------------------------------------
#
# The build() method processes every object stored in the
# content list and writes them to the PDF file in order.
#
# Once this step completes, sales_report.pdf is created
# and ready to be viewed or distributed.
#
# ----------------------------------------------------------

pdf.build(content)

print("Reports generated successfully.")
