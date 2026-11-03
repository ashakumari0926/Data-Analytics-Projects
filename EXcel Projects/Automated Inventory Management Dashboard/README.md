📦 Automated Inventory Management Dashboard

An Excel-based Automated Inventory Management System designed to simplify inventory tracking, manage purchase and sales transactions, monitor stock availability, and analyze business performance through an interactive dashboard.

Built using Microsoft Excel, Excel Tables, Lookup Formulas, Pivot Tables, and Charts, this project automates inventory calculations and provides meaningful business insights through a centralized dashboard.

📊 Dashboard Preview

Automated Inventory Management Dashboard

🎯 Project Overview

The Automated Inventory Management Dashboard is designed to make inventory management easier, faster, and more efficient.

It provides a centralized system to manage customers, products, vendors, purchases, sales, and inventory records.

The workbook uses sample and randomly generated data to demonstrate inventory management and automation. It automatically retrieves product and customer information, calculates transaction amounts, updates available stock, and generates low-stock alerts.

The interactive dashboard presents important business KPIs and visualizations to help users monitor inventory and business performance.

✨ Key Features
Interactive Dashboard: Displays important business KPIs and visual charts.
Customer Management: Maintains customer master records.
Product Management: Stores product IDs, names, and pricing details.
Vendor Management: Maintains vendor information and contact details.
New Entry: Form-style worksheet for adding records.
Purchase Management: Records purchase transactions with automatic data retrieval.
Sales Management: Records sales transactions with automatic customer and product details.
Automated Inventory Tracking: Calculates available stock based on purchase and sales records.
Low-Stock Alerts: Identifies products with fewer than 5 available units and displays vendor contact information.
Top 5 Products: Visualizes top-performing products.
Automated Calculations: Uses Excel formulas to calculate transaction amounts and stock values.
Pivot Table Analysis: Generates dashboard KPIs and supporting charts.

🛠️ Tools & Technologies
Tools / Features	Purpose
Microsoft Excel	Dashboard development and data management
Excel Tables	Structured data storage
Pivot Tables	Data summarization and analysis
Pivot Charts	Data visualization
VLOOKUP	Automatic data retrieval
SUMIF	Purchase and sales calculations
IF / IFERROR	Conditional calculations and error handling
GETPIVOTDATA	Extracting KPI values from Pivot Tables
Data Validation	Drop-down lists for data entry
Excel Formulas	Inventory and financial calculations

📁 Workbook Structure
Worksheet	Description
Dashboard	Main dashboard with KPIs and charts
Customers	Customer master data
Products	Product master data
Vendors	Vendor master data
New Entry	Form-style data entry sheet
Purchase	Purchase transaction records
Sales	Sales transaction records
Inventory	Automated stock calculation and low-stock alerts
Pivot_table	Pivot Tables supporting dashboard calculations

📋 Excel Tables & Named Ranges
Excel Tables
Table Name	Worksheet	Purpose
Customer_Data	Customers	Customer master records
Productsdata	Products	Product master records
Vendors_data	Vendors	Vendor master records
Purchase_data	Purchase	Purchase transactions
Sales_data	Sales	Sales transactions
Inventory_data	Inventory	Inventory calculations
Named Ranges

The following named ranges are used for drop-down lists through Excel Data Validation:

Cust_ID
ProductID
Product_No

These named ranges help maintain consistent customer and product references during purchase and sales entry.

⚙️ Automation & Formulas

The workbook uses Excel formulas to automate data retrieval, transaction calculations, and inventory management.

1. Purchase Sheet
Formula	Purpose
VLOOKUP	Automatically retrieves product names
VLOOKUP	Retrieves vendor information
VLOOKUP	Fetches product details
Cost Price × Units	Calculates total purchase amount
IFERROR	Handles lookup errors

Automatically Retrieve Product Name:

excel
=IFERROR(VLOOKUP([@ProductID],Productsdata[#All],2,FALSE),"")


Calculate Total Purchase Amount:

excel
=IFERROR([@[Cost_Price (₹)]]*[@Units],"")

2. Sales Sheet
Formula	Purpose
VLOOKUP	Retrieves customer names
VLOOKUP	Retrieves product names
VLOOKUP	Fetches available stock
VLOOKUP	Retrieves product selling price
Sell Unit × Price	Calculates total sales amount

Automatically Retrieve Customer Name:

excel
=IFERROR(VLOOKUP([@Cust_ID],Customer_Data[#All],2,0),"")


Calculate Total Sales Amount:

excel
=IFERROR([@Sell_Unit]*[@Price],"")

3. Inventory Sheet

Inventory is automatically calculated using purchase and sales transaction records.

Formula	Purpose
SUMIF	Calculates total purchased units
SUMIF	Calculates total sold units
Purchase Units - Sales Units	Calculates available stock
Selling Price × Available Stock	Calculates stock value
IF	Identifies low-stock products
VLOOKUP	Retrieves vendor contact information

Available Stock Formula:

excel
=[@[Purchase Unit]]-[@[Sell Unit]]


Stock Value Formula:

excel
=[@[Selling_Price (₹)]]*[@[Available_Stock]]


Low-Stock Alert Formula:

excel
=IF([@Available_Stock]<5," 📞 "&[@ProductName]&" Less Quantity available order now! PH "&IFERROR(VLOOKUP([@ProductID],Vendors_data[#All],5,0),""),"")


Products with fewer than 5 available units are automatically flagged for restocking.

4. Pivot Table Sheet – Dashboard KPIs

The GETPIVOTDATA function extracts summarized values from Pivot Tables to display important dashboard metrics.

KPI	Formula / Calculation
Total Customers	=GETPIVOTDATA("Customer_name",$A$1)
Total Products	=GETPIVOTDATA("ProductName",$A$5)
Total Purchase Amount	=GETPIVOTDATA("Total_Amount (₹)",$A$13)
Total Sales Amount	=GETPIVOTDATA("Total_Amount (₹)",$A$17)
Total Stock Value	=GETPIVOTDATA("Stock Amount(₹)",$A$21)
COGS	=C14-C22
Profit / Loss	=C22+C18-C14
Profit / Loss Ratio	=C25/E14

These calculations provide a summarized view of inventory and financial performance.

Note: The Profit / Loss and COGS calculations reflect the workbook's current calculation logic and should be validated against the business's accounting method before financial reporting.

📈 Business Insights

The dashboard helps users:

Monitor overall inventory performance.
Identify the top 5 products.
Track purchase and sales transactions.
Analyze total purchase and sales amounts.
Monitor available stock levels.
Identify products that require immediate restocking.
Evaluate total inventory value.
Review estimated profit/loss performance.

📌 Dataset Information
The workbook contains sample and randomly generated data.
No external raw dataset or separate CSV file is required.
All sample records are maintained within the Excel workbook.
Users can replace the sample data with their own inventory records.
The dashboard and calculations demonstrate automated inventory tracking and reporting.

👨‍💻 Author

Asha Kumari
Project: Automated Inventory Management Dashboard
Platform: Microsoft Excel
Domain: Inventory Management & Data Analytics