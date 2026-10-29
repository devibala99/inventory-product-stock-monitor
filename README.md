# Product Inventory Stock-Level Monitor

An Excel workbook I built to track stock levels for 200 products. It flags what is low or out of stock, shows stock value by category and supplier, and gives a list of what to reorder.

Data as-of date: 27 October 2026. Days since restock are worked out from this fixed date, not from today's date, so the numbers stay the same whenever you open the file.

## Why I built it
Checking 200 products by eye is slow and easy to get wrong. Running out of stock loses sales, and overstocking ties up money. I wanted one place that shows both problems and tells you what to order and from whom.

## Dashboard
![Dashboard](screenshots/dashboard.png)

![Summary](screenshots/summary.png)

## What I found
- 73 of the 200 products need attention. 11 are out of stock and 62 are low.
- Personal Care has 22 products on alert (out of 24) and Electronics has 25 (out of 35). Together they make up 47 of the 73 alerts.
- Jasmine General Trading has the most alerts of any supplier (10). The four suppliers with the most alerts account for 33 of the 73.
- Grocery and Stationery hold about 67% of all units but only about 37% of the stock value.
- Total stock value is about ₹95.73 lakh.

## What I did in Excel
- Cleaned the data with TRIM, UPPER and PROPER, removed duplicates, and added dropdown lists with Data Validation.
- Looked up supplier names, lead times and prices with XLOOKUP and INDEX-MATCH.
- Worked out stock status, stock value, reorder quantity and days since restock with IFS, IF, IFERROR and DATEDIF.
- Built summaries by category and supplier with SUMIFS, COUNTIFS and AVERAGEIFS, plus check cells that should all show 0.
- Added conditional formatting, three PivotTables, charts and slicers.
- Built a dashboard with KPI cards (using named ranges) and a reorder list that updates by itself using FILTER.

## How the reorder rules work
- Quantity 0 is Out of Stock.
- Quantity below the reorder level is Low Stock.
- Anything else is In Stock.
- Reorder quantity is (2 x reorder level) minus the current quantity. It is 0 for products that are In Stock.

## Sheets in the workbook
1. `Dashboard`: KPI cards and charts
2. `Summary`: category and supplier tables, totals and checks
3. `Pivot`: PivotTables, charts and slicers
4. `Reorder_List`: products that need reordering
5. `Insights`: my business observations
6. `Inventory_Data`: the main product table
7. `Supplier_Master`: supplier details
8. `Product_Master`: product reference prices
9. `Lists`: category list used by the dropdown
10. `Data_Issues_Log`: errors I added and how I fixed them
11. `Data_Dictionary`: what each column means
12. `Formula_Log`: the formulas I used
13. `Answer_Key`: a reference copy I used to check my formulas

## The data
The 200 products are synthetic, generated for this project. They are not real company data. I added some mistakes on purpose (duplicate rows, extra spaces, wrong capital letters) so I could practise cleaning. I lost the original uncleaned file, so only the cleaned workbook is here. The `Data_Issues_Log` sheet and the report show what was fixed and the counts before and after.

## Files
- [Workbook](excel/inventory_stock_monitor.xlsx)
- [Report (PDF)](docs/Inventory_Stock_Monitor_Report.pdf)
- [Screenshots](screenshots/)

## How to open it
Open the workbook in Excel (Microsoft 365 or Excel 2021 or newer, because of the FILTER function). Start with the Dashboard and Summary sheets. If you change any data, go to the Pivot sheet and use Data, then Refresh All.

## Limits
- I used Excel Online, so there are no macros.
- The slicers are on the Pivot sheet and only filter PivotTable 3. I could not move them to the dashboard.
- There is no sales history, so the reorder levels are assumptions.
- The reorder rule is my own simple one. It ignores each supplier's standard order size.

## What I would do next
- Add macros to refresh and reset in desktop Excel.
- Connect the slicers to all the PivotTables.
- Test the reorder levels with real sales and supplier lead-time data.
