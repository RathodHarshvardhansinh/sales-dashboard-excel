# Simple Sales Dashboard (Excel)

**Track:** Data Analytics Internship | **Task 3:** Simple Sales Dashboard Design
**Tool:** Microsoft Excel (formulas, data validation dropdown filters, charts, conditional formatting)

## Objective
Turn sales data and KPIs into an interactive view that a stakeholder can use on their own: sales by product, region and month.

## Files
| File | Purpose |
|---|---|
| `Sales_Dashboard.xlsx` | The interactive dashboard (3 sheets: Dashboard, Data, Calc) |
| `README.md` | This file |

## Dataset
1,200 Superstore-style orders from 2025, with columns: Order ID, Order Date, Month, Region, Category, Product, Quantity, Sales, Profit. This is a generated sample for practice, not a real company's data. A real dataset (for example Sample Superstore) can be pasted into the `Data` sheet, columns A to I from row 2, keeping Month as text (Jan to Dec).

## How it works
The workbook has three sheets that work together:

1. **Data**: the raw orders. Nothing here is calculated.
2. **Calc**: helper tables (sales and profit by product, by region, by category). Each cell is a `SUMIFS` formula that adds up Sales or Profit only for rows matching the current filters.
3. **Dashboard**: the filters, KPI cards and charts a viewer sees.

### Filters
- The yellow cells `C5` (Month) and `F5` (Region) use Excel **Data Validation** to create dropdown lists.
- On the Calc sheet, `B1` and `B2` convert the dropdown choice into a criterion: if the choice is "All" it becomes `*` (match everything), otherwise it is the chosen value (for example `Mar` or `West`).
- Every formula uses those criteria, so changing a dropdown recalculates everything.

Example formula (Total Sales):
`=SUMIFS(Sales, Month, MonthCriteria, Region, RegionCriteria)`

### KPIs
| KPI | How it is calculated |
|---|---|
| Total Sales | `SUMIFS` of Sales with both filters |
| Profit Margin | Total Profit / Total Sales |
| Units Sold | `SUMIFS` of Quantity with both filters |
| Top Region | Region with the highest sales (`INDEX` + `MATCH` on the max); uses the month filter only |

### Charts
| Chart | Source | Filters applied |
|---|---|---|
| Sales by Product | Calc, product table | Month + Region |
| Sales by Region | Calc, region table | Month only (so regions can be compared) |
| Monthly Sales Trend | Dashboard, monthly table | Region only (so all months are shown) |
| Profit by Category | Calc, category table | Month + Region |

### Colour coding (conditional formatting)
- Month-over-month growth: **green** if above 0, **red** if below 0.
- Profit Margin card: **green** if at or above the target, **red** if below.

## How to use
1. Open `Sales_Dashboard.xlsx` and go to the **Dashboard** sheet.
2. Pick a month and a region from the yellow dropdowns (or leave both on All).
3. Read the KPI cards, then the charts. See `How_to_Read_Dashboard.docx` for details.

## Outcome
Full-year sample: about $1.25M sales, 9.0% margin (target 10%). West is the top region and sales peak in Nov/Dec. Furniture is loss-making overall (Tables and Bookcases) because of discounts.

## Assumptions
- 10% target margin (editable on `Calc!B3`).
- Sample data, not a real company's figures. Currency shown in USD.

##Author
- Rathod Harshvardhansinh 
