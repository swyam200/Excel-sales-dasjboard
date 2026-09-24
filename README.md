# Interactive Excel Sales Dashboard

A 4-sheet Excel workbook that turns raw retail order data into a live,
formula-driven sales dashboard — built with `openpyxl`.

## What it does
- **Raw Data** sheet holds 2,000 order records
- **Region Summary**, **Category Summary**, and **Monthly Trend** sheets use
  `SUMIFS`/`COUNTIFS` formulas (not hardcoded numbers) to roll up revenue,
  order counts, and average order value — the sheet recalculates automatically
  if the raw data changes
- Each summary sheet includes a matching native Excel chart (bar, pie, line)

## Tech
Python, openpyxl, Excel formulas (SUMIFS, COUNTIFS)

## Files
- `Sales_Dashboard.xlsx` — the deliverable
- `build_dashboard.py` — script that generates it from `retail_sales.csv`
