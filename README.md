# Shop Sales Analysis — Excel Project

Excel-based end-to-end analysis of a small retail shop's sales data (300 synthetic
transactions, Jan–Dec 2025). Built to practice the full Data Analyst workflow:
**Data Cleaning → Data Processing → Data Analysis → Data Slicing → Data Insights.**

## Workflow

### 1. Data Cleaning
- Checked for duplicate Order IDs — none found
- Checked for blank/missing values across all columns — none found
- Verified date formatting (proper date type, not text)
- Verified categorical consistency (Region, Payment Mode values standardized)

### 2. Data Processing
- Added **Order Month** and **Weekday** columns extracted from Date
- Added **Sales Category** column classifying each order as High (>₹2000),
  Medium (₹1000–2000), or Low (≤₹1000) using `IFS()`
- Unit Price and Category pulled via lookup logic to keep the source data consistent

### 3. Data Analysis
- Built Pivot Tables for Sales by Category, Sales by Region, and Average Order Value
- Visualized with Bar Chart (Category) and Pie Chart (Region)

### 4. Data Slicing
- Added a Slicer on the pivot tables for interactive filtering by Region / Category / Payment Mode

### 5. Data Insights

| Metric | Value |
|---|---|
| Total Sales | ₹11,95,595 |
| Total Orders | 300 |
| Average Order Value | ₹3,985 |
| Best-Selling Category | Furniture |

**Sales by Category**
- Furniture leads by a large margin (₹7,13,713) — driven by high-ticket items like Study Table and Office Chair
- Apparel is second (₹2,06,554), followed by Electronics (₹1,40,196) and Home & Kitchen (₹90,655)
- **Insight:** Furniture drives revenue through high unit price rather than high volume — a few expensive items contribute disproportionately to total sales

**Sales by Region**
- South leads (₹4,05,976), followed by West (₹3,06,761), East (₹2,48,941), North (₹2,33,917)
- **Insight:** South and West together contribute ~60% of total sales — worth investigating what's driving stronger performance there (marketing reach, store presence, demographics) and applying similar strategies in North

**Payment Mode**
- Card is the most used (₹3,67,784), Cash and UPI are close behind (~₹2,80,000 each), Net Banking lowest (₹2,66,107)
- **Insight:** Fairly even split across modes means no single payment channel can be deprioritized — keep all options supported

**Top 5 Products by Sales Value**
1. Study Table — ₹3,65,878
2. Office Chair — ₹2,79,920
3. Jeans — ₹1,57,179
4. Bluetooth Speaker — ₹82,445
5. LED Desk Lamp — ₹67,915

**Sales Category Distribution (order-level)**
- High-value orders: 43.3% (130 orders)
- Low-value orders: 40.7% (122 orders)
- Medium-value orders: 16.0% (48 orders)
- **Insight:** Orders cluster at the two extremes rather than the middle — customers are either buying big-ticket furniture/electronics or small everyday items, with fewer mid-range purchases. A "mid-range bundle" promotion could help capture that gap.

**Recommendation:** Focus marketing spend on the South/West regions where sales
are already strongest, and test cross-selling low-value items (stationery, personal
care) alongside high-value Furniture purchases to lift the Medium-order segment.

## Files
- `Shop_Sales_Analysis.xlsx` — full workbook: Dashboard, Raw Data, Analysis (pivot-style + charts), Price Lookup

## Skills demonstrated
`INDEX`/`MATCH`, `SUMIFS`, `IFS`, `COUNTBLANK`, Pivot Tables, Slicers, Bar/Pie/Line Charts, Data Cleaning checks

## Tools
Built with `openpyxl` (Python), recalculated with LibreOffice to verify zero formula errors.

## Note
Data is synthetically generated for practice purposes, not real shop data.
