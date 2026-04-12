# Profit Performance Analysis Project

A four-year transactional sales analysis built entirely in Microsoft Excel, delivering two interactive dashboards that surface profit trends, customer behavior, and product performance across time, geography, and demographics.

🔍 Excel File? Check it out here: [Excel Project File](https://github.com/Shihab412/Profit_Performance_Analysis_in_Excel/tree/main/Excel%20File)

---

## Project Overview

This project was built in response to a business request to visualize four years of sales data and assess company performance — with a primary focus on profit. The analysis spans two dashboards and four supporting analysis sheets, covering time-series trends, product profitability, customer segmentation, and geographic distribution.

**Tool:** Microsoft Excel (with VBA/Macros enabled — `.xlsm`)  
**Source:** [Data with Decision](https://www.youtube.com/@datalab365)  
**Dataset:** Four years of transactional sales data

---

## File Structure

| Sheet | Type | Purpose |
|---|---|---|
| Time Series Dashboard | Dashboard | KPI comparisons, monthly/quarterly/weekly profit trends |
| Dashboard 2 | Dashboard | Product, customer, and geographic profit analysis |
| Analysis 1 | Analysis | KPI aggregations and yearly performance data |
| Analysis 1 (2) | Analysis | Time-series breakdowns: monthly, weekly, quarterly |
| Product Analysis | Analysis | Product profit rankings, color analysis, pricing types |
| Customer Analysis | Analysis | Customer rankings, age groups, gender, country mapping |

---

## Dashboard 1 — Time Series Performance

This dashboard provides a longitudinal view of business health, allowing stakeholders to assess performance relative to the prior year and identify seasonal or temporal patterns.

**What it covers:**

- **KPI Cards (vs. Prior Year):** COGS, Revenue, Order Quantity, Profit, Profit Margin, and Transaction count — each benchmarked against the previous year.
- **Yearly Performance Highlights:** Filters and displays only the years that exceeded average performance across Revenue, Profit, and Transactions.
- **Monthly Profit Trend:** A time-series chart tracking profit month by month to identify seasonal cycles.
- **Profit by Week Type:** Breaks profit down between weekdays and weekends to assess when sales are strongest.
- **Quarterly Profit Analysis:** Evaluates profit contribution by quarter to reveal seasonal peaks and troughs.
- **Profit by Weekday:** Day-level profit breakdown (Sunday through Saturday) to identify the most and least profitable days.

---

## Dashboard 2 — Product, Customer & Geographic Analysis

This dashboard focuses on the *who* and *what* behind profitability — identifying top performers, demographic patterns, and geographic concentration.

**What it covers:**

- **Top 5 Profitable Products:** Ranks the five highest-profit products and shows their combined percentage contribution against all others.
- **Top 5 Profitable Customers:** Identifies the five highest-value customers and their share of total profit.
- **Profit by Gender:** Compares profit generated across male and female customer segments.
- **Profit by Product Color:** Ranks product colors by total profit, with conditional highlighting of the best-selling color.
- **Profit by Pricing Type:** Analyzes profit performance across pricing tiers (e.g., Expensive vs. Less Expensive).
- **Country-wise Profit (Custom Map):** A custom map visualization showing profit distribution across customer countries — useful for identifying geographic concentration or expansion opportunities.
- **Profit by Age Group:** Segments customers into age bands and compares their respective profit contributions.

---

## Technical Highlights

- Dynamic chart data tables using `LARGE`, `INDEX`, and `MATCH` for ranked visualizations
- `XLOOKUP` used for geographic coordinate mapping to power the custom country map
- Array formulas for conditional highlighting of top performers across multiple categories
- Percentage contribution calculations for top-N vs. "Others" comparisons
- Week type and weekday segmentation logic built into pivot-ready analysis tables
- Fully formula-driven analysis sheets — no hardcoded aggregates

---

## How to Use

1. Open `Project.xlsm` in **Microsoft Excel** (not Google Sheets or LibreOffice — VBA macros are required).
2. Enable macros if prompted.
3. Navigate to **Time Series Dashboard** or **Dashboard 2** using the sheet tabs.
4. Use any slicers or filters on the dashboards to explore the data interactively.

> The four analysis sheets (Analysis 1, Analysis 1 (2), Product Analysis, Customer Analysis) contain the underlying aggregation tables that power the dashboard charts. They are not intended for direct presentation.

---

## Key Insights Enabled

- Which years outperformed the company average — and by how much
- Whether weekday or weekend transactions drive more profit
- Which quarter consistently underperforms or over-performs
- Which five products and five customers account for a disproportionate share of total profit
- How profit distributes across countries, genders, age groups, and pricing tiers

---

## Author

**Sahadat Ullah Mollah**  
Data Analytics Portfolio Project  
[LinkedIn](www.linkedin.com/in/sahadat-ullah)
