# Excel Sales Performance Analysis & Dashboard

## Project Overview

This Excel project performs **Sales Performance Analysis, Data Analysis, What-If Analysis, Customer Analysis, Product Analysis, and Dashboard Visualization** using a 250-transaction sales dataset.

The workbook demonstrates practical Excel functions, formulas, PivotTables, PivotCharts, Conditional Formatting, Scenario Manager, Goal Seek, Linear Regression, FILTER, TEXT functions, and dashboard reporting.

---

## Workbook Sheets

### 1. Raw Data
Contains the complete sales dataset with 250 transactions.

Main fields include:
- Transaction ID
- Date
- Customer ID
- Customer Name
- Product ID
- Product Name
- Category
- Quantity
- Unit Price
- Payment Method
- Region
- Customer Segment
- Customer Since
- Total Amount
- Customer Months
- Month End
- Today
- Current Date Time
- Customer Tenure
- Timestamp

### 2. Pivot Table
Contains PivotTable-based analysis including:
- Revenue by Region
- Revenue by Product
- Quantity by Product
- Grand Total Revenue
- Grand Total Quantity

### 3. Scenario Summary
Contains Scenario Manager results for:
- Normal
- High Sales
- Low Sales

### 4. Summary
Contains the **Linear Regression using Data Analysis ToolPak**.

Regression analysis includes:
- Regression Statistics
- ANOVA
- Coefficients
- Standard Error
- R Square
- Significance F
- Quantity coefficient analysis

### 5. Analysis
Contains the main analytical calculations and business insights, including:
- Sales Performance KPIs
- Total Revenue
- Total Transactions
- Total Quantity Sold
- Average Transaction Value
- Monthly Revenue Analysis
- Region Analysis
- Product Analysis
- Goal Seek Analysis
- Scenario Analysis
- Arrow/Indicator Analysis
- FILTER – High Value Customers
- High-Value Customer Analysis
- Most Frequently Purchased Product
- Repeating Customer Names
- Compare Two Lists
- Matching Names
- Abbreviations using TEXT functions
- Business Findings and Storytelling

### 6. Visualizations
Used for the visual representation of the analysis through charts.

Charts include:
- Revenue by Product – Bar Chart
- Monthly Revenue Trend – Line Chart
- Revenue Distribution by Region – Pie Chart

### 7. Dashboard
Contains the final Sales Performance Dashboard with KPI indicators and Revenue Performance analysis.

---

# Key Excel Functions Used

## Date & Time Functions

### TODAY
Returns the current date.

Example:
```excel
=TODAY()
```

### NOW
Returns the current date and time.

Example:
```excel
=NOW()
```

### DATEDIF
Calculates the difference between two dates.

Used for:
- Customer Months
- Customer Tenure

### EOMONTH
Returns the last day of a month.

Example:
```excel
=EOMONTH(B2,0)
```

---

# FILTER Function

The FILTER function is used to return multiple matching records dynamically.

Example:
```excel
=FILTER(Customer_Range,Condition_Range)
```

It is used in the workbook for customer analysis and matching records.

> Note: FILTER results are dynamic-array results and should remain as a normal Excel range rather than being converted into an Excel Table.

---

# Conditional Formatting

Conditional Formatting is used to make important values easier to understand visually.

The workbook includes indicator-based analysis such as:

- ↑ High Revenue
- → Average / Stable Revenue
- ↓ Low Revenue

The Dashboard also contains a Revenue Performance indicator.

Indicator logic:
- Green / High = Target achieved
- Yellow / Average = Close to target
- Red / Low = Below target

---

# Timestamp

A Timestamp column is included in the Raw Data sheet.

Formula used:

```excel
=IF(A2="","",IF(T2="",NOW(),T2))
```

Iterative calculation is enabled so that the timestamp can remain fixed after being generated.

---

# What-If Analysis

## Scenario Manager

Three scenarios are included:

| Scenario | Revenue |
|---|---:|
| Normal | ₹229,192.47 |
| High Sales | ₹250,000 |
| Low Sales | ₹180,000 |

This allows comparison of different sales situations.

## Goal Seek

Goal Seek is used to determine the required value needed to reach a target revenue.

Example analysis:
- Quantity = 10
- Target Revenue = ₹5,000
- Required Unit Price = ₹500

Therefore:

**10 × ₹500 = ₹5,000**

---

# Linear Regression

Linear Regression is performed using the Excel **Data Analysis ToolPak**.

The model analyzes the relationship between:

**Quantity → Total Amount**

Important regression results include:

- Observations: 250
- R Square: approximately 0.210
- Multiple R: approximately 0.458
- Quantity coefficient: approximately ₹353.79

### Interpretation

The regression indicates a positive relationship between Quantity and Total Amount. As quantity increases, total sales amount generally tends to increase.

---

# Sales Performance KPIs

The workbook calculates the following major KPIs:

- **Total Revenue:** ₹229,192.47
- **Total Transactions:** 250
- **Total Quantity Sold:** 753
- **Total Customers:** 50
- **Average Transaction Amount:** approximately ₹916.77

These KPIs are displayed on the Dashboard for quick business monitoring.

---

# Product Analysis

The workbook analyzes revenue and quantity across products such as:

- Laptop
- Smartphone
- Desk
- Monitor
- Bookshelf
- Office Chair
- Keyboard
- Headphones
- Blender
- Coffee Maker

### Key Product Insights

- **Laptop** generated the highest revenue.
- **Smartphone** generated the second-highest revenue.
- **Bookshelf** had the highest quantity sold.

---

# Region Analysis

Revenue is analyzed across:

- East
- North
- West
- Central
- South

### Key Region Insight

**East** generated the highest revenue.

**South** generated the lowest revenue.

This can help management identify strong-performing and underperforming regions.

---

# Customer Analysis

The workbook includes customer-level analysis using:

- SUMIF
- COUNTIF
- FILTER
- UNIQUE
- Customer ID
- Customer Name
- Total Purchase

The project also includes:
- High-Value Customer analysis
- Repeating Customer analysis
- Most frequently purchased product analysis
- Customer name matching

---

# Repeating Customers

COUNTIF is used to identify customers appearing multiple times in the transaction data.

Example:

```excel
=COUNTIF('Raw Data'!D:D,G42)
```

This helps identify repeat customers and customer purchasing frequency.

---

# Compare Two Lists

The workbook includes a section for comparing two customer-name lists and extracting matching names.

Example:

```excel
=FILTER(G112:G121,COUNTIF(H112:H121,G112:G121)>0,"No Matching Names")
```

This demonstrates dynamic list comparison using FILTER and COUNTIF.

---

# Abbreviations Using TEXT Functions

Customer names are converted into initials using TEXT functions.

Example:

```excel
=UPPER(LEFT(TEXTBEFORE(J26," "),1)&LEFT(TEXTAFTER(J26," "),1))
```

Example:

**Rahul Patel → RP**

This demonstrates the use of:
- TEXTBEFORE
- TEXTAFTER
- LEFT
- UPPER

---

# Dashboard

The final Dashboard provides a quick overview of sales performance.

## Dashboard KPIs

- Total Revenue
- Total Transactions
- Total Quantity
- Total Customers

## Revenue Performance

The Dashboard compares:

- Actual Revenue
- Target Revenue
- Performance Indicator

Current target:

**₹250,000**

The indicator evaluates whether actual revenue is:
- At/above target
- Close to target
- Below target

---

# Charts

## 1. Bar Chart
**Title:** Revenue by Product

Shows the revenue contribution of each product.

## 2. Line Chart
**Title:** Monthly Revenue Trend

Shows how revenue changes over time.

## 3. Pie Chart
**Title:** Revenue Distribution by Region

Shows the percentage contribution of each region to total revenue.

---

# Business Insights

Based on the analysis:

1. Total revenue generated is **₹229,192.47**.
2. The dataset contains **250 transactions**.
3. Total quantity sold is **753 units**.
4. There are **50 customers**.
5. Laptop is the highest-revenue product.
6. Smartphone is the second-highest-revenue product.
7. Bookshelf has the highest quantity sold.
8. East is the highest-revenue region.
9. South is the lowest-revenue region.
10. Quantity has a positive relationship with Total Amount.
11. Repeat-customer analysis helps identify customer purchasing frequency.
12. Dashboard KPIs provide a quick view of overall sales performance.

---

# Tools & Excel Features Used

- Microsoft Excel
- Excel Tables
- Formulas & Functions
- TODAY
- NOW
- DATEDIF
- EOMONTH
- EDATE
- SUMIF
- SUMIFS
- COUNTIF
- COUNTA
- AVERAGE
- SUM
- FILTER
- UNIQUE
- SORT
- INDEX
- MATCH
- IF
- IFERROR
- TEXTBEFORE
- TEXTAFTER
- LEFT
- UPPER
- Conditional Formatting
- Icon Sets / Arrows
- PivotTables
- PivotCharts
- Slicers
- Timeline
- Scenario Manager
- Goal Seek
- Data Analysis ToolPak
- Linear Regression
- Dashboard Design
- Data Storytelling

---

# Conclusion

This project demonstrates how Excel can be used as a complete **Sales Analytics and Business Intelligence tool**.

The workbook transforms raw transaction data into:
- Structured analysis
- KPI reporting
- Customer insights
- Product insights
- Regional insights
- What-If analysis
- Regression analysis
- Visual charts
- An interactive sales dashboard

The final result provides both **data-driven insights and a professional visual summary** for business decision-making.

---

## Project File

**Excel Workbook:** `Final_Pr(2).xlsx`

**Prepared for:** Excel Data Analysis / Dashboard Project
