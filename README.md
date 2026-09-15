
# 📊 Sales & Data Intelligence Dashboard – Excel Project

<p align="center">
  <img src="https://readme-typing-svg.demolab.com?font=Fira+Code&size=25&pause=1000&color=36BCF7&center=true&vCenter=true&width=700&lines=Sales+%26+Data+Analytics+Dashboard;Excel+Data+Analysis+Project;Pivot+Tables+%7C+Formulas+%7C+Visualizations" alt="Typing SVG" />
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/Data-Analysis-blue?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Pivot-Tables-orange?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Visualization-green?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Dashboard-Excel-purple?style=for-the-badge" />
</p>

---

## 📌 Project Overview

The **Sales & Data Intelligence Dashboard** is an Excel-based data analysis project created to analyze sales transactions, customers, products, regions, payment methods, customer segments, revenue, and purchasing patterns.

The project demonstrates the practical use of **Microsoft Excel** for data cleaning, formula-based calculations, date and time analysis, What-If Analysis, regression analysis, Pivot Tables, charts, conditional formatting, and dashboard creation.

The workbook provides a complete workflow from raw transaction data to business insights and visual reporting.

---

## 🎯 Objectives

The main objectives of this project are:

- Analyze sales transaction data using Microsoft Excel.
- Organize and structure raw sales data.
- Calculate total revenue and transaction-related KPIs.
- Apply Date & Time functions.
- Use Excel formulas for business analysis.
- Use the FILTER function for multi-value results.
- Create timestamps using NOW().
- Apply Conditional Formatting and Arrow Indicators.
- Perform What-If Analysis using Scenario Manager.
- Perform Goal Seek analysis.
- Perform Linear Regression using the Data Analysis ToolPak.
- Analyze products and regions.
- Identify high-value customers.
- Identify frequently purchased products.
- Analyze repeating customer names.
- Compare two customer lists.
- Create customer abbreviations using TEXT functions.
- Create Pivot Tables.
- Create charts and visualizations.
- Build a centralized Excel Dashboard.
- Present business insights through data storytelling.

---

## 🚀 Project Highlights

- 📊 **250 sales transactions analyzed**
- 💰 **Total Revenue: ₹229,192.47**
- 📦 **Total Quantity Sold: 753**
- 👥 **50 Unique Customers**
- 📈 Average Transaction Value: approximately **₹916.77**
- 🔎 FILTER-based customer analysis
- 📅 TODAY, NOW, DATEDIF and EOMONTH functions
- ⏱️ Timestamp using NOW()
- 🎯 Scenario Manager
- 🎯 Goal Seek
- 📉 Linear Regression using Data Analysis ToolPak
- 📊 Pivot Table analysis
- 📈 Bar, Line and Pie Charts
- 🎨 Conditional Formatting with Arrow Indicators
- 📋 High-Value Customer Analysis
- 🔁 Repeating Customer Analysis
- 🔤 Customer Name Abbreviations
- 🔎 Two-List Matching Analysis
- 📊 Interactive Dashboard

---

## 🛠️ Technologies Used

- **Microsoft Excel**
- Excel Tables
- Excel Formulas
- TODAY
- NOW
- DATEDIF
- EOMONTH
- FILTER
- SORT
- UNIQUE
- SUMIF
- SUMIFS
- COUNTIF
- INDEX
- MATCH
- IF
- Scenario Manager
- Goal Seek
- Data Analysis ToolPak
- Linear Regression
- Pivot Tables
- Pivot Charts
- Conditional Formatting
- Dashboard Design

---

## 📂 Project Structure

```text
Sales & Data Intelligence Dashboard
│
├── Raw Data
├── Pivot Table
├── Scenario Summary
├── Summary
├── Analysis
├── Visualizations
└── Dashboard
````

---

## 📘 Workbook Information

| Information         | Details                              |
| ------------------- | ------------------------------------ |
| Project Type        | Sales & Data Analytics               |
| Tool Used           | Microsoft Excel                      |
| Total Transactions  | 250                                  |
| Total Quantity Sold | 753                                  |
| Total Revenue       | ₹229,192.47                          |
| Unique Customers    | 50                                   |
| Main Analysis       | Sales, Products, Customers & Revenue |
| Regression          | Quantity vs Total Amount             |
| What-If Analysis    | Scenario Manager & Goal Seek         |
| Visualization       | Bar Chart, Line Chart & Pie Chart    |
| Dashboard           | Excel Sales Dashboard                |
| Dataset             | Sales Transaction Dataset            |

---

# 📑 Worksheets

## 1️⃣ Raw Data

The **Raw Data** sheet contains the main sales transaction dataset.

The dataset contains **250 transaction records** and includes fields such as:

* Transaction ID
* Date
* Customer ID
* Customer Name
* Product ID
* Product Name
* Category
* Quantity
* Unit Price
* Payment Method
* Region
* Customer Segment
* Customer Since
* Total Amount
* Customer Months
* Month End
* Today
* Current Date Time
* Customer Tenure
* Timestamp

The Raw Data sheet is the main source for the different analysis sections of the project.

---

# 📅 Date & Time Analysis

The project demonstrates multiple Excel Date & Time functions.

## TODAY Function

Used to display the current date.

Example:

```excel
=TODAY()
```

The result is stored in the **Today** column.

---

## NOW Function

Used to display the current date and time.

Example:

```excel
=NOW()
```

The result is stored in the **Current Date Time** column.

---

## DATEDIF Function

DATEDIF is used to calculate customer-related time periods.

Examples include:

* Customer Months
* Customer Tenure

Example:

```excel
=DATEDIF(Customer_Since,Date,"m")
```

This helps analyze how long customers have been associated with the business.

---

## EOMONTH Function

EOMONTH is used to find the end date of a month.

Example:

```excel
=EOMONTH(B2,0)
```

This helps with monthly sales and date-based analysis.

---

# 🔎 FILTER Function

The **FILTER** function is used to return multiple matching records dynamically.

The project uses FILTER for customer-based analysis and multi-value results.

Example:

```excel
=FILTER(G112:G121,COUNTIF(H112:H121,G112:G121)>0,"No Matching Names")
```

This formula extracts names that are present in both lists.

---

# ⏱️ Timestamp Analysis

A timestamp is created using the **NOW()** function.

The project uses an iterative calculation approach so that the timestamp can remain fixed after the transaction is entered.

Example:

```excel
=IF(A2="","",IF(T2="",NOW(),T2))
```

This provides a transaction timestamp based on the entry.

---

# 🎨 Conditional Formatting

Conditional Formatting is used to make important values easier to understand visually.

The project includes:

* Revenue performance indicators
* Arrow indicators
* High Revenue indicators
* Average Revenue indicators
* Low Revenue indicators
* KPI-based visual indicators

Example indicators:

```text
↑ High Revenue
→ Average Revenue
↓ Low Revenue
→ Stable
```

These indicators provide quick visual interpretation of business performance.

---

# 📊 KPI Analysis

The **Analysis** and **Dashboard** sheets contain important sales KPIs.

Main KPIs include:

### 💰 Total Revenue

```text
₹229,192.47
```

### 🧾 Total Transactions

```text
250
```

### 📦 Total Quantity Sold

```text
753
```

### 👥 Total Customers

```text
50
```

### 💵 Average Transaction Value

Approximately:

```text
₹916.77
```

These KPIs provide a quick overview of overall sales performance.

---

# 📈 Monthly Revenue Analysis

Monthly revenue is calculated using SUMIFS.

Example:

```excel
=SUMIFS(SalesData[Total_Amount],SalesData[Date],">="&A10,SalesData[Date],"<"&EDATE(A10,1))
```

The analysis covers monthly revenue from:

```text
April 2024
May 2024
June 2024
July 2024
August 2024
September 2024
October 2024
November 2024
December 2024
January 2025
February 2025
March 2025
April 2025
```

This analysis is useful for identifying revenue trends over time.

---

# 🌍 Region Analysis

The project analyzes revenue generated by different regions.

| Region  | Total Revenue |
| ------- | ------------: |
| East    |    ₹59,288.39 |
| North   |    ₹50,808.31 |
| West    |    ₹41,408.68 |
| Central |    ₹41,288.34 |
| South   |    ₹36,398.75 |

### Key Insight

**East** generated the highest revenue, while **South** generated the lowest revenue.

---

# 🛍️ Product Analysis

The project analyzes product-wise revenue and quantity.

The dataset contains the following products:

* Laptop
* Smartphone
* Desk
* Monitor
* Bookshelf
* Office Chair
* Keyboard
* Headphones
* Blender
* Coffee Maker

---

## 💰 Product Revenue Analysis

| Product      | Total Revenue |
| ------------ | ------------: |
| Laptop       |    ₹67,499.25 |
| Smartphone   |    ₹67,199.04 |
| Desk         |    ₹23,399.22 |
| Monitor      |    ₹21,999.12 |
| Bookshelf    |    ₹15,298.98 |
| Office Chair |    ₹10,399.48 |
| Keyboard     |     ₹8,009.11 |
| Headphones   |     ₹7,049.53 |
| Blender      |     ₹5,219.13 |
| Coffee Maker |     ₹3,119.61 |

### Key Insight

**Laptop generated the highest revenue**, followed closely by **Smartphone**.

---

# 📦 Quantity Analysis

Product quantity is analyzed using the Quantity field.

The project identifies:

* Total quantity sold
* Product-wise quantity
* Most frequently purchased products
* Product purchase counts

**Bookshelf recorded the highest quantity sold among the analyzed products.**

---

# 👥 Customer Analysis

Customer analysis is performed using:

* Customer ID
* Customer Name
* Customer Segment
* Customer Since
* Customer Months
* Customer Tenure
* Total Purchase

The workbook contains **50 unique customers**.

---

# 💎 High-Value Customer Analysis

The project includes a **High-Value Customer** analysis section.

Customer purchases are calculated using SUMIF.

Example:

```excel
=SUMIF(SalesData[Customer_ID],G18,SalesData[Total_Amount])
```

This helps identify customers contributing higher purchase amounts.

A customer can be classified using a threshold such as:

```excel
=IF(H26>=10000,"High Value","Normal")
```

The analysis supports the concept of customer grouping and segmentation.

---

# 🛒 Most Frequently Purchased Product

The project identifies products based on purchase frequency.

Purchase count is calculated using:

```excel
=COUNTIF('Raw Data'!F:F,G29)
```

Total quantity is calculated using:

```excel
=SUMIF('Raw Data'!F:F,G29,'Raw Data'!H:H)
```

This analysis helps determine which products are purchased most frequently.

---

# 🔁 Repeating Names

The project includes a **Repeating Names** analysis.

Purchase frequency is calculated using:

```excel
=COUNTIF('Raw Data'!D:D,G42)
```

This helps identify customers who appear multiple times in the transaction dataset.

The project also includes a separate **Repeating Customers** section showing how many times selected customer names appear.

---

# 🔎 Compare Two Lists

The workbook contains a section for comparing two customer lists.

The purpose is to identify names appearing in both lists.

Example:

```excel
=FILTER(G112:G121,COUNTIF(H112:H121,G112:G121)>0,"No Matching Names")
```

This demonstrates the use of:

* FILTER
* COUNTIF
* List comparison
* Multi-value extraction

---

# 🔤 Customer Name Abbreviations

The project creates abbreviations from customer names using TEXT functions.

Example:

```excel
=UPPER(LEFT(TEXTBEFORE(J26," "),1)&LEFT(TEXTAFTER(J26," "),1))
```

For example:

```text
Paul Baker → PB
Michael Brown → MB
Karen Hill → KH
```

This demonstrates the use of:

* TEXTBEFORE
* TEXTAFTER
* LEFT
* UPPER

---

# 🎯 What-If Analysis

The project includes **What-If Analysis** to understand how changes in sales assumptions affect revenue.

## Scenario Manager

Three scenarios are included:

```text
Normal
High Sales
Low Sales
```

Scenario values include:

| Scenario   |     Revenue |
| ---------- | ----------: |
| Normal     | ₹229,192.47 |
| High Sales |    ₹250,000 |
| Low Sales  |    ₹180,000 |

The Scenario Summary sheet provides a comparison of these scenarios.

---

# 🎯 Goal Seek Analysis

Goal Seek is used to determine the required value needed to reach a target amount.

Example:

```text
Quantity = 10
Target Revenue = ₹5,000
Required Unit Price = ₹500
```

The expected amount is calculated using:

```excel
=K8*K9
```

Goal Seek demonstrates how Excel can determine an unknown input based on a desired output.

---

# 📉 Linear Regression Analysis

Linear Regression is performed using the **Data Analysis ToolPak**.

The regression analyzes the relationship between:

```text
Quantity
       ↓
Total Amount
```

### Regression Statistics

| Metric            |  Value |
| ----------------- | -----: |
| Multiple R        | 0.4581 |
| R Square          | 0.2098 |
| Adjusted R Square | 0.2067 |
| Standard Error    | 919.61 |
| Observations      |    250 |

The regression helps understand the relationship between quantity purchased and total transaction amount.

---

# 📊 Pivot Table Analysis

The **Pivot Table** sheet summarizes the sales dataset.

Pivot Tables are used for:

* Region-wise revenue
* Product-wise revenue
* Product-wise quantity
* Sales performance
* Business comparisons

Example region analysis:

| Region  |    Revenue |
| ------- | ---------: |
| East    | ₹59,288.39 |
| North   | ₹50,808.31 |
| West    | ₹41,408.68 |
| Central | ₹41,288.34 |
| South   | ₹36,398.75 |

The Pivot Table provides a quick and flexible way to summarize the dataset.

---

# 📊 Data Visualization

The **Visualizations** sheet contains graphical representations of the analysis.

The project includes:

### 📊 Bar Chart

Used to compare revenue across products.

The chart makes it easier to identify the highest and lowest revenue-generating products.

---

### 📈 Line Chart

Used to display the monthly revenue trend.

The line chart helps identify changes in revenue over time.

---

### 🥧 Pie Chart

Used to display the distribution of revenue across regions.

The pie chart provides a quick visual comparison of regional contribution.

---

# 📊 Dashboard

The **Dashboard** sheet provides a centralized view of the sales performance.

The main KPI cards include:

```text
Total Revenue
Total Transactions
Total Quantity
Total Customers
```

The dashboard also contains a **Revenue Performance** section with:

```text
Actual Revenue
Target Revenue
Indicator
```

The target revenue is:

```text
₹250,000
```

The dashboard is designed to provide a quick overview of the overall sales performance.

---

# 🎨 Dashboard Design

The Dashboard focuses on:

* Clear KPI cards
* Simple visual hierarchy
* Sales performance indicators
* Revenue analysis
* Easy-to-understand charts
* Business-oriented presentation

The purpose is to allow users to understand the major sales insights without reviewing the complete dataset.

---

# 📋 Key Calculations

## Total Revenue

```excel
=SUM(SalesData[Total_Amount])
```

Result:

```text
₹229,192.47
```

---

## Total Transactions

```excel
=COUNTA(SalesData[Transaction_ID])
```

Result:

```text
250
```

---

## Total Quantity

```excel
=SUM(SalesData[Quantity])
```

Result:

```text
753
```

---

## Average Transaction Value

```excel
=AVERAGE(SalesData[Total_Amount])
```

Result:

```text
Approximately ₹916.77
```

---

## Region Revenue

```excel
=SUMIF(SalesData[Region],D14,SalesData[Total_Amount])
```

---

## Product Revenue

```excel
=SUMIF(SalesData[Product_Name],G22,SalesData[Total_Amount])
```

---

## Product Purchase Count

```excel
=COUNTIF('Raw Data'!F:F,G29)
```

---

## Product Quantity

```excel
=SUMIF('Raw Data'!F:F,G29,'Raw Data'!H:H)
```

---

# 📚 Excel Functions Covered

The project demonstrates a wide range of Excel functions.

### Date & Time Functions

```text
TODAY
NOW
DATEDIF
EOMONTH
EDATE
```

### Lookup & Reference Functions

```text
INDEX
MATCH
```

### Dynamic Array Functions

```text
FILTER
SORT
UNIQUE
```

### Logical Functions

```text
IF
```

### Calculation Functions

```text
SUM
AVERAGE
SUMIF
SUMIFS
COUNTIF
COUNTA
```

### Text Functions

```text
TEXTBEFORE
TEXTAFTER
LEFT
UPPER
```

---

# 📈 Data Storytelling

The project uses data storytelling to convert numerical analysis into meaningful business insights.

Important findings include:

1. **Laptop generated the highest revenue.**
2. **East region generated the highest revenue.**
3. **Smartphone was the second-highest revenue product.**
4. **Bookshelf had the highest quantity sold.**
5. **South region generated the lowest revenue.**

These insights help explain the overall sales performance of the business.

---

# 💡 Business Insights

### Product Insight

Laptop and Smartphone are the strongest revenue-generating products.

### Regional Insight

East is the strongest-performing region based on total revenue.

### Low-Performing Region

South generated the lowest regional revenue and may require additional sales attention.

### Customer Insight

Repeated transactions indicate that several customers make purchases multiple times.

### Sales Insight

The business generated total revenue of approximately **₹229K** from 250 transactions.

---

# 🔄 Project Workflow

The complete project follows this workflow:

```text
Raw Transaction Data
        ↓
Date & Time Analysis
        ↓
Formula-Based Analysis
        ↓
Customer & Product Analysis
        ↓
What-If Analysis
        ↓
Linear Regression
        ↓
Pivot Table
        ↓
Visualizations
        ↓
Dashboard
        ↓
Business Insights
```

---

# 📌 Project Output

The final workbook contains:

```text
Raw Data
      ↓
Pivot Table
      ↓
Scenario Summary
      ↓
Regression Summary
      ↓
Analysis
      ↓
Visualizations
      ↓
Dashboard
```

The workbook provides a complete Excel-based data analytics solution.

---

# 🧠 Concepts Practiced

```text
✔ Data Organization
✔ Excel Tables
✔ Date & Time Functions
✔ Dynamic Array Functions
✔ Conditional Formatting
✔ Timestamp Creation
✔ Customer Analysis
✔ Product Analysis
✔ Regional Analysis
✔ High-Value Customer Analysis
✔ Frequently Purchased Product Analysis
✔ Repeating Customer Analysis
✔ List Comparison
✔ Text Functions
✔ Scenario Manager
✔ Goal Seek
✔ Linear Regression
✔ Pivot Tables
✔ Data Visualization
✔ Dashboard Creation
✔ Data Storytelling
```

---

# 📚 Learning Outcomes

Through this project, the following skills were practiced:

* Excel Data Analysis
* Excel Formula Implementation
* Dynamic Array Functions
* Date & Time Analysis
* Customer Analysis
* Product Analysis
* Revenue Analysis
* Regional Analysis
* What-If Analysis
* Regression Analysis
* Pivot Table Creation
* Chart Creation
* Dashboard Development
* Conditional Formatting
* Business Data Interpretation
* Data Storytelling

---

# 💻 Skills Demonstrated

## Technical Skills

* Microsoft Excel
* Excel Tables
* Excel Formulas
* TODAY
* NOW
* DATEDIF
* EOMONTH
* FILTER
* SORT
* UNIQUE
* SUMIF
* SUMIFS
* COUNTIF
* INDEX-MATCH
* IF
* TEXTBEFORE
* TEXTAFTER
* Conditional Formatting
* Scenario Manager
* Goal Seek
* Data Analysis ToolPak
* Linear Regression
* Pivot Tables
* Charts
* Dashboard Design

## Analytical Skills

* Sales Analysis
* Revenue Analysis
* Product Performance Analysis
* Regional Performance Analysis
* Customer Analysis
* Customer Segmentation
* Purchase Frequency Analysis
* Trend Analysis
* Scenario Analysis
* Statistical Analysis
* Business Insight Generation
* Data Storytelling

---

# 🏆 Project Achievements

* Successfully analyzed **250 sales transactions**
* Calculated total revenue of **₹229,192.47**
* Analyzed **753 total quantities sold**
* Identified **50 unique customers**
* Implemented multiple Date & Time functions
* Implemented FILTER-based analysis
* Created dynamic timestamp functionality
* Applied Conditional Formatting with indicators
* Completed Scenario Manager analysis
* Performed Goal Seek analysis
* Completed Linear Regression using ToolPak
* Created Pivot Table analysis
* Created Bar, Line and Pie Charts
* Created a centralized Sales Dashboard
* Performed customer and product analysis
* Added business insights and data storytelling

---

# 📋 Dataset Fields

The Raw Data sheet contains the following fields:

```text
Transaction_ID
Date
Customer_ID
Customer_Name
Product_ID
Product_Name
Category
Quantity
Unit_Price
Payment_Method
Region
Customer_Segment
Customer_Since
Total_Amount
Customer Months
Month End
Today
Current Date Time
Customer Tenure
Timestamp
```

---

# 📊 Key Dataset Statistics

| Metric                    |       Value |
| ------------------------- | ----------: |
| Total Transactions        |         250 |
| Total Quantity            |         753 |
| Total Revenue             | ₹229,192.47 |
| Unique Customers          |          50 |
| Number of Products        |          10 |
| Number of Regions         |           5 |
| Number of Categories      |           3 |
| Number of Payment Methods |           4 |
| Regression Observations   |         250 |

---

# 🛍️ Product Categories

The dataset contains three main product categories:

```text
Electronics
Furniture
Appliances
```

These categories are used for product and sales analysis.

---

# 💳 Payment Methods

The dataset contains the following payment methods:

```text
Cash
Credit Card
Debit Card
PayPal
```

Payment methods are used to analyze customer purchasing behavior.

---

# 👤 Customer Segments

The dataset contains three customer segments:

```text
Premium
Standard
Basic
```

Customer segmentation helps understand differences in purchasing behavior.

---

# 🌍 Regions

The sales dataset contains five regions:

```text
East
North
West
Central
South
```

Regional analysis is used to compare revenue performance.

---

# 🔎 Business Analysis Areas

The project focuses on the following business questions:

### Sales Performance

What is the total revenue and average transaction value?

### Product Performance

Which products generate the highest revenue?

### Regional Performance

Which regions contribute the most revenue?

### Customer Performance

Which customers generate higher purchase amounts?

### Purchase Frequency

Which products are purchased most frequently?

### Customer Retention

Which customer names appear repeatedly in transactions?

### Revenue Forecasting

How does changing sales assumptions affect projected revenue?

### Statistical Relationship

Is there a relationship between quantity sold and total transaction amount?

---

# 🎯 What-If Business Questions

The What-If Analysis supports questions such as:

* What happens if sales increase?
* What happens if sales decrease?
* What revenue target can be achieved?
* What unit price is required to reach a target amount?
* How can changing assumptions affect projected revenue?

---

# 📉 Regression Business Question

The regression analysis investigates:

```text
Does Quantity Sold have a meaningful relationship
with Total Transaction Amount?
```

The analysis uses **250 observations** and provides regression statistics such as R Square, Multiple R, Standard Error, and significance information.

---

# 📊 Dashboard KPIs

The final Dashboard presents the main business KPIs:

```text
💰 Total Revenue
🧾 Total Transactions
📦 Total Quantity
👥 Total Customers
```

A Revenue Performance indicator is also included to compare actual revenue against the target.

---

# 📝 How to Use

1. Open the Excel workbook.
2. Start from the **Raw Data** sheet.
3. Review the transaction dataset.
4. Check the Date & Time analysis columns.
5. Open the **Pivot Table** sheet for summarized data.
6. Open **Scenario Summary** to review What-If scenarios.
7. Open **Summary** to view regression results.
8. Open **Analysis** to review formulas and business analysis.
9. Open **Visualizations** to view charts.
10. Open **Dashboard** to view the final sales dashboard.

---

# 📌 Project Summary

The **Sales & Data Intelligence Dashboard** is a practical Microsoft Excel project that transforms raw sales transaction data into meaningful business insights.

The project combines:

```text
Data Organization
        +
Excel Formulas
        +
Date & Time Analysis
        +
Dynamic Array Functions
        +
Customer Analysis
        +
Product Analysis
        +
What-If Analysis
        +
Linear Regression
        +
Pivot Tables
        +
Data Visualization
        +
Dashboard
        +
Data Storytelling
        =
Complete Sales Data Analytics Solution
```

The project demonstrates how Microsoft Excel can be used to analyze sales performance, revenue, customers, products, regions, purchasing patterns, and business scenarios.

---

# 👨‍💻 Author

**Sarth Thakar**
```
```
