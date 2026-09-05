```
# 📊 Global E-Commerce & Retail Performance Dashboard (2021 - 2024)

An institutional-grade, end-to-end **E-Commerce & Retail Analytics Dashboard** designed with a modern **Dark-Mode UI** in Microsoft Excel. The solution utilizes **Power Pivot Data Modeling (Star Schema)**, custom **DAX measures**, and optimized visual layouts to track high-level financial KPIs, revenue growth trends, regional profitability, and product line distributions.

---

## 📸 Dashboard Preview

![Dashboard Preview](dashboard_preview.png)

---

## 🎯 Executive Summary & Core Business Objectives

This project was developed to deliver executive-level decision support for a global retail entity. By unifying multi-table relational transaction databases into a unified analytical environment, the dashboard provides interactive answers to critical business questions regarding profitability, cost structures, product category performance, and multi-year growth trajectories.

### Key Performance Summary (All-Time Baseline):
* **Total Revenue:** $5,284,387.70
* **Total Profit:** $1,437,638.31
* **Total Cost:** $3,846,749.39
* **Profit Margin %:** 27.21%
* **Total Orders:** 10,000
* **Average Order Value (AOV):** $528.44
* **Total Active Customers:** 5,348

---

## 🏗️ Data Architecture & Star Schema Model

The analytics layer is driven by an enterprise **Star Schema Data Model** built inside **Excel Power Pivot**, ensuring maximum query efficiency, data integrity, and dynamic cross-filtering capabilities.


```

```
              ┌─────────────────┐
              │  Dim_Customers  │
              └────────┬────────┘
                       │ 1
                       │

```

┌──────────────┐   * ┌────┴─────┐ *   ┌──────────────┐
│ Dim_Products ├─────┤Fact_Sales├─────┤ Dim_Geography│
└──────────────┘ 1   └────┬─────┘ 1   └──────────────┘
│
│ *
┌────────┴────────┐
│    Dim_Date     │
└─────────────────┘

```

### Table Relationships & Cardinality:
1. **`Fact_Sales`** *(Fact Table)*: Contains 10,000 transactional line records including order IDs, dates, quantities, unit prices, revenue, profit, and foreign key references.
2. **`Dim_Customers`** *(Dimension)*: Stores unique customer IDs (5,348 records), customer segments (New, Premium, Regular, VIP), and demographics. Linked to `Fact_Sales[Customer_ID]` (1-to-Many).
3. **`Dim_Products`** *(Dimension)*: Catalog of items (70 products) across 5 primary categories. Linked to `Fact_Sales[Product_ID]` (1-to-Many).
4. **`Dim_Geography`** *(Dimension)*: Regional clusters (North America, Europe, Asia, Middle East) and countries. Linked to `Fact_Sales[Country_ID]` (1-to-Many).
5. **`Dim_Date`** *(Dimension)*: Date lookup table supporting Time Intelligence DAX logic. Linked to `Fact_Sales[Order_Date]` (1-to-Many).

---

## 🧮 Complete DAX Measures Dictionary

All core metrics were calculated using explicit **DAX (Data Analysis Expressions)** inside the Data Model rather than implicit pivot fields:

### 1. Financial Metrics
* **Total Revenue:**
  ```dax
  Total Revenue := SUM('Fact_Sales'[Sales_Amount])

```

* **Total Cost:**
```dax
Total Cost := SUM('Fact_Sales'[Cost_Amount])

```


* **Total Profit:**
```dax
Total Profit := SUM('Fact_Sales'[Profit_Amount])

```


* **Profit Margin %:**
```dax
Profit Margin % := DIVIDE([Total Profit], [Total Revenue], 0)

```



### 2. Operational Metrics

* **Total Orders:**
```dax
Total Orders := DISTINCTCOUNT('Fact_Sales'[Order_ID])

```


* **Total Customers:**
```dax
Total Customers := DISTINCTCOUNT('Fact_Sales'[Customer_ID])

```


* **Average Order Value (AOV):**
```dax
AOV := DIVIDE([Total Revenue], [Total Orders], 0)

```



### 3. Time Intelligence & Growth Measures

* **Prior Year Revenue:**
```dax
Prior Year Revenue := CALCULATE([Total Revenue], SAMEPERIODLASTYEAR('Dim_Date'[Date]))

```


* **YoY Growth %:**
```dax
YoY Growth % := 
VAR CurrentRevenue = [Total Revenue]
VAR PriorRevenue = [Prior Year Revenue]
RETURN
IF(
    ISBLANK(PriorRevenue) || PriorRevenue = CurrentRevenue,
    BLANK(),
    DIVIDE(CurrentRevenue - PriorRevenue, PriorRevenue, 0)
)

```



---

## 🎨 Dashboard Visual Layout & Charts Breakdown

The dashboard UI was engineered on a structured 12-column grid system using a custom **Dark-Mode theme** designed for executive reading.

```
┌────────────────────────────────────────────────────────────────────────┐
│                        TITLE BLOCK & HEADER                            │
├────────────────────────────────────────────────────────────────────────┤
│                      HEADERLESS DYNAMIC SLICERS                        │
│ [Region Slicer]       [Category Slicer]        [Customer Segment]      │
├───────────────────┬────────────────────────────────────────────────────┤
│   KPI SUMMARY     │               CHART 1: REVENUE TREND               │
│   CARD MATRIX     │         Line Chart (2021 - 2024 Time Series)       │
│ (3x3 Grid Layout) │                                                    │
├───────────────────┴─────────────────┬──────────────────────────────────┤
│    CHART 2: REGIONAL COMPARISON     │     CHART 3: TOP CATEGORIES      │
│ Clustered Column (Revenue vs Profit)│    Horizontal Bar (Revenue)      │
└─────────────────────────────────────┴──────────────────────────────────┘

```

### Visual Components & Chart Specifications:

1. **Executive KPI Card Grid (Top Left):**
* **Structure:** 9 synchronized visual summary blocks linking directly to underlying DAX Pivot cells.
* **Metrics:** Revenue, Profit, Margin %, Orders, AOV, Cost, Customers, Prior Year Revenue, and YoY Growth %.


2. **Chart 1: Revenue Trend over Time (2021 - 2024)**
* **Visual Type:** 2D Line Chart with Data Markers.
* **X-Axis:** Calendar Years (`Dim_Date[Year]`).
* **Y-Axis:** Total Revenue ($).
* **UI Treatment:** Clean line styling, hidden Y-axis gridlines, direct Data Labels enabled above data markers for enhanced readability.


3. **Chart 2: Revenue vs Profit by Region (Bottom Left)**
* **Visual Type:** Clustered Column Chart.
* **X-Axis:** Geographic Region (`Dim_Geography[Region]`).
* **Y-Axis:** Dual metric series (`Total Revenue` vs `Total Profit`).
* **Color Palette:** Deep navy blue (`#0F2537`) for Revenue contrasted with forest green (`#273E2B`) for Profit.


4. **Chart 3: Top Categories by Revenue (Bottom Right)**
* **Visual Type:** Clustered Horizontal Bar Chart.
* **Y-Axis:** Product Category (`Dim_Products[Category]`).
* **X-Axis:** Total Revenue ($).
* **Sorting:** Sorted in descending order to immediately draw focus to top revenue generators (e.g., *Electronics* and *Home & Kitchen*).



---

## 📁 Repository File Structure

```
├── Global_Ecommerce_Dashboard.xlsx     # Fully functional Excel workbook (Data Model, DAX, Dark Mode Dashboard)
├── Global_Ecommerce_Dashboard.pdf      # High-resolution PDF export (Formatted Fit-to-Page Layout)
├── dashboard_preview.png               # Visual asset preview for documentation
└── README.md                           # Comprehensive technical documentation

```

---

## 👤 Author & Contact

* **Mr. Fathallah Saied** — *Data Analyst*
* **LinkedIn:** [Fathallah Saied Profile](https://www.linkedin.com/)
* **GitHub Portfolio:** [GitHub Repositories](https://github.com)

---

*If you find this project structure useful for your portfolio, feel free to star ⭐️ this repository!*

```
```
