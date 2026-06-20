<div align="center">

# 📊 Excel Sales Analysis Dashboard

### 🚀 Advanced Microsoft Excel Project for Sales Analytics & Business Intelligence



<br>

<a href="#overview">
<img src="https://img.shields.io/badge/📖_Overview-Click_Here-blue?style=for-the-badge"/>
</a>

<a href="#project-features">
<img src="https://img.shields.io/badge/✨_Features-Explore-success?style=for-the-badge"/>
</a>

<a href="#dashboard">
<img src="https://img.shields.io/badge/📊_Dashboard-View-orange?style=for-the-badge"/>
</a>


<a href="#key-findings">
<img src="https://img.shields.io/badge/📈_Insights-Check-purple?style=for-the-badge"/>
</a>



</div>

---

## 📌 Table of Contents

- [🎯 Overview](#overview)
- [✨ Project Features](#project-features)
- [📁 Dataset](#dataset)
- [📅 Date & Time Functions](#date-time-functions)
- [🔍 Filter Analysis](#filter-analysis)
- [⏰ Timestamp Creation](#timestamp-creation)
- [🎯 What-If Analysis](#what-if-analysis)
- [📈 Linear Regression Analysis](#linear-regression-analysis)
- [👥 High Value Customer Analysis](#high-value-customer-analysis)
- [🏆 Most Purchased Product](#most-purchased-product)
- [🤝 Customer Matching Analysis](#customer-matching-analysis)
- [✂️ Text Function Analysis](#text-function-analysis)
- [📊 Dashboard](#dashboard)
- [🎓 Key Findings](#key-findings)
- [🛠 Technologies Used](#technologies-used)
- [📋 Project Structure](#project-structure)


---

<a id="overview"></a>

## 🎯 Overview

**Excel Sales Analysis Dashboard** is a comprehensive Microsoft Excel project designed to transform raw sales transaction data into meaningful business insights.

### 🎯 Objectives

- Analyze Customer Behavior
- Identify High Value Customers
- Discover Top Selling Products
- Perform Statistical Analysis
- Build Interactive Dashboards
- Generate Business Insights

---

<a id="project-features"></a>

## ✨ Project Features

- Date & Time Functions

- Dynamic FILTER Function

- Scenario Manager

- Goal Seek

- Linear Regression Analysis

- Customer Segmentation

- Most Purchased Product Analysis

- Customer Matching Analysis

- Pivot Tables & Charts

- Interactive Dashboard

---

<a id="dataset"></a>

## 📁 Dataset

The dataset contains sales transaction records including:

| Column | Description |
|----------|-------------|
| Transaction ID | Unique Transaction Number |
| Customer ID | Customer Identifier |
| Customer Name | Customer Name |
| Product ID | Product Code |
| Product Name | Product Purchased |
| Category | Product Category |
| Quantity | Quantity Purchased |
| Unit Price | Product Price |
| Payment Method | Payment Type |
| Region | Sales Region |
| Customer Segment | Customer Classification |
| Date | Transaction Date |
| Total Amount | Final Sales Amount |

---

<a id="date-time-functions"></a>

## 📅 Date & Time Functions

### Functions Used

```excel
=TODAY()
=NOW()
=EOMONTH()
=DATEDIF()
```

### Results

| Function | Output |
|----------|---------|
| TODAY() | 20-06-2026 |
| NOW() | 20-06-2026 13:28 |
| EOMONTH() | 30 June 2026 |
| DATEDIF() | 10 |

---

<a id="filter-analysis"></a>

## 🔍 Filter Analysis

### Requirement

Find transactions where:

- Region = East
- Payment Method = Cash

### Formula

```excel
=FILTER(A:N,(J:J="Cash")*(K:K="East"))
```

### Sample Output

- TRX0239 → Patricia Moore → Smartphone
- TRX0220 → Michael Brown → Bookshelf
- TRX0093 → Julie Foster → Blender

---

<a id="timestamp-creation"></a>

## ⏰ Timestamp Creation

### Formulas

```excel
=NOW()-1
=NOW()+1
=NOW()+2
```

### Results

| Event | Timestamp |
|--------|---------|
| Yesterday | 20-06-2026 13:28 |
| Tomorrow | 21-06-2026 13:28 |
| After Tomorrow | 22-06-2026 13:28 |

---

<a id="what-if-analysis"></a>

## 🎯 What-If Analysis

### Scenario Manager

| Quantity | Unit Price | Total |
|----------|------------|--------|
| 3 | 249.99 | 749.97 |
| 2 | 249.99 | 499.98 |
| 5 | 149.99 | 749.95 |
| 3 | 899.99 | 2699.97 |

### Goal Seek

#### Target

```text
Total Sales = ₹10,000
```

#### Result

```text
Required Unit Price = ₹3333.33
```

---

<a id="linear-regression-analysis"></a>

## 📈 Linear Regression Analysis

### Regression Statistics

| Metric | Value |
|----------|---------|
| Multiple R | 0.8354 |
| R Square | 0.6979 |
| Adjusted R Square | 0.6967 |
| Standard Error | 151.64 |
| Observations | 249 |

### Interpretation

📊 Approximately 70% of variation is explained by the model.

📈 Strong positive relationship between variables.

---

<a id="high-value-customer-analysis"></a>

## 👥 High Value Customer Analysis

### Top Customers

| Customer | Revenue |
|-----------|---------|
| Mark Carter | ₹15,659.65 |
| Edward Mitchell | ₹11,919.77 |
| Barbara Young | ₹10,649.80 |
| Patricia Moore | ₹9,799.73 |

### Grand Total

```text
₹93,867.66
```

---

<a id="most-purchased-product"></a>

## 🏆 Most Purchased Product

### Result

```text
📚 Bookshelf
```

### Insight

Bookshelf was the most frequently purchased product.

---

<a id="customer-matching-analysis"></a>


### Formula Used

```excel
=MATCH()
=XLOOKUP()
=FILTER()
```

---

<a id="text-function-analysis"></a>

## ✂️ Text Function Analysis

### Formula

```excel
=LEFT(A2,1)
```

### Output

| Segment | Short Form |
|----------|------------|
| Basic | B |
| Premium | P |
| Standard | S |

---

<a id="dashboard"></a>

## 📊 Dashboard

### Dashboard Components

<img width="964" height="594" alt="image" src="https://github.com/user-attachments/assets/b1a973e2-d7d9-4d7c-bfb1-7241f98a90ea" />


---

<a id="key-findings"></a>

## 🎓 Key Findings

### 💰 Revenue Insights

High-value customers contribute significantly to overall revenue.

### 👥 Customer Insights

Repeat customers generate consistent business growth.

### 📦 Product Insights

Bookshelf is the most purchased product.

### 📈 Statistical Insights

Regression analysis shows strong predictive capability.

---

<a id="technologies-used"></a>

## 🛠 Technologies Used

| Technology | Purpose |
|------------|----------|
| Microsoft Excel | Data Analysis |
| FILTER Function | Dynamic Filtering |
| Pivot Tables | Aggregation |
| Pivot Charts | Visualization |
| Goal Seek | Forecasting |
| Scenario Manager | What-If Analysis |
| Linear Regression | Statistical Modeling |

---

<a id="project-structure"></a>

## 📋 Project Structure

```text
Excel_Sales_Analysis_Dashboard/
│
├── Dataset.xlsx
├── Date_Time_Functions.xlsx
├── Filter_Analysis.xlsx
├── Timestamp_Creation.xlsx
├── What_If_Analysis.xlsx
├── Regression_Analysis.xlsx
├── Customer_Analysis.xlsx
├── Dashboard.xlsx
└── README.md
```

---


