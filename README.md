# NYC Market Dynamics Dashboard (Interactive Excel BI)

## 📌 Executive Summary
This project demonstrates end-to-end data analysis and Business Intelligence (BI) reporting using **Advanced Excel**. Utilizing a dataset of nearly 50,000 raw real estate records, this project involved handling missing values, engineering a custom "Estimated Minimum Revenue" proxy, and building a dynamic aggregation engine using Pivot Tables. The final deliverable is an interactive, stakeholder-ready UI that allows users to seamlessly drill down into neighborhood-specific pricing power and market yield.

## 🎯 Business Objective
To analyze high-volume market supply, dynamic pricing trends, and revenue concentration across different geographic constraints and product categories. The goal was to build an automated reporting tool that identifies the most profitable segments of a market to inform yield optimization strategies.

## 📊 Dataset
* **Source:** Kaggle (New York City Airbnb Open Data)
* **Volume:** 48,895 rows of raw listing data
* **Dimensions:** Geography (Boroughs), Product Category (Room Types), Pricing, and Engagement Metrics.

## 🛠️ Methodology & Technical Execution

### 1. ETL & Data Cleaning
* Transformed raw CSV data into a dynamic relational table.
* Handled null values in engagement metrics using bulk replacement logic to ensure data integrity for calculations.

### 2. Feature Engineering (The Revenue Proxy)
Raw pricing data does not equate to actual business yield. To solve this, I engineered a new metric called `Estimated_min_REV` using the formula `=[@price]*[@minimum_nights]`. This allowed for accurate modeling of the minimum baseline revenue generated per listing.

### 3. Data Modeling & Aggregation Engine
Built a hidden backend "Engine Room" using cross-linked Pivot Tables to aggregate massive row data into three distinct business views:
* **Market Supply:** Count of active listings grouped by Neighborhood.
* **Pricing Power:** Average nightly rate commanded by each Neighborhood.
* **Estimated Yield:** Total estimated revenue grouped by Room Category.

### 4. Interactive Dashboard (UI/UX)
Designed a front-facing, gridless BI dashboard containing three primary visualizations:
* **Market Supply Distribution** (Bar Chart)
* **Average Nightly Rate Trends** (Line Chart)
* **Revenue Yield by Room Category** (Pie/Donut Chart)

**💡 Technical Workaround (Cloud Accessibility):** To ensure this dashboard remained fully functional and interactive in cloud-based web environments (like Excel Online, which limits traditional macro-based Slicer UI), I engineered the interactivity using native PivotTable Report Filters. This maintains dynamic drill-down capabilities, allowing stakeholders to filter the entire dashboard by specific boroughs and room types without breaking the web architecture.

## 🚀 Skills Demonstrated
* **Tools:** Advanced Microsoft Excel (Office 365)
* **Data Engineering:** ETL, Data Cleaning, Feature Engineering, Calculated Fields
* **Data Analysis:** Yield Optimization, Trend Analysis, Aggregation Models
* **Visualization:** Dashboard Design, PivotCharts, Interactive BI Reporting
