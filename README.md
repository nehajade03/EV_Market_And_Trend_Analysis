# Electric Vehicles Market Analysis Dashboard

## Introduction

The **Electric Vehicles Analysis Dashboard** is an interactive Microsoft Excel-based tool designed to analyze trends in electric vehicle (EV) adoption. It provides insights into:

- Total number of electric vehicles.
- Distribution of EVs across different states.
- Vehicle model-wise distribution.
- Growth trends over the years.
- Clean Alternative Fuel Vehicle (CAFV) eligibility.

Leveraging Excel's **PivotTables**, **PivotCharts**, **slicers**, and **conditional formatting**, this dashboard allows real-time data filtering and interactive exploration.

## Objective

The primary goals of this dashboard are to:

- **Provide a comprehensive view** of electric vehicle adoption trends.
- **Enable interactive filtering** to analyze data by city, electric utility, and vehicle type.
- **Offer insightful visualizations** for understanding EV distribution, model popularity, and CAFV eligibility.
- **Facilitate data-driven decision-making** for businesses, policymakers, and researchers.
  
## Data Source

The dataset contains information about electric vehicle models, manufacturers, states, CAFV eligibility, and battery range. The data is structured and processed using Excel formulas, PivotTables, and dynamic filtering for smooth interactivity.

# 📊 EV Dataset – Full Formula Reference

## 1. City-Level Analysis

| Metric | Formula |
|--------|---------|
| **Total EVs by City** | `=COUNTIF(EV_Data!A:A, A3)` |
| **Electric Utility by City** | `=XLOOKUP(TRIM(CLEAN(A3)), EV_Data!A:A, EV_Data!M:M, "Not Found")` |

---

## 2. Summary Metrics (Top Section)

| Metric | Formula |
|--------|---------|
| **Total EVs** | `=COUNTA(EV_Data!D:D)` |
| **Count of BEVs** | `=COUNTIF(EV_Data!E:E, "BEV")` |
| **Count of PHEVs** | `=COUNTIF(EV_Data!E:E, "PHEV")` |
| **EV Makers (Unique)** | `=COUNTA(UNIQUE(EV_Data!H:H))` |
| **EV Models (Unique)** | `=COUNTA(UNIQUE(EV_Data!I:I))` |
| **Count by Specific Counties** | `=COUNTIF(EV_Data!C:C, "King")` *(example—adjust county name as needed)* |

---

## 3. CAFV Eligibility Analysis

| Metric | Formula |
|--------|---------|
| **CAFV Eligible** | `=COUNTIF(EV_Data!L:L, "Clean Alternative Fuel Vehicle Eligible")` |
| **CAFV Not Eligible** | `=COUNTIF(EV_Data!L:L, "Not eligible due to low battery range")` |

---

## 4. Year-Over-Year Growth Table

### **EV Registrations (Column H)**   
Formula if calculated manually per year:  
`=COUNTIF(EV_Data!D:D, G8)`

### **EV Count Difference (Column I)**  
=COUNTIF(EV_Data!D:D, G8) - COUNTIF(EV_Data!D:D, G7)

# 📊 Year-over-Year (YOY) Growth Formula Set  
(Using COUNTIF only — matches your original logic)

| Purpose | Formula |
|--------|----------|
| **Previous Year EV Count** | `=COUNTIF(EV_Data!D:D, G7)` |
| **EV Count Difference (YOY Change)** | `=COUNTIF(EV_Data!D:D, G8) - COUNTIF(EV_Data!D:D, G7)` |
| **YOY Percentage Growth** | `=((COUNTIF(EV_Data!D:D, G8) - COUNTIF(EV_Data!D:D, G7)) / COUNTIF(EV_Data!D:D, G7))` |


<img width="1910" height="950" alt="image" src="https://github.com/user-attachments/assets/105083e3-7bb7-42e4-b1fe-4b59676e18db" />


# 📄 **Summary Report**

### 🔹 **Top EV Cities**
Cities like **Seattle, Redmond, Bellevue, Sammamish, Issaquah, and Kirkland** dominate EV registrations.  
Most are served by **Puget Sound Energy** and **Seattle City Light**.

### 🔹 **EV Totals**
- **Total EVs:** 177,824  
- **BEVs:** 139,168  
- **PHEVs:** 38,656  

### 🔹 **Manufacturers & Models**
- **Unique Makers:** 42  
- **Unique Models:** 139  

### 🔹 **CAFV Eligibility**
- **Eligible:** 66,289  
- **Not Eligible:** 19,585

### 🔹 **Year-Over-Year Growth**
- Explosive early growth (2011–2013)  
- Consistent increases between 2015–2022  
- **2023 adds the largest EV volume jump: +29,811**  
- 2024 is negative due to incomplete data  

### 🔹 **Key Insight**
EV adoption is accelerating fastest in high-income, infrastructure-dense regions. Utilities serving these regions show strong alignment with EV growth patterns.

---

## Conditional Formatting Applied:
- **Headers** styled with bright green background and bold text.
- **Total EVs** column color-coded from yellow (low) to green (high) using gradient scale.
- **City-wise % Share of Total EVs** visualized with blue data bars for easier comparison.
- **Consistent black borders** for clean table presentation.

This manual method ensured cross-verification with automated VBA processes, enhancing both accuracy and report presentation quality.

## Pivot tables
![image](https://github.com/user-attachments/assets/3490bae0-0906-4926-9412-16fc41ea3178)

![image](https://github.com/user-attachments/assets/60c7939d-5d69-4d0a-8449-40d1d94101a6)

![image](https://github.com/user-attachments/assets/1e82b03f-3473-498b-b128-01baaa2086a6)




## Dashboard 
![image](https://github.com/user-attachments/assets/4d70cd9a-8e0b-4105-b908-e065361406a2)


## Dashboard Overview

The dashboard consists of multiple interactive sections, each focusing on different insights:

### 1. Key Performance Indicators (KPIs) Section

Located on the left side of the dashboard, this section displays summary metrics:
- **Total Vehicles**: Total count of electric vehicles.
- **Average Electric Range**: Mean range (in miles) for all EVs.
- **BEV & PHEV Distribution**: Pie charts displaying the proportion of Battery Electric Vehicles (BEV) vs. Plug-in Hybrid Electric Vehicles (PHEV).

### 2. Interactive Filter Panel

Positioned at the top, this panel allows users to filter data dynamically by:
- **Electric Utility Provider**: Filter by electricity provider (e.g., Avista Corp).
- **City**: Select a city (e.g., Aberdeen).
- **Electric Vehicle Type**: Filter by BEV or PHEV.

### 3. Total Vehicles by Model Year (Line Chart)

This chart shows the growth trend of EVs from 2010 onwards, illustrating:
- EV adoption over the years.
- Peaks and dips in adoption, highlighting significant years.
  
![image](https://github.com/user-attachments/assets/50d393fb-7819-4134-98a8-0b2c1354ef05)

### 4. Total Vehicles by State (Geographical Map)

This U.S. map visualizes the state-wise distribution of EVs using different shades of green to represent varying adoption levels.
![image](https://github.com/user-attachments/assets/f59ec464-fb3d-481f-9eee-7e6dd28f96c2)


### 5. Total Vehicles by Model Maker (Bar Chart)

This horizontal bar chart displays the most popular EV manufacturers. Tesla leads, followed by Nissan, Chevrolet, Ford, and BMW.
![image](https://github.com/user-attachments/assets/2e3bb61c-0541-4858-b3be-e1cdc0901db7)


### 6. Total Vehicles by CAFV Eligibility (Donut Chart)

Categorizes vehicles based on their eligibility for the Clean Alternative Fuel Vehicle (CAFV) program:
- **Eligible for CAFV Program**: 37%
- **Eligibility Unknown**: 11%
- **Not Eligible (Low battery range)**: 52%
  
![image](https://github.com/user-attachments/assets/73158c6a-1f82-4c59-8df5-758b36ef75f8)

### 7. Total Vehicles by Model (Treemap Visualization)

A treemap chart displays various EV models (e.g., Tesla Model 3, Nissan Leaf). The size of each box represents the number of vehicles of that model.

![image](https://github.com/user-attachments/assets/1515ab42-b7be-479a-8abb-3a1a31958d6b)


## Key Insights from the Dashboard

- **Tesla** dominates the EV market, leading in vehicle count.
- EV adoption has **increased significantly** over the years, with recent years showing a peak.
- Certain states have **higher EV adoption**, indicating better infrastructure support.
- A large percentage of EVs are **not CAFV eligible**, highlighting potential areas for improvement.

## Future Enhancements

- **Advanced Filtering Options**: Add filters for battery capacity, charging time, and energy efficiency.
- **Predictive Analytics**: Implement forecasting models to predict future EV adoption trends.
- **Expanded Data Sources**: Integrate government and industry datasets for deeper insights.
- **Automated Data Updates**: Link to external databases for real-time data refresh.

## Conclusion

The **Electric Vehicles Analysis Dashboard** in Excel provides a powerful and interactive tool for analyzing EV adoption trends. With dynamic filtering, clear visual insights, and user-friendly navigation, the dashboard serves as a valuable resource for:
- **Businesses** looking to understand EV market trends.
- **Policymakers** planning EV infrastructure.
- **Researchers & Analysts** studying EV adoption patterns.


# Electric Vehicle (EV) Report Generator – Excel VBA

This project contains an **Excel VBA tool** that generates automated reports for Electric Vehicle (EV) data. Users can generate **City** or **Year** reports by simply clicking a button.

## Features

### 1. Interactive Buttons
- **Generate City Report** – Enter a city to see EV data and insights.
- **Generate Year Report** – Enter a year to see EV registrations and YoY growth.
- Buttons are **locked and fixed**, ensuring consistent placement.

### 2. Automated Report
- Clears previous report content automatically.
- Calculates **total EVs**, top manufacturers, EV types, CAFV eligibility, and utility providers.
- Generates **YoY growth analysis** for year reports.
- Applies formatting: headers, borders, highlights, and trend indicators.

### 3. Data Insights
Automatically provides key insights, e.g.:
- Top manufacturer market share
- BEV vs PHEV split
- Top 3 makes share
- CAFV eligibility distribution
- Main utility provider

### 4. Benefits
- **Time-saving:** Instantly generates formatted reports without manual calculations.
- **Consistency:** Standardized formatting and calculations every time.
- **User-friendly:** No need for VBA knowledge; just click a button.

## How It Works
1. Open `EV_Report.xlsm`.
2. Click **Generate City Report** or **Generate Year Report**.
3. Enter the city or year when prompted.
4. The report will automatically populate below the buttons.

### Buttons
<img width="625" height="62" alt="image" src="https://github.com/user-attachments/assets/412fc7fb-c951-4c82-8f66-d3917a7a5470" />


### Sample City Report
<img width="1526" height="656" alt="image" src="https://github.com/user-attachments/assets/58bebe06-78bc-4b98-8d8a-5a13067b89de" />

### Sample Year Report
<img width="1123" height="755" alt="image" src="https://github.com/user-attachments/assets/a8a957d0-e06c-4e0f-9972-e705879586cc" />

## VBA Code Highlights
- `SetupReportButtons()` – Creates and positions buttons.
- `GenerateReport()` – Core function that creates tables, insights, and YoY analysis.
- `City_Report()` / `Year_Report()` – Input validation and triggers report generation.

## Features
- Top selling city
- Most popular EV make
- Most popular EV model
- Year-wise vehicle distribution

## How to Use
1. Place your EV data in an Excel sheet named `EV_Data`.
2. Make sure the columns are:
   - Column A: City
   - Column D: Model Year
   - Column E: Make
   - Column F: Model
3. Run the `GenerateEVReport` macro.
4. The report will be generated in a new sheet called `Report`.

## Requirements
- Microsoft Excel
- VBA enabled

---
*Created by Neha Jade*


