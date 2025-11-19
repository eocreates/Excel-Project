
Adventure Works
----
A professional and interactive Excel dashboard designed to analyze Adventure Works’ sales performance, product profitability, regional market trends, and customer purchasing behavior across four years of transactional data.

## Table of Contents
* [Project Overview](#Project-Overview)
* [Context and Problem](#Context-and-Problem)
* [Dataset Description](#Dataset-Description)
* [Data Cleaning](#Data-Cleaning)
* [Analysis](#Analysis)
* [Dashboard](#Dashboard)
* [Recommendations](#Recommendations)

## Project Overview
This project presents a comprehensive sales and profit analysis based on four years of transactional data from the Adventure Works dataset.
Using Microsoft Excel (PivotTables, PivotCharts, Power Query, and Data Modeling), the project delivers an interactive dashboard that helps business stakeholders understand revenue drivers, product performance, customer geography trends, and time-based patterns.
The primary focus is profit analysis, with a deep exploration of products, customer locations, and time trends.

## Context and Problem
The Adventure Works management team wants to evaluate four years of performance to understand:
* Which products contribute most to profit
* Which countries/regions generate the highest profitability
* How profit changes across months, weekdays, and years
* Whether there are any noticeable patterns or seasonal trends
* How sales channels (online vs offline) influence performance
The business goal is to transform raw transactional data into insights that support decision-making, strategic planning, and operational improvements.

## Dataset Description
* This project uses a structured dataset extracted from the [Adventure Works company’s](https://github.com/eocreates/Excel-Project/blob/main/Adventure%20Works%20Sales%20Analysis/AdventureWorks%20Dataset.xlsx) multi-year sales and transactional records. The dataset includes detailed information on products, customer regions, sales channels, pricing, and profit metrics.


## Data Cleaning
* Power Query was used  to clean, transform, and load the raw Adventure Works dataset into Excel. This included removing errors, standardizing data types, formatting dates, creating calculated fields, and preparing a structured data model for analysis and dashboard reporting.

## Analysis
* Yearly Profit Analysis
````Excel
Total Profit per Year
=SUMIFS(Profit, Year, A2)

Average Profit (All Years)
=AVERAGE(Profit_Range)

Above Average Flag
=IF([@TotalProfit] > AverageProfit, [@TotalProfit], "")

% of Profit from Highlighted Years
=SUM(HighlightedYearsProfit) / SUM(AllYearsProfit)
`````
  - Result: 51.07% of total profit came from highlighted years

* Revenue, Profit, and Units Sold by Year
````Excel
Total Revenue
=SUMIFS(Revenue, Year, A2)

Units Sold
=SUMIFS(Units_Sold, Year, A2)
````

* Monthly Profit Analysis
````Excel
=SUMIFS(Profit, MonthName, A2)

% Contribution of Highlighted Months
=SUM(Highlighted_Month_Profit) / SUM(All_Month_Profit)
````
  - Result: 39.32%

* Weekday Profit Analysis
````Excel
=SUMIFS(Profit, Weekday, A2)
````
* Top 3 Weekdays
````Excel
=LARGE(Profit_Range, 1)
=LARGE(Profit_Range, 2)
=LARGE(Profit_Range, 3)
````
* Sales Channel Performance
````Excel
Channel Profit
=SUMIFS(Profit, SalesChannel, A2)

Channel Percentage
=[@Profit] / SUM(TotalChannelProfit)
````

## Dashboard
 <a href="https://github.com/eocreates/Excel-Project/blob/main/Adventure%20Works%20Sales%20Analysis/ADVN.png">
  <img src="https://github.com/eocreates/Excel-Project/blob/main/Adventure%20Works%20Sales%20Analysis/ADVN.png" width="1000">
</a>


## Recommendations
* Product Strategy
  - Invest more in Cosmetics, Household, and Office Supplies
  - Improve pricing or reduce costs for low-margin categories

* Region Strategy
  - Expand in Asia, Europe, North America
  - Analyze underperforming regions for logistics or pricing barriers

 * Seasonal Strategy
   - Target peak months (May, July, November)
   - Boost weekend promotions (Saturday highest profit day)

* Sales Channel Strategy
  - Increase investment in the Online channel
  - Improve offline partnerships and retail distribution


