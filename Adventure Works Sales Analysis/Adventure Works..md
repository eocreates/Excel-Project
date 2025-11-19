
Adventure Works
----
A professional and interactive Excel dashboard designed to analyze Adventure Works’ sales performance, product profitability, regional market trends, and customer purchasing behavior across four years of transactional data.

## Table of Contents
* Project Overview
* Context and Problem
* Dataset Description
* Data Cleaning
* Analysis
* Dashboard
* Recommendations

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
🔍 1. Yearly Profit Analysis
Dataset Table Preview
Year	Total Profit	Average Profit	Above Average
2010	£07 M	£06 M	£07 M
2011	£03 M	£06 M	—
2012	£09 M	£06 M	£09 M
2013	£07 M	£06 M	£07 M
2014	£06 M	£06 M	£06 M
2015	£04 M	£06 M	—
2016	£05 M	£06 M	—
2017	£04 M	£06 M	—
👉 Function Codes Used
Total Profit per Year
=SUMIFS(Profit, Year, A2)

Average Profit (All Years)
=AVERAGE(Profit_Range)

Above Average Flag
=IF([@TotalProfit] > AverageProfit, [@TotalProfit], "")

% of Profit from Highlighted Years
=SUM(HighlightedYearsProfit) / SUM(AllYearsProfit)


Result: 51.07% of total profit came from highlighted years

🔍 2. Revenue, Profit, and Units Sold by Year
Year	Total Revenue	Total Profit	Units Sold
2010	£19.19 M	£6.63 M	61.57K
2011	£11.13 M	£2.74 M	54.77K
2012	£31.90 M	£9.21 M	97.97K
2013	£20.33 M	£6.72 M	64.66K
2014	£16.63 M	£5.88 M	92.04K
2015	£12.43 M	£4.00 M	49.48K
2016	£12.37 M	£4.90 M	43.16K
2017	£13.37 M	£4.09 M	49.23K
Function Codes
Total Revenue
=SUMIFS(Revenue, Year, A2)

Units Sold
=SUMIFS(Units_Sold, Year, A2)

🔍 3. Monthly Profit Analysis
Month	Profit
Jan	£4M
Feb	£2M
Mar	£1M
Apr	£3M
May	£6M
Jun	£4M
Jul	£6M
Aug	£3M
Sep	£2M
Oct	£5M
Nov	£6M
Dec	£2M

Total = £44M

Function Code
=SUMIFS(Profit, MonthName, A2)

% Contribution of Highlighted Months
=SUM(Highlighted_Month_Profit) / SUM(All_Month_Profit)


Result: 39.32%

🔍 4. Weekday Profit Analysis
Day	Profit
Mon	£4.33M
Tue	£8.73M
Wed	£8.23M
Thu	£2.18M
Fri	£7.10M
Sat	£9.29M
Sun	£4.31M
Function Code
=SUMIFS(Profit, Weekday, A2)

Top 3 Weekdays
=LARGE(Profit_Range, 1)
=LARGE(Profit_Range, 2)
=LARGE(Profit_Range, 3)


Top performers:

Sat – £9.29M

Tue – £8.73M

Wed – £8.23M

🔍 5. Sales Channel Performance
Channel	Profit	Percentage
Offline	£24.92M	56.42%
Online	£19.25M	43.58%
Function Codes
Channel Profit
=SUMIFS(Profit, SalesChannel, A2)

Channel Percentage
=[@Profit] / SUM(TotalChannelProfit)

Dashboard

The Excel dashboard includes:

KPI cards (Revenue, Profit, Cost, Units Sold, Unit Price, Unit Cost)

Monthly profit line chart

Weekday profit bar chart

Year-over-year comparison

Product category & region breakdowns

Online vs. Offline donut chart

Interactive slicers (Year, Month, Region, Product Category)

Recommendations
✔ Product Strategy

Invest more in Cosmetics, Household, and Office Supplies

Improve pricing or reduce costs for low-margin categories

✔ Region Strategy

Expand in Asia, Europe, North America

Analyze underperforming regions for logistics or pricing barriers

✔ Seasonal Strategy

Target peak months (May, July, November)

Boost weekend promotions (Saturday highest profit day)

✔ Sales Channel Strategy

Increase investment in the Online channel

Improve offline partnerships and retail distribution

✔ Data Strategy

Add forecasting

Add customer segmentation (RFM)

Automate data refresh with Power Query
