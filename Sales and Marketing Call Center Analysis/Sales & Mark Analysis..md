📞 Sales & Marketing Call Center Dashboard
------
 A professional and interactive Excel dashboard for analyzing call center performance, agent activity, and customer behavior.

## Table of Contents
  * Project Overview
  * Context and Problem
  * Dataset Description
  * Data Cleaning
  * Analysis
  * Dashboard
  * Recommendations

## Project Overview
This project presents an interactive Power BI dashboard designed to help Sales & Marketing Call Center teams track performance metrics, understand customer behavior, and evaluate agent effectiveness.
The dashboard enables business leaders to make data-driven decisions by surfacing trends, inefficiencies, and optimization opportunities.

## Context and Problem
A call center receives thousands of customer calls daily across different agents, products, time periods, and regions.
However, raw call logs alone make it difficult to answer critical questions such as:
  * How many calls were successful, failed, or abandoned?
  * Which agents are performing well or poorly?
  * What products generate the most inbound calls?
  * What regions call the most frequently?
  * Why are customers abandoning calls?
  * How do monthly call trends fluctuate over time?
    
Leaders needed a dashboard that transforms raw call data into clear operational insights, supports agent-specific drilldowns, and highlights areas needing improvement.

## Dataset Description
This project uses a structured dataset extracted from the call center’s daily operational logs.[ Download here](https://github.com/eocreates/Excel-Project/blob/main/Sales%20and%20Marketing%20Call%20Center%20Analysis/Dataset%20Sales%20and%20Marketing%20Call%20Center.xlsx)
The dataset serves as the foundation for all analysis, KPIs, and dashboard visualizations.

## DATA CLEANING
The following steps were applied before analysis:
* Standardized Category Values
 - Regions (e.g., “London” vs “london”)
* Product Names
 - Outcome labels (“Success”, “Successful”, etc.)
* Converted Data Types
 - Dates → Excel Date format
* Call Duration → Numeric
* Ratings → Numeric
* Time → Extracted "Morning", "Afternoon", "Evening"
*  Added Derived Columns
 - Month = TEXT(Date, "MMM")
 - Outcome_Flag = IF(Call_Outcome="Successful",1,0)
 - Time Slot = IF(HOUR(Time)<12,"Morning",IF(HOUR(Time)<17,"Afternoon","Evening"))

## ANALYSIS
The Analysis sheet computes all KPIs used in the dashboard.
Below are the major analytical components:

* Total Calls
```EXCEL
=COUNTA(Data[Call_ID])
```
* Successful Calls
```EXCEL
=COUNTIF(Data[Call_Outcome],"Successful")
```
* Failed Calls
```EXCEL
=COUNTIF(Data[Call_Outcome],"Failed")
```
* Abandoned Calls
```EXCEL
=COUNTIF(Data[Call_Outcome],"Abandoned")
```
* Calls by Agent
```EXCEL
=COUNTIF(Data[Agent Full Name], AgentNameCell)
```
* Agent Success Rate
```EXCEL
= Successful_Calls_By_Agent / Total_Calls_By_Agent
```
* Average Rating
```EXCEL
=AVERAGEIF(Data[Agent Full Name], AgentName, Data[Rating])
```
* Reasons for Abandonment
```EXCEL
=COUNTIFS(Data[Call_Outcome],"Abandoned", Data[Reason],"Long Wait Time")
```
* Monthly Call Trends
```EXCEL
=COUNTIF(MonthColumn, "Jan")
```

VIEW 1 — Agent Performance Summary
```EXCEL
Agent_Name,
Total_Calls        = COUNTIF(Data[Agent], Name)
Successful_Calls   = COUNTIFS(Data[Agent], Name, Data[Outcome],"Successful")
Failed_Calls       = COUNTIFS(Data[Agent], Name, Data[Outcome],"Failed")
Abandoned_Calls    = COUNTIFS(Data[Agent], Name, Data[Outcome],"Abandoned")
Avg_Rating         = AVERAGEIF(Data[Agent], Name, Data[Rating])
Avg_Duration       = AVERAGEIF(Data[Agent], Name, Data[Call Duration])
Success_Rate       = Successful_Calls / Total_Calls
```
VIEW 2 — Product Call Summary
```EXCEL
Product,
Success_Count  = COUNTIFS(Data[Product], Product, Data[Outcome],"Successful")
Fail_Count     = COUNTIFS(Data[Product], Product, Data[Outcome],"Failed")
```

VIEW 3 — Abandonment Reason Summary
````EXCEL
Reason,
Total_Abandoned = COUNTIFS(Data[Call_Outcome],"Abandoned", Data[Reason], Reason)
````

VIEW 4 — Region-Level Call Summary
```EXCEL
Region,
Call_Count = COUNTIF(Data[Region], Region)
```

VIEW 5 — Monthly Trend Summary
`````EXCEL
Month,
Total_Calls = COUNTIF(Month_Column, Month)
`````

##  Dashboard
 <a href="https://github.com/eocreates/Excel-Project/blob/main/Sales%20and%20Marketing%20Call%20Center%20Analysis/Sales.png">
  <img src="https://github.com/eocreates/Excel-Project/blob/main/Sales%20and%20Marketing%20Call%20Center%20Analysis/Sales.png" width="1000">
</a>


## RECOMMENDATIONS
* Reduce Long Wait Times
 - Since long wait time is a major abandonment reason, improve staffing during peak periods.
* Improve Underperforming Agents
 - Agents with low success rates or poor ratings need coaching.
* Enhance Technical Stability
 - Technical issue–related abandonment needs IT evaluation.
* Monitor Product Inquiry Patterns
 - Products with high failure rates (e.g., Loans, Internet Packages) may require script updates.





