# Project Overview – Excel Business Operations Analysis

## 1. Project Introduction

This project focuses on analysing business operational data using Microsoft Excel to generate meaningful insights, monitor operational performance, and support data-driven decision-making.

The project demonstrates an end-to-end Excel-based analytics workflow, covering data preparation, formula-based transformations, KPI calculations, Pivot Table analysis, and interactive dashboard development.

The primary objective is to transform raw operational data into structured reports and visual insights that help management understand business performance across different accounts and reporting periods.

## 2. Business Objective

The dashboard is designed to help business stakeholders monitor and analyse operational performance through the following key objectives:

- Monitor overall claims volume and performance.
- Analyse approved, PUA, and declined claims.
- Track monthly operational trends.
- Evaluate account-wise performance.
- Monitor Customer Success Manager (CSM) query volumes.
- Calculate operational KPIs and performance percentages.
- Identify changes in workload across different reporting periods.
- Provide management with an interactive reporting interface.

## 3. Dataset Overview

The project uses an operational dataset containing 18,325 records and 20 columns.

| Dataset Attribute | Details |
|---|---|
| Total Records | 18,325 |
| Total Columns | 20 |
| Unique Account Names (Raw) | 94 |
| Unique Account Names (Normalised) | 83 |
| Reporting Start Date | January 2024 |
| Reporting End Date | April 2025 |
| Total Claims | 808,002 |
| Approved Claims | 666,762 |
| PUA Claims | 102,154 |
| Declined Claims | 27,038 |
| CSM Queries | 11,138 |

The dataset contains account-level operational information, reporting dates, claims volumes, claim outcomes, and CSM query information.

## 4. Tools and Excel Functionalities

### A. Data Preparation and Transformation

- Excel Structured Tables.
- Calculated columns.
- Date and text transformations.
- Data type validation.
- Account-wise aggregations.
- Monthly data aggregations.
- Formula-based data checks.

### B. Advanced Excel Functions

The project demonstrates practical applications of:

| Excel Function | Application |
|---|---|
| SUMIFS | Account-wise and monthly aggregations |
| IF | Conditional calculations |
| IFERROR | Error handling and percentage calculations |
| TEXT | Extracting year and month from dates |
| VALUE | Numeric conversion |
| DATE | Date-based conditions |
| ISNUMBER | Data type validation |
| GETPIVOTDATA | Retrieving values from Pivot Tables |

### C. Pivot Table Analysis

Pivot Tables are used to summarise operational data and analyse:

- Total claims by account.
- Monthly claims performance.
- Approved claims.
- PUA claims.
- Declined claims.
- CSM query volumes.
- Account-level performance comparisons.

### D. Interactive Dashboard

The dashboard includes:

- Operational KPI cards.
- Monthly claims trend analysis.
- PUA trend analysis.
- CSM query trend analysis.
- Interactive Pivot Charts.
- Account-based filtering.
- Year-based filtering.
- Rolling-period filtering using slicers.

## 5. Dashboard KPIs

The project focuses on the following operational performance indicators:

| KPI | Description |
|---|---|
| Total Accounts | Distinct account count |
| Total Claims | Overall claims volume |
| Approved Claims | Total approved claims |
| PUA Claims | Total PUA claims |
| Declined Claims | Total declined claims |
| CSM Queries | Total CSM queries |

These indicators provide a consolidated overview of operational workload and claim outcomes.

## 6. Key Business Findings

### Overall Claims Performance

- Total claims recorded: 808,002.
- Approved claims: 666,762 (82.52%).
- PUA claims: 102,154 (12.64%).
- Declined claims: 27,038 (3.35%).

The remaining 12,048 claims require reconciliation against the reported outcome categories.

### Monthly Performance

- December 2024 recorded the highest monthly claim volume, with 84,860 claims.
- February 2024 recorded the lowest monthly claim volume, with 35,250 claims.
- January–April 2025 recorded approximately 31.43% more claims than January–April 2024.

### Account-wise Performance

The five highest-volume accounts were:

1. EA – 140,412 claims.
2. Airbnb – 128,930 claims.
3. Procore – 58,153 claims.
4. Pinterest – 57,177 claims.
5. Prudential – 56,731 claims.

Together, these accounts contributed approximately 54.63% of total claims.

### Data Quality Observations

The analysis also identified:

- 70 exact duplicate rows requiring review.
- Inconsistent account naming conventions.
- Differences between total claims and the sum of reported claim categories.
- A difference between raw account-name counts and normalised account counts.

These observations highlight the importance of data validation and reconciliation in operational reporting.

## 7. Project Structure

The Excel workbook contains three primary worksheets:

| Worksheet | Description |
|---|---|
| Dataset | Source operational data and calculated columns |
| Pivot Table | KPI calculations and analytical summaries |
| Dashboard | Visual representation of operational performance |

## 8. Project Deliverables

- Excel Business Operations Dashboard.
- Operational Dataset.
- Formula-based Data Transformations.
- Pivot Table Analysis.
- Operational KPI Reporting.
- Interactive Dashboard.
- Business Findings Documentation.
- Data Dictionary.

## 9. Project Outcome

This project demonstrates the practical application of Microsoft Excel for business operations analysis, MIS reporting, KPI monitoring, and management dashboard development.

It highlights the ability to work with large operational datasets, perform formula-based analysis, create structured reports, identify data quality issues, and communicate meaningful business insights through interactive visualisations.

## 10. Author

**Deepak Rana**

Project Category: Data Analytics | Advanced Excel | MIS Reporting | Business Intelligence

Primary Tool: Microsoft Excel
