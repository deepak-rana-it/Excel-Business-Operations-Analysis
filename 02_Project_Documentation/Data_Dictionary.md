# Data Dictionary – Business Operations Analysis

## 1. Introduction

This document provides the definitions, descriptions, and business purposes of all 20 columns available in the Dataset worksheet of the Excel Business Operations Analysis project.

The dataset contains operational claims information across multiple accounts and reporting periods.

It includes source fields, calculated columns, operational metrics, and validation indicators used for analysis and dashboard reporting.

## 2. Dataset Information

| Attribute | Details |
|---|---|
| Worksheet Name | Dataset |
| Total Records | 18,325 |
| Total Columns | 20 |
| Reporting Period | January 2024 – April 2025 |
| Data Format | Excel Structured Table |
| Primary Analysis | Claims and Operational Performance |

## 3. Column Definitions

### A. Source Data Columns

These columns contain the original operational information.

| No. | Column Name | Data Type | Description | Business Purpose |
|---|---|---|---|---|
| 1 | Year | Text | Reporting year extracted from the transaction date. | Year-based reporting and filtering. |
| 2 | Month | Text | Reporting month extracted from the transaction date. | Monthly trend analysis. |
| 3 | Date | Date | Reporting date associated with the operational record. | Time-based analysis. |
| 4 | Day | Text | Day of the week associated with the reporting date. | Day-level analysis. |
| 5 | Total Claims | Number | Total claims recorded for the account and reporting period. | Measures overall claims volume. |
| 6 | Approved | Number | Number of approved claims. | Tracks successful claim outcomes. |
| 7 | PUA | Number | Number of PUA claims. | Monitors PUA claim volume. |
| 8 | Declined | Number | Number of declined claims. | Tracks unsuccessful claim outcomes. |
| 9 | CSM Queries | Number | Number of CSM-related queries. | Monitors customer support query volume. |
| 10 | Account Name | Text | Name of the business account associated with the record. | Account-wise performance analysis. |

### B. Calculated and Validation Columns

These columns are generated using Excel formulas to support data analysis and reporting.

| No. | Column Name | Data Type | Description | Excel Function / Logic |
|---|---|---|---|---|
| 11 | Column1 | Boolean | Checks whether the Date field contains a valid numeric Excel date value. | ISNUMBER |
| 12 | Rolling Column | Text | Identifies records falling between April 1, 2024 and April 30, 2025. | IF + DATE |
| 13 | PUA Total | Number | Aggregated PUA claims for the same account, year, and month. | SUMIFS + VALUE |
| 14 | Claim Total | Number | Aggregated total claims for the same account, year, and month. | SUMIFS + VALUE |
| 15 | PUA % | Percentage | Percentage of PUA claims against total claims. | IFERROR |
| 16 | Total CSM | Number | Aggregated CSM queries for the same account, year, and month. | SUMIFS + VALUE |
| 17 | CSM % | Percentage | CSM queries as a percentage of total claims. | IFERROR |
| 18 | checkPUA | Boolean | Checks whether the CSM % field contains a numeric value. | ISNUMBER |
| 19 | Decline Total | Number | Aggregated declined claims for the same account, year, and month. | SUMIFS + VALUE |
| 20 | Decline% | Percentage | Declined claims as a percentage of total claims. | IFERROR |

## 4. Important Excel Formulas

### 4.1 Year Extraction

**Column:** Year

```excel
=TEXT([@Date],"yyyy")
```

Purpose: Extracts the year from the reporting date.

### 4.2 Month Extraction

**Column:** Month

```excel
=TEXT([@Date],"mmmm")
```

Purpose: Extracts the full month name from the reporting date.

### 4.3 Date Validation

**Column:** Column1

```excel
=ISNUMBER(C2)
```

Purpose: Checks whether the date value is stored as a numeric Excel date.

### 4.4 Rolling Period Identification

**Column:** Rolling Column

```excel
=IF([@Date]>=DATE(2024,4,1),
   IF([@Date]<=DATE(2025,4,30),
      "Rolling 13 Months Data"," "))
```

Purpose: Identifies records within the selected 13-month reporting window.

### 4.5 PUA Aggregation

**Column:** PUA Total

```excel
=VALUE(SUMIFS(G:G,B:B,[@Month],
A:A,[@Year],
J:J,[@[Account Name]]))
```

Purpose: Calculates the total PUA claims for the corresponding account, month, and year.

### 4.6 Total Claims Aggregation

**Column:** Claim Total

```excel
=VALUE(SUMIFS(E:E,B:B,[@Month],
A:A,[@Year],
J:J,[@[Account Name]]))
```

Purpose: Calculates total claims for the corresponding account, month, and year.

### 4.7 PUA Percentage

**Column:** PUA %

```excel
=IFERROR([@[PUA Total]]/[@[Claim Total]],0)
```

Purpose: Calculates PUA claims as a percentage of total claims.

### 4.8 CSM Query Aggregation

**Column:** Total CSM

```excel
=VALUE(SUMIFS(I:I,B:B,[@Month],
A:A,[@Year],
J:J,[@[Account Name]]))
```

Purpose: Calculates total CSM queries for the corresponding account, month, and year.

### 4.9 CSM Percentage

**Column:** CSM %

```excel
=IFERROR([@[Total CSM]]/[@[Claim Total]],0)
```

Purpose: Calculates the proportion of CSM queries relative to total claims.

### 4.10 Declined Claims Aggregation

**Column:** Decline Total

```excel
=VALUE(SUMIFS(H:H,B:B,[@Month],
A:A,[@Year],
J:J,[@[Account Name]]))
```

Purpose: Calculates total declined claims for the corresponding account, month, and year.

### 4.11 Declined Claims Percentage

**Column:** Decline%

```excel
=IFERROR([@[Decline Total]]/[@[Claim Total]],0)
```

Purpose: Calculates declined claims as a percentage of total claims.

## 5. Key Business Metrics

The following metrics are derived from the dataset and used in operational reporting.

| Metric | Calculation |
|---|---|
| Total Claims | SUM of Total Claims |
| Approved Claims | SUM of Approved |
| PUA Claims | SUM of PUA |
| Declined Claims | SUM of Declined |
| Total CSM Queries | SUM of CSM Queries |
| Approval Percentage | Approved Claims / Total Claims |
| PUA Percentage | PUA Claims / Total Claims |
| Declined Percentage | Declined Claims / Total Claims |
| CSM Query Percentage | CSM Queries / Total Claims |

## 6. Data Quality Considerations

The following data quality checks are relevant to this dataset:

1. Validate reporting dates and date formats.
2. Standardise account naming conventions.
3. Identify duplicate records.
4. Check numeric fields for invalid values.
5. Reconcile total claims against approved, PUA, and declined categories.
6. Validate calculated percentages.
7. Ensure Pivot Tables reflect the latest source data.

### Important Observations

- The source dataset contains 94 distinct account-name values.
- After trimming spaces and normalising letter case, there are 83 distinct account names.
- There are 70 exact duplicate rows requiring review.
- Some records have differences between Total Claims and the sum of Approved, PUA, and Declined.

These observations should be investigated before making changes to the original dataset.

## 7. Reporting and Analysis Usage

The dataset supports the following reporting dimensions:

| Analysis Dimension | Relevant Columns |
|---|---|
| Yearly Analysis | Year |
| Monthly Analysis | Year, Month |
| Account Performance | Account Name |
| Claims Performance | Total Claims |
| Approval Analysis | Approved |
| PUA Analysis | PUA, PUA Total, PUA % |
| Declined Claims Analysis | Declined, Decline Total, Decline% |
| CSM Query Analysis | CSM Queries, Total CSM, CSM % |
| Rolling Period Analysis | Date, Rolling Column |

## 8. Conclusion

The Data Dictionary provides a structured reference for understanding the dataset, its calculated fields, and the business purpose of each column.

It supports transparency, maintainability, and consistency in Excel-based operational reporting and business analysis.
