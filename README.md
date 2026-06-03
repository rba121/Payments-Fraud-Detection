# Payment Fraud Detection Dashboard

A data analysis pipeline for detecting fraudulent and anomalous payment transactions, built with Python, SQL, and Excel (with VBA automation).

---

## Overview

This project simulates a payment review workflow similar to those used in government and insurance organizations. It ingests transaction records, runs SQL-based fraud detection queries, and outputs a formatted multi-sheet Excel report with flagged transactions, regional breakdowns, and trend analysis.

**Designed to mirror real-world payment compliance workflows.**

---

## Tech Stack

| Tool | Usage |
|------|-------|
| Python | Data generation, pipeline orchestration, Excel report building |
| SQL (SQLite) | Fraud detection queries, aggregations, joins |
| Excel (openpyxl) | Multi-sheet formatted report with charts |
| VBA | Auto-formatting macros, filter setup, export automation |
| pandas | Data manipulation and query result handling |

---

## Features

- **Synthetic dataset** — 300 realistic payment transactions with injected fraud patterns
- **10 SQL fraud detection queries** covering:
  - Overpayment detection (threshold-based)
  - Duplicate payment detection (JOIN-based)
  - High-risk claimant identification
  - Regional fraud rate analysis
  - Monthly trend tracking
  - Payment type fraud rate breakdown
  - Processor-level flag review
- **7-sheet Excel report** auto-generated with:
  - Executive Summary with KPI metrics
  - Overpayments sheet (flagged red)
  - High Risk Claimants sheet
  - Regional Analysis with bar chart
  - Monthly Trend sheet
  - Payment Type Analysis
  - Full Flagged Transactions log
- **VBA macros** for:
  - Auto-formatting all sheets
  - Highlighting overpayment rows
  - Adding AutoFilters for review workflows
  - Exporting flagged records to a new workbook
  - Generating quick stats summary

---

## Project Structure

```
payment_fraud_detector/
├── data/
│   ├── payments.db              # SQLite database
│   └── raw_transactions.csv     # Raw transaction export
├── sql/
│   └── fraud_detection_queries.sql   # All 10 SQL queries
├── scripts/
│   ├── generate_data.py         # Generates synthetic dataset + DB
│   ├── analyze_and_export.py    # Runs queries, builds Excel report
│   └── fraud_dashboard_macros.vba    # VBA macros for Excel
└── reports/
    └── Payment_Fraud_Report.xlsx     # Final output report
```

---

## How to Run

```bash
# 1. Install dependencies
pip install pandas openpyxl

# 2. Generate the dataset and database
python scripts/generate_data.py

# 3. Run analysis and export Excel report
python scripts/analyze_and_export.py
```

The report will be saved to `reports/Payment_Fraud_Report.xlsx`.

---

## Fraud Detection Logic

| Flag Type | Detection Method |
|-----------|-----------------|
| Overpayment | `WHERE amount > 4500` threshold query |
| Duplicate Payment | Self-JOIN on `claimant_id + payment_date + amount` |
| Inactive Claimant | Status flag in transaction record |
| Multiple Same-Day Payments | GROUP BY with HAVING COUNT > 1 |

---

## Sample Output

- **300 transactions** reviewed across 2024
- **37 flagged** (12.33% flag rate)
- **$173,472** in flagged payment value identified
- Breakdown by region, payment type, and processor

---

## Key SQL Concepts Used

- `JOIN` (self-join for duplicate detection)
- `GROUP BY` + `HAVING` for aggregation filtering
- `CASE WHEN` for conditional aggregation
- `SUBSTR()` for date truncation to month
- Subqueries and multi-table aggregation

---

## VBA Macros Included

Open the `.vba` file in the Excel VBA editor (Alt + F11) to use:

1. `FormatAllSheets` — Auto-fits columns and standardizes fonts
2. `HighlightOverpayments` — Colors overpayment rows red
3. `AddFlagFilter` — Adds dropdown filters to flagged transactions
4. `ExportFlaggedToNewWorkbook` — Saves flagged records to desktop
5. `GenerateSummaryStats` — Creates a quick stats sheet

---

## Skills Demonstrated

- SQL query design for fraud detection use cases
- Data pipeline: raw data → database → analysis → Excel report
- Excel report automation with Python (openpyxl)
- VBA macro development for workflow automation
- Data analysis and anomaly detection patterns
