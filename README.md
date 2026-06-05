# Payment Fraud Detection Dashboard

A data analysis pipeline for detecting fraudulent and anomalous payment transactions, built with Python, SQL, and Excel (with VBA automation).

---

## Overview

This project simulates a payment review workflow similar to those used in government and insurance organizations. It ingests transaction records, runs SQL-based fraud detection queries, and outputs a formatted multi-sheet Excel report with flagged transactions, regional breakdowns, and trend analysis.

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

## Sample Output

- **300 transactions** reviewed across 2024
- **37 flagged** (12.33% flag rate)
- **$173,472** in flagged payment value identified
- Breakdown by region, payment type, and processor


---

## VBA Macros Included

Open the `.vba` file in the Excel VBA editor (Alt + F11) to use:

1. `FormatAllSheets` — Auto-fits columns and standardizes fonts
2. `HighlightOverpayments` — Colors overpayment rows red
3. `AddFlagFilter` — Adds dropdown filters to flagged transactions
4. `ExportFlaggedToNewWorkbook` — Saves flagged records to desktop
5. `GenerateSummaryStats` — Creates a quick stats sheet

---

