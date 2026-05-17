# Financial Data Pipeline — Agentic ETL & Categorization System

> **Built to demonstrate Sales Operations skills:** multi-source data ingestion, normalization, automated categorization, and master data management — the same core competencies used in CRM ops, pipeline reporting, and revenue analytics.

---

## What This Project Does

This pipeline ingests raw financial transaction data from **7 different file formats** across **multiple bank accounts and credit cards**, normalizes everything into a single source of truth, auto-categorizes each transaction using a two-layer AI-assisted engine, and safely appends the results to a historical master ledger — with full deduplication and QA reporting.

**Built with:** Python · Pandas · OpenPyXL · Kiro AI (agentic development environment)

---

## Why This Matters for Sales Operations

| Pipeline Concept | Sales Ops Equivalent |
|---|---|
| Multi-format file ingestion | CRM exports from Salesforce, HubSpot, Outreach |
| Sign normalization & business rules | Commission plan logic, deal stage rules |
| Categorization engine (1,300+ patterns) | Lead scoring, opportunity classification |
| Deduplication before master append | CRM dedup, contact merge logic |
| QA reconciliation report | Pipeline health reporting, forecast accuracy |
| Human-in-the-loop validation | Sales ops review workflow before data goes live |

---

## Pipeline Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        RAW INPUT FILES                          │
│  Chase CC (CSV) │ Chase Bank (CSV) │ Amex Gold (XLSX)           │
│  Capital One (CSV) │ Alaska BofA (CSV) │ HSBC (CSV) │ US Bank   │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│              SCRIPT 1: process_transactions.py                  │
│                                                                 │
│  • Detects file format per source                               │
│  • Applies per-source sign normalization rules                  │
│  • Filters by configurable date cutoffs                         │
│  • Outputs uniform 9-column DataFrame                           │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                 SCRIPT 2: categorize.py                         │
│                                                                 │
│  Layer 1: Historical lookup (Master List — 1,300+ descriptions) │
│  Layer 2: Keyword rule engine (150+ regex patterns)             │
│  Layer 3: Manual overrides (context-dependent edge cases)       │
│                                                                 │
│  → Human-in-the-loop: "Needs Review" sheet for ambiguous rows   │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│              SCRIPT 3: append_to_master.py                      │
│                                                                 │
│  • Deduplicates against existing Master List                    │
│  • Maps account IDs to canonical card names                     │
│  • Normalizes date formats                                      │
│  • Appends net-new rows only                                    │
│  • Original file NEVER overwritten (safe append pattern)        │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                    MASTER LIST (XLSX)                           │
│         Historical + New transactions, fully categorized        │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Metrics (Sample Dataset)

| Metric | Value |
|---|---|
| Source files processed | 7 (across 5 financial institutions) |
| File formats handled | CSV (6 variants) + XLSX (1 variant) |
| Transactions normalized | 103 rows |
| Auto-categorization rate | 99% (102/103) |
| Duplicates detected & skipped | 6 |
| Net new rows appended | 97 |
| Categories applied | 15 categories, 28 sub-categories |

---

## Sign Normalization Logic

One of the core engineering challenges: each financial institution uses a different sign convention. This pipeline standardizes everything to a single rule:

```
+ Positive = Expense / charge / purchase
- Negative = Payment / credit / deposit / income
```

| Source | Raw Convention | Transformation |
|---|---|---|
| Chase Credit Cards | Expenses negative, payments positive | Multiply by -1 |
| Chase Bank Accounts | Withdrawals negative, deposits positive | Multiply by -1 |
| Alaska BofA (D/C type) | D rows positive, C rows negative | No flip needed |
| Capital One Venture X | Split Debit/Credit columns | Merge: Debit=+, Credit*-1 |
| Amex Gold (XLSX) | Already correct | No flip needed |
| HSBC Bank | Withdrawals negative, deposits positive | Multiply by -1 |
| US Bank | Debits negative, credits positive | Multiply by -1 |

---

## Categorization Engine

The two-layer categorization system mirrors how a Sales Ops analyst would build a lead scoring or opportunity classification model:

**Layer 1 — Historical Lookup**
Matches transaction descriptions against a Master List of previously categorized transactions. Most-frequent category wins on ties. This is equivalent to using historical CRM data to auto-classify new leads.

**Layer 2 — Keyword Rule Engine**
150+ regex patterns organized by category priority. Rules are ordered to prevent misclassification (e.g. "Capital One Mobile Payment" must be caught as a Payment before the Automotive rule can incorrectly match "Capital One").

**Layer 3 — Manual Overrides**
Context-dependent corrections applied post-engine. For example: Lyft rides during a work trip are classified as `Official > Official A`, not `Travel > Cabs`. This demonstrates understanding that the same merchant can have different business meanings depending on context.

---

## Project Structure

```
Project1-Financial-Data-Pipeline/
│
├── README.md                          ← You are here
├── requirements.txt                   ← pip install -r requirements.txt
│
├── data/
│   ├── Master_List_Sample.xlsx        ← Historical reference data (dummy)
│   ├── sample_inputs/                 ← Raw input files (all dummy data)
│   │   ├── Chase1154_Activity...CSV
│   │   ├── Chase8829_Bank_Account.CSV
│   │   ├── Alaska_BofA_0131_4044_CC.csv
│   │   ├── VentureX_Jan_Apr_2026_Credit_Card.csv
│   │   ├── Amex_Gold_Jan_Apr_2026_CC.xlsx
│   │   ├── HSBC_7393_Bank_Account.csv
│   │   └── USBank_6846_Jan_Apr_2026_Bank_Account.csv
│   └── sample_output/                 ← Generated by running the pipeline
│       ├── MASTER_Transactions_normalized.xlsx
│       ├── MASTER_Transactions_categorized.xlsx
│       └── Master_List_Updated.xlsx
│
├── scripts/
│   ├── process_transactions.py        ← Step 1: Ingest & normalize
│   ├── categorize.py                  ← Step 2: Categorize transactions
│   └── append_to_master.py            ← Step 3: Append to master ledger
│
└── docs/
    └── pipeline_diagram.png           ← Architecture diagram
```

---

## How to Run

**Prerequisites**
```bash
python --version   # Requires Python 3.10+
pip install -r requirements.txt
```

**Run the full pipeline**
```bash
# Step 1: Normalize all input files
python scripts/process_transactions.py

# Step 2: Categorize transactions
python scripts/categorize.py

# Step 3: Append to master list
python scripts/append_to_master.py
```

**Expected output**
```
[Step 1] Processing Chase Credit Cards...
[Step 2] Processing Chase Bank Accounts...
...
QA REPORT
  Chase1154_Activity...CSV     20 rows
  Chase8829_Bank_Account.CSV   16 rows
  ...
  TOTAL ROWS     : 103
  Expenses (+)   : $48,036.95
  Credits  (-)   : $-60,701.75
  Net            : $-12,664.80
```

---

## Agentic Development Approach

This project was built using **Kiro AI** as an agentic development environment — an AI-assisted workflow where the human defines requirements, reviews outputs, and makes judgment calls, while the AI handles implementation, iteration, and error correction.

This mirrors how modern Sales Operations teams are adopting AI:
- The human defines the business rules (what counts as a "work trip expense")
- The AI implements and iterates on the categorization logic
- The human validates edge cases and provides corrections
- The AI propagates those corrections consistently across all data

The result is a **human-in-the-loop pipeline** — not fully automated, not fully manual. This is the pattern that scales.

---

## Skills Demonstrated

**Data Engineering**
- Multi-format ETL pipeline (CSV, XLSX, no-header files, split-column formats)
- Business rule application per data source
- Date normalization and filtering
- Deduplication logic

**Analytics & Operations**
- Category taxonomy design (15 categories, 28 sub-categories)
- Pattern matching at scale (1,300+ historical lookups + 150+ regex rules)
- QA reconciliation reporting
- Master data management with safe append pattern

**AI & Agentic Workflows**
- Human-in-the-loop validation design
- Iterative rule refinement based on feedback
- Context-dependent override logic
- Prompt-driven development with Kiro AI

---

## About

Built as a portfolio project to demonstrate data pipeline and Sales Operations skills.
All data in this repository is **completely fabricated** — no real financial information is included.

**Author:** Karishma Jariwala
**LinkedIn:** [linkedin.com/in/karishmajariwala24-gif](https://linkedin.com/in/karishmajariwala24-gif)


