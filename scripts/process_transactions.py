"""
=============================================================================
SCRIPT 1: process_transactions.py
=============================================================================
Multi-Source Financial Transaction Normalizer

PURPOSE:
    Ingests raw financial data from 7 different file formats across multiple
    bank accounts and credit cards. Normalizes all transactions into a single
    unified DataFrame with consistent column names, date formats, and sign
    conventions.

SIGN CONVENTION (standardized output):
    +  Positive amount = Expense / charge / purchase
    -  Negative amount = Payment / credit / deposit / income

INPUT:  Raw CSV/XLSX files from sample_inputs/
OUTPUT: MASTER_Transactions_normalized.xlsx (Transactions sheet)

SKILLS DEMONSTRATED:
    - Multi-format ETL pipeline design
    - Data normalization across heterogeneous sources
    - Per-source business rule application
    - Date filtering and deduplication logic
    - QA reconciliation reporting
=============================================================================
"""

import pandas as pd
import os
from io import StringIO
from datetime import datetime

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
BASE      = os.path.join(os.path.dirname(__file__), "..", "data", "sample_inputs")
ML_FILE   = os.path.join(os.path.dirname(__file__), "..", "data", "Master_List_Sample.xlsx")
OUT_DIR   = os.path.join(os.path.dirname(__file__), "..", "data", "sample_output")
OUT_FILE  = os.path.join(OUT_DIR, "MASTER_Transactions_normalized.xlsx")

os.makedirs(OUT_DIR, exist_ok=True)

# Global date cutoff — only include transactions from this date onwards
# Prevents re-importing historical data already in the Master List
DATE_MIN = pd.Timestamp("2025-11-01")

# Per-account stricter cutoffs where files contain excessive history
DATE_OVERRIDES = {
    "Chase x8829" : pd.Timestamp("2025-12-27"),
    "Chase x0480" : pd.Timestamp("2026-01-02"),
    "HSBC x7393"  : pd.Timestamp("2026-01-02"),
}

# Output columns — uniform across all sources
COLS = ["Transaction Date", "Description", "Amount", "Type",
        "Financial Institution", "Account", "Source File"]

all_dfs = []


# ── HELPER FUNCTIONS ──────────────────────────────────────────────────────────

def clean_amount(value):
    """
    Safely converts any amount string or number to float.
    Handles commas, dollar signs, and empty strings.
    """
    if pd.isna(value):
        return 0.0
    return float(str(value).replace(",", "").replace("$", "").strip() or 0)


def add_to_master(df, financial_institution, account_name, source_filename):
    """
    Applies date filtering and appends a normalized DataFrame to the master list.

    Args:
        df: DataFrame with columns [Transaction Date, Description, Amount, Type]
        financial_institution: Bank/institution name (e.g. 'Chase', 'Amex')
        account_name: Account identifier (e.g. 'Chase x1154')
        source_filename: Original filename for traceability
    """
    df = df.copy()
    df["Transaction Date"] = pd.to_datetime(df["Transaction Date"], errors="coerce")

    # Apply per-account date cutoff (or global minimum)
    cutoff = DATE_OVERRIDES.get(account_name, DATE_MIN)
    before = len(df)
    df = df[df["Transaction Date"] >= cutoff]
    dropped = before - len(df)
    if dropped > 0:
        print(f"  [{account_name}] Filtered {dropped} rows before {cutoff.date()}")

    df["Financial Institution"] = financial_institution
    df["Account"]               = account_name
    if "Source File" not in df.columns:
        df["Source File"] = source_filename

    all_dfs.append(df[COLS])


# ── STEP 1: Chase Credit Cards ────────────────────────────────────────────────
# Format: Transaction Date, Post Date, Description, Category, Type, Amount, Memo
# Sign rule: Expenses are NEGATIVE in source → flip to positive
# Payments are POSITIVE in source → flip to negative

print("\n[Step 1] Processing Chase Credit Cards...")
chase_cc_files = {
    "Chase1154_Activity20260101_20260430_Credit_Card.CSV": "Chase x1154",
}
for filename, account in chase_cc_files.items():
    filepath = os.path.join(BASE, filename)
    if not os.path.exists(filepath):
        print(f"  WARNING: {filename} not found, skipping.")
        continue
    df = pd.read_csv(filepath)
    df["Amount"] = df["Amount"].apply(clean_amount) * -1  # Flip sign
    df["Type"]   = df["Type"].fillna("Sale")
    add_to_master(
        df[["Transaction Date", "Description", "Amount", "Type"]],
        "Chase", account, filename
    )
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── STEP 2: Chase Bank Accounts ───────────────────────────────────────────────
# Format: Details, Posting Date, Description, Amount, Type, Balance, Check#
# Sign rule: Withdrawals are NEGATIVE in source → flip to positive
#            Deposits are POSITIVE in source → flip to negative

print("\n[Step 2] Processing Chase Bank Accounts...")
chase_bank_files = {
    "Chase8829_Bank_Account.CSV": "Chase x8829",
}
for filename, account in chase_bank_files.items():
    filepath = os.path.join(BASE, filename)
    if not os.path.exists(filepath):
        print(f"  WARNING: {filename} not found, skipping.")
        continue
    df = pd.read_csv(filepath)
    df["Amount"] = df["Amount"].apply(clean_amount) * -1  # Flip sign
    df = df.rename(columns={"Posting Date": "Transaction Date"})
    df["Type"] = df["Type"].fillna("Debit")
    add_to_master(
        df[["Transaction Date", "Description", "Amount", "Type"]],
        "Chase", account, filename
    )
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── STEP 3: Alaska BofA Credit Cards (Combined Account) ───────────────────────
# Format: 4 summary header rows, then data with CardHolder, Card#, Dates, Amount
# Sign rule: Transaction Type D (Debit) = already positive (expense)
#            Transaction Type C (Credit) = already negative (payment)
#            NO flip needed — source signs are already correct
# Note: Cards 0131 and 4044 are nested cards on the same account

print("\n[Step 3] Processing Alaska BofA Credit Cards (combined 0131/4044)...")
alaska_frames = []
for filename in ["Alaska_BofA_0131_4044_CC.csv"]:
    filepath = os.path.join(BASE, filename)
    if not os.path.exists(filepath):
        print(f"  WARNING: {filename} not found, skipping.")
        continue
    raw = open(filepath, encoding="utf-8", errors="replace").readlines()
    df  = pd.read_csv(StringIO("".join(raw[4:])))  # Skip 4 summary header rows
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Description"].notna() & (df["Description"].str.strip() != "")]
    df["Amount"] = df["Amount"].apply(clean_amount)  # No flip — source correct
    df["Type"]   = df["Transaction Type"].map({"D": "Sale", "C": "Payment"}).fillna("Sale")
    df = df.rename(columns={"Trans. Date": "Transaction Date"})
    df["Source File"] = filename
    alaska_frames.append(df[["Transaction Date", "Description", "Amount", "Type", "Source File"]])

if alaska_frames:
    alaska_all = pd.concat(alaska_frames, ignore_index=True)
    add_to_master(alaska_all, "Bank of America", "Alaska BofA (0131/4044)", "Alaska_BofA_0131_4044_CC.csv")
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── STEP 4: Capital One Venture X ─────────────────────────────────────────────
# Format: Transaction Date, Posted Date, Card No., Description, Category, Debit, Credit
# Sign rule: Debit column = expense (positive)
#            Credit column = payment/return (negative)
# Note: Two separate columns must be merged into one Amount column

print("\n[Step 4] Processing Capital One Venture X...")
for filename in ["VentureX_Jan_Apr_2026_Credit_Card.csv"]:
    filepath = os.path.join(BASE, filename)
    if not os.path.exists(filepath):
        print(f"  WARNING: {filename} not found, skipping.")
        continue
    df = pd.read_csv(filepath)
    df["Debit"]  = df["Debit"].apply(clean_amount)
    df["Credit"] = df["Credit"].apply(clean_amount)
    # Merge: Debit = positive expense, Credit = negative payment
    df["Amount"] = df.apply(
        lambda r: r["Debit"] if r["Debit"] != 0 else r["Credit"] * -1, axis=1
    )
    df["Type"] = df.apply(
        lambda r: "Sale" if r["Debit"] != 0 else "Payment", axis=1
    )
    add_to_master(
        df[["Transaction Date", "Description", "Amount", "Type"]],
        "Capital One", "Capital One Venture X", filename
    )
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── STEP 5: Amex Gold XLSX ────────────────────────────────────────────────────
# Format: XLSX with 6 metadata rows, row 7 is the actual header, data from row 8
# Sign rule: Expenses already positive, payments already negative — NO flip needed
# Note: Must use header=6 (0-indexed) to skip metadata rows

print("\n[Step 5] Processing Amex Gold XLSX...")
for filename in ["Amex_Gold_Jan_Apr_2026_CC.xlsx"]:
    filepath = os.path.join(BASE, filename)
    if not os.path.exists(filepath):
        print(f"  WARNING: {filename} not found, skipping.")
        continue
    df = pd.read_excel(filepath, sheet_name="Transaction Details", header=6)
    df = df[["Date", "Description", "Amount"]].copy()
    df = df[df["Date"].notna() & (df["Date"].astype(str).str.strip() != "")]
    df["Amount"] = df["Amount"].apply(clean_amount)
    df["Type"]   = df["Amount"].apply(lambda x: "Payment" if x < 0 else "Sale")
    df = df.rename(columns={"Date": "Transaction Date"})
    add_to_master(
        df[["Transaction Date", "Description", "Amount", "Type"]],
        "American Express", "Amex Gold", filename
    )
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── STEP 6: HSBC Bank Account ─────────────────────────────────────────────────
# Format: NO header row — columns are Date, Description, Amount, Balance
# Sign rule: Withdrawals NEGATIVE → flip to positive
#            Deposits POSITIVE → flip to negative

print("\n[Step 6] Processing HSBC Bank Account...")
hsbc_path = os.path.join(BASE, "HSBC_7393_Bank_Account.csv")
if os.path.exists(hsbc_path):
    df = pd.read_csv(hsbc_path, header=None,
                     names=["Transaction Date", "Description", "Amount", "Balance"])
    df["Amount"] = df["Amount"].apply(clean_amount) * -1  # Flip sign
    df["Type"]   = df["Amount"].apply(lambda x: "Payment" if x < 0 else "Sale")
    add_to_master(
        df[["Transaction Date", "Description", "Amount", "Type"]],
        "HSBC", "HSBC x7393", "HSBC_7393_Bank_Account.csv"
    )
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── STEP 7: US Bank Account ───────────────────────────────────────────────────
# Format: Date, Transaction, Name, Memo, Amount
# Sign rule: Debits NEGATIVE → flip to positive
#            Credits POSITIVE → flip to negative

print("\n[Step 7] Processing US Bank Account...")
usbank_path = os.path.join(BASE, "USBank_6846_Jan_Apr_2026_Bank_Account.csv")
if os.path.exists(usbank_path):
    df = pd.read_csv(usbank_path)
    df.columns = [c.strip() for c in df.columns]
    df["Amount"] = df["Amount"].apply(clean_amount) * -1  # Flip sign
    df["Type"]   = df["Amount"].apply(lambda x: "Payment" if x < 0 else "Sale")
    df = df.rename(columns={"Date": "Transaction Date", "Name": "Description"})
    add_to_master(
        df[["Transaction Date", "Description", "Amount", "Type"]],
        "US Bank", "US Bank x6846", "USBank_6846_Jan_Apr_2026_Bank_Account.csv"
    )
print(f"  Total rows so far: {sum(len(d) for d in all_dfs)}")


# ── CONSOLIDATE ───────────────────────────────────────────────────────────────
print("\n[Consolidating] Merging all sources into master DataFrame...")
master = pd.concat(all_dfs, ignore_index=True)
master["Transaction Date"] = pd.to_datetime(master["Transaction Date"], errors="coerce")
master = master.sort_values("Transaction Date")
master["Transaction Date"] = master["Transaction Date"].dt.strftime("%m/%d/%Y")
master["Category"]     = ""  # To be filled by categorize.py
master["Sub-Category"] = ""

final_cols = ["Financial Institution", "Account", "Source File",
              "Transaction Date", "Description", "Type", "Amount",
              "Category", "Sub-Category"]
master = master[final_cols]


# ── QA REPORT ─────────────────────────────────────────────────────────────────
print("\n" + "=" * 65)
print("QA REPORT — process_transactions.py")
print("=" * 65)
counts = master.groupby("Source File").size().reset_index(name="Rows")
for _, row in counts.iterrows():
    print(f"  {row['Source File']:<55} {row['Rows']:>4} rows")

total = len(master)
pos   = master[master["Amount"] > 0]["Amount"].sum()
neg   = master[master["Amount"] < 0]["Amount"].sum()
print(f"\n  TOTAL ROWS     : {total}")
print(f"  Expenses (+)   : ${pos:,.2f}")
print(f"  Credits  (-)   : ${neg:,.2f}")
print(f"  Net            : ${pos + neg:,.2f}")
print("=" * 65)


# ── WRITE OUTPUT ──────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
    master.to_excel(writer, sheet_name="Transactions", index=False)

print(f"\nOutput written to:\n  {OUT_FILE}")
print(f"\nNext step: Run categorize.py to assign categories to each transaction.")
