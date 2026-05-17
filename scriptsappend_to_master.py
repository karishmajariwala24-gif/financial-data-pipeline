"""
=============================================================================
SCRIPT 3: append_to_master.py
=============================================================================
Master List Append & Deduplication Engine

PURPOSE:
    Safely appends newly categorized transactions to the historical Master List.
    Performs deduplication to prevent double-counting transactions that may
    already exist in the Master List from a previous quarterly run.

DEDUPLICATION LOGIC:
    Matches on: Transaction Date + Description + Amount
    If a new row matches an existing row on all three fields, it is flagged
    as a duplicate and excluded from the append.

OUTPUT: Master_List_Updated.xlsx — the original file is NEVER overwritten.
        A new versioned file is always created for safety.

SKILLS DEMONSTRATED:
    - Safe data append with rollback protection
    - Deduplication across datasets
    - Data type normalization (date formatting)
    - Audit trail generation
    - Master data management
=============================================================================
"""

import pandas as pd
import os
from datetime import datetime

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
BASE_DIR   = os.path.join(os.path.dirname(__file__), "..", "data")
INPUT      = os.path.join(BASE_DIR, "sample_output", "MASTER_Transactions_categorized.xlsx")
ML_FILE    = os.path.join(BASE_DIR, "Master_List_Sample.xlsx")
ML_SHEET   = "Transaction Log"
DATE_ADDED = datetime.today().strftime("%#m/%#d/%Y")  # Windows format (no leading zero)

# Output — always a NEW file, never overwrites the original
OUT_FILE = os.path.join(BASE_DIR, "sample_output", "Master_List_Updated.xlsx")


# ── ACCOUNT → CARD NAME MAPPING ───────────────────────────────────────────────
# Maps our internal account identifiers to the Card names used in the Master List.
# This ensures new rows are consistent with historical naming conventions.

CARD_MAP = {
    "Chase x1154"             : "Chase Sapphire Preferred - A",
    "Chase x5813"             : "Chase World of Hyatt",
    "Chase x6082"             : "Chase Amazon Visa",
    "Chase x9338"             : "Chase Sapphire Preferred - K",
    "Chase x8829"             : "Chase Debit",
    "Chase x2020"             : "Chase Debit - Miracle",
    "Chase x0480"             : "Chase Debit",
    "Alaska BofA (0131/4044)" : "Alaska BofA - Business",
    "Capital One Venture X"   : "Capital One Venture - K",
    "Amex Gold"               : "Amex Gold",
    "Chase Sapphire x6810"    : "Chase Sapphire Preferred - A",
    "HSBC x7393"              : "HSBC Debit",
    "US Bank x6846"           : "US Bank Debit",
}


def fmt_date(val):
    """Converts any date value to mm/dd/yyyy string format."""
    if pd.isna(val):
        return ""
    try:
        dt = pd.to_datetime(val, errors="coerce")
        if pd.isna(dt):
            return str(val)
        return dt.strftime("%#m/%#d/%Y")
    except Exception:
        return str(val)


# ── STEP 1: LOAD INPUTS ───────────────────────────────────────────────────────
print("\n[Step 1] Loading categorized transactions...")
tx = pd.read_excel(INPUT, sheet_name="Transactions")
tx["Amount"] = pd.to_numeric(tx["Amount"], errors="coerce")
print(f"  Loaded {len(tx)} categorized transactions")

print("\n[Step 2] Loading existing Master List...")
ml = pd.read_excel(ML_FILE, sheet_name=ML_SHEET)
ml["Amount"] = pd.to_numeric(ml["Amount"], errors="coerce")
print(f"  Existing Master List: {len(ml)} rows")


# ── STEP 2: BUILD NEW ROWS IN MASTER LIST FORMAT ──────────────────────────────
print("\n[Step 3] Building new rows in Master List format...")
new_rows = []
unmapped_accounts = []

for _, row in tx.iterrows():
    acct = row["Account"]
    card = CARD_MAP.get(acct)
    if not card:
        unmapped_accounts.append(acct)
        card = acct  # Fallback: use account name as-is

    txdate = pd.to_datetime(row["Transaction Date"], errors="coerce")
    month  = txdate.strftime("%b") if pd.notna(txdate) else ""
    year   = int(txdate.year)      if pd.notna(txdate) else ""

    new_rows.append({
        "Date Added"       : DATE_ADDED,
        "Card"             : card,
        "Category"         : row["Category"],
        "Sub Category"     : row["Sub-Category"],
        "Month"            : month,
        "Transaction Date" : fmt_date(row["Transaction Date"]),
        "Amount"           : row["Amount"],
        "Description"      : row["Description"],
        "Year"             : year,
    })

if unmapped_accounts:
    print(f"  WARNING: Unmapped accounts (used as-is): {set(unmapped_accounts)}")

new_df = pd.DataFrame(new_rows)
print(f"  Built {len(new_df)} new rows")


# ── STEP 3: DEDUPLICATION ─────────────────────────────────────────────────────
print("\n[Step 4] Checking for duplicates against existing Master List...")

# Normalize dates for comparison
ml["_dt"]     = ml["Transaction Date"].apply(fmt_date)
new_df["_dt"] = new_df["Transaction Date"].apply(fmt_date)

# Build match key: date + description (stripped) + amount (rounded)
ml["_key"]     = ml["_dt"]     + "|" + ml["Description"].astype(str).str.strip() + "|" + ml["Amount"].round(2).astype(str)
new_df["_key"] = new_df["_dt"] + "|" + new_df["Description"].astype(str).str.strip() + "|" + new_df["Amount"].round(2).astype(str)

existing_keys = set(ml["_key"])
new_df["_is_duplicate"] = new_df["_key"].isin(existing_keys)

duplicates = new_df[new_df["_is_duplicate"]]
clean_rows = new_df[~new_df["_is_duplicate"]]

print(f"  Duplicates found (excluded): {len(duplicates)}")
print(f"  Net new rows to append     : {len(clean_rows)}")

if len(duplicates) > 0:
    print("\n  Duplicate rows (already in Master List):")
    pd.set_option("display.max_colwidth", 45)
    print(duplicates[["Transaction Date", "Card", "Description", "Amount"]].to_string(index=False))


# ── STEP 4: APPEND ────────────────────────────────────────────────────────────
print("\n[Step 5] Appending to Master List...")

# Clean up helper columns before appending
ml_clean     = ml.drop(columns=["_dt", "_key"])
clean_rows_f = clean_rows.drop(columns=["_dt", "_key", "_is_duplicate"])

# Normalize existing Master List date columns to same string format
ml_clean["Date Added"]       = ml_clean["Date Added"].apply(fmt_date)
ml_clean["Transaction Date"] = ml_clean["Transaction Date"].apply(fmt_date)

combined = pd.concat([ml_clean, clean_rows_f], ignore_index=True)
print(f"  Original rows : {len(ml_clean)}")
print(f"  New rows added: {len(clean_rows_f)}")
print(f"  Final total   : {len(combined)}")


# ── STEP 5: WRITE OUTPUT ──────────────────────────────────────────────────────
print("\n[Step 6] Writing updated Master List...")
with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
    combined.to_excel(writer, sheet_name=ML_SHEET, index=False)
    ws = writer.sheets[ML_SHEET]
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

print(f"\nOutput written to:\n  {OUT_FILE}")
print(f"\nDate Added stamp : {DATE_ADDED}")
print(f"Original rows    : {len(ml_clean)}")
print(f"New rows added   : {len(clean_rows_f)}")
print(f"Duplicates skipped: {len(duplicates)}")
print(f"Final row count  : {len(combined)}")
print("\nThe original Master_List_Sample.xlsx has NOT been modified.")
print("Review Master_List_Updated.xlsx before replacing the original.")
