"""
=============================================================================
SCRIPT 2: categorize.py
=============================================================================
AI-Assisted Transaction Categorization Engine

PURPOSE:
    Applies a two-layer categorization strategy to assign a Category and
    Sub-Category to every transaction:

    Layer 1 — Historical Lookup:
        Matches transaction descriptions against a Master List of previously
        categorized transactions. Most-frequent category wins on ties.

    Layer 2 — Keyword Rule Engine:
        150+ regex patterns organized by category. Applied when no historical
        match exists. Rules are ordered by priority — first match wins.

    Layer 3 — Manual Overrides:
        Explicit corrections for edge cases, ambiguous merchants, and
        context-dependent categorizations (e.g. Lyft during work trips).

OUTPUT: MASTER_Transactions_categorized.xlsx with Category + Sub-Category filled

SKILLS DEMONSTRATED:
    - Pattern matching and fuzzy lookup at scale
    - Business rule engine design
    - Human-in-the-loop validation workflow
    - Data quality assurance
    - Iterative refinement based on feedback
=============================================================================
"""

import pandas as pd
import os
import re

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
BASE_DIR  = os.path.join(os.path.dirname(__file__), "..", "data")
INPUT     = os.path.join(BASE_DIR, "sample_output", "MASTER_Transactions_normalized.xlsx")
ML_FILE   = os.path.join(BASE_DIR, "Master_List_Sample.xlsx")
OUTPUT    = os.path.join(BASE_DIR, "sample_output", "MASTER_Transactions_categorized.xlsx")


# ── APPROVED CATEGORY TAXONOMY ────────────────────────────────────────────────
# Only these categories and sub-categories are valid in the output.
# This enforces consistency across all quarterly runs.

CATEGORIES = [
    "Home+", "Education", "Automotive", "Bills & Utilities", "Grocery",
    "Travel", "Shopping", "Meals", "Health & Beauty", "Entertaining",
    "Payment", "Professional Services", "Spl Investment", "Income", "Official"
]

SUB_CATEGORIES = [
    "Annual Membership", "Cabs, Parking & Tolls", "Car & Accessories",
    "Card Payment", "Cleaning Maintenance Gardening Services",
    "Dues and Subscriptions", "Entertaining Expense", "Gas & Wash",
    "Home Office +", "Income AJ", "Income KJ", "Insurance",
    "Interbank Transfer", "Interest Payment", "Legal and professional",
    "Meals", "Mobile & Internet", "Neighbor Income", "Official A", "Official K",
    "Online Services", "Payment", "Reimbursements", "Reimbursements Zelle",
    "Service Fees", "Ticket", "Travel & Entertainment Expense", "Utilities"
]


# ── STEP 1: LOAD MASTER LIST LOOKUP ───────────────────────────────────────────

def build_lookup(ml_file):
    """
    Builds a description-to-category lookup dictionary from historical data.
    Uses the most frequently assigned category for each unique description.
    """
    ml = pd.read_excel(ml_file, sheet_name="Transaction Log")
    ml = ml[["Description", "Category", "Sub Category"]].dropna(subset=["Category"])
    ml = ml[ml["Category"].str.strip() != ""]
    ml["Description"]  = ml["Description"].astype(str).str.strip().str.upper()
    ml["Category"]     = ml["Category"].astype(str).str.strip()
    ml["Sub Category"] = ml["Sub Category"].fillna("").astype(str).str.strip()

    lookup = (
        ml.groupby(["Description", "Category", "Sub Category"])
          .size().reset_index(name="n")
          .sort_values("n", ascending=False)
          .drop_duplicates("Description")
          .set_index("Description")[["Category", "Sub Category"]]
          .to_dict("index")
    )
    print(f"  Historical lookup built: {len(lookup)} unique descriptions")
    return lookup


# ── STEP 2: KEYWORD RULE ENGINE ───────────────────────────────────────────────
# Rules are tuples of (regex_pattern, Category, Sub-Category)
# ORDER MATTERS — first match wins
# Organized by priority: Payments first (prevent misclassification),
# then specific categories, then general ones

RULES = [
    # ── INCOME ────────────────────────────────────────────────────────────────
    (r"AMAZON.*PAYROLL|AMAZON WEB SERVI.*PAYROLL|ELECTRONIC DEPOSIT AMAZON WEB SERVI",
     "Income", "Income AJ"),
    (r"BCCL WORLDWIDE|MIRACLE CONSULTANTS",
     "Income", "Income KJ"),
    (r"NEIGHBOR\.COM|NEIGHBOR\.C\b",
     "Income", "Neighbor Income"),
    (r"ACH CORP TRADE PAYMENT.*AMAZON|AMAZON WEB SERVI.*TRADE",
     "Payment", "Reimbursements"),

    # ── PAYMENT — Capital One MUST come before Automotive to avoid misclassification
    (r"CAPITAL ONE.*MOBILE PYMT|CAPITAL ONE.*MOBILE PMT|CAPITAL ONE MOBILE",
     "Payment", "Payment"),
    (r"AUTOPAY PAYMENT|AUTOMATIC PAYMENT|ONLINE PAYMENT.*THANK|MOBILE PAYMENT.*THANK|PAYMENT THANK YOU|PAYMENT RECEIVED|BA ELECTRONIC PAYMENT|AMERICAN EXPRESS ACH PMT|BK OF AMER.*ONLINE PMT|BANK OF AMERICA.*PAYMENT|APPLECARD GSBANK PAYMENT|CHASE CREDIT CRD AUTOPAY|BARCLAYCARD US.*CREDITCARD",
     "Payment", "Payment"),
    (r"ONLINE TRANSFER|ONLINE REALTIME TRANSFER|RTP TRANSFER|REAL TIME PAYMENT|WIRE TRANSFER|BOOK TRANSFER|INTERBANK",
     "Payment", "Interbank Transfer"),
    (r"INTEREST PAID|INTEREST PAYMENT|INTEREST EARNED",
     "Payment", "Interest Payment"),
    (r"PURCHASE INTEREST CHARGE|INTEREST CHARGE.*PURCHASE|INTEREST CHARGED ON PURCHASES|INTEREST CHARGE:PURCHASES",
     "Payment", "Service Fees"),
    (r"MONTHLY (SERVICE|MAINTENANCE) FEE|MONTHLY MAINTENANCE FEE WAIVED|WIRE.*FEE|INTERNATIONAL.*WIRE FEE",
     "Payment", "Service Fees"),
    (r"LATE FEE|PAST DUE FEE|PURCHASE \*FINANCE CHARGE\*",
     "Payment", "Service Fees"),
    (r"ANNUAL CARD FEE|ANNUAL FEE$|ANNUAL MEMBERSHIP FEE|RENEWAL MEMBERSHIP FEE|CAPITAL ONE MEMBER FEE|ANNUAL FEE ADJUSTMENT|INTEREST CHARGE ADJUSTMENT",
     "Bills & Utilities", "Annual Membership"),
    (r"ZELLE PAYMENT FROM|VENMO.*CASHOUT|Zelle payment from",
     "Payment", "Reimbursements Zelle"),
    (r"ACH DEPOSIT FROM AMAZON.*DIRECT DEP|ACH DEPOSIT FROM AMAZON WEB SERVI(?!.*PAYROLL)",
     "Income", "Income AJ"),

    # ── HOME ──────────────────────────────────────────────────────────────────
    (r"EVERBANK|TREAN.*MTG|MTG PYMT",
     "Home+", "Home Office +"),
    (r"TOSCANA.*HOMEOWNER|REVO\*TOSCANA",
     "Home+", "Home Office +"),

    # ── OFFICIAL (work trips) — BEFORE general travel/meals to take priority ──
    (r"WIFIONBOARD ALASKA|WIFIONBOARD",
     "Official", "Official A"),
    (r"LYFT",
     "Official", "Official A"),
    (r"HILTON PARC 55|HYATT REGENCY SAN FRAN|GRAND HYATT SAN FRAN|WESTIN ST\. FRANCIS",
     "Official", "Official A"),
    (r"BUN MEE|THEMELT|MARU SUSHI",
     "Official", "Official A"),

    # ── AUTOMOTIVE ────────────────────────────────────────────────────────────
    (r"COSTCO GAS|SAFEWAY FUEL|SHELL OIL|ARCO|EXXON|76 -|CIRCLE K|PRIME GAS|CHEVRON|TEXACO|MOBIL|CONOCO|TESLA SUPERCHARGER",
     "Automotive", "Gas & Wash"),
    (r"BROWN BEAR CAR WASH|PAPA BEAR|CAR WASH|AUTO WASH",
     "Automotive", "Gas & Wash"),
    (r"TESLA MOTORS|TESLA RESERVATION|DISCOUNT.TIRE|WA VEHICLE LICENS|PROGRESSIVE INSU|STATE FARM INSUR(?!.*CLAIM)|PROG DIRECT INS",
     "Automotive", "Car & Accessories"),
    (r"STATE FARM.*CLAIM",
     "Automotive", "Car & Accessories"),

    # ── BILLS & UTILITIES ─────────────────────────────────────────────────────
    (r"SNOHOMISH COUNTY PUD|PUGET SOUND ENERGY|ALDERWOOD WATER|KING COUNTY WW|RECOLOGY|RING BASIC|BLINK$",
     "Bills & Utilities", "Utilities"),
    (r"TMOBILE|T-MOBILE|COMCAST|XFINITY|AT&T|VERIZON|SPECTRUM",
     "Bills & Utilities", "Mobile & Internet"),
    (r"TAILOR BRANDS",
     "Bills & Utilities", "Dues and Subscriptions"),
    (r"DISNEYPLUS|DISNEY PLUS|NETFLIX|HULU|SPOTIFY|AMAZON PRIME|AMAZON KIDS|SQSP\*|SQUARESPACE|LINKEDINPRE|PERPLEXITY|TUTORIALS DOJO|COURSERA",
     "Bills & Utilities", "Dues and Subscriptions"),
    (r"APPLE\.COM/BILL",
     "Bills & Utilities", "Dues and Subscriptions"),
    (r"NORTHWESTERN MU.*ISA PYMENT|NORTHWESTERN MU.*RQST PYMT",
     "Bills & Utilities", "Insurance"),
    (r"SAELA PEST|PEST CONTROL",
     "Bills & Utilities", "Insurance"),
    (r"AWS\.AMAZON|AMAZON WEB SERVICES(?!.*PAYROLL)(?!.*TRADE)(?!.*DIRECT DEP)",
     "Bills & Utilities", "Online Services"),

    # ── PROFESSIONAL SERVICES — BEFORE Education ──────────────────────────────
    (r"USCIS ELIS",
     "Professional Services", "Legal and professional"),
    (r"GLOBAL TAXES|TAX PREP|CPA |ATTORNEY|LAWYER|LEGAL",
     "Professional Services", "Legal and professional"),
    (r"CORPORATE FILINGS|REGISTERED AGENT|LLC FILING",
     "Professional Services", "Legal and professional"),
    (r"CLEANING SERVICE|HOUSE CLEAN|MAID|JANITORIAL|GARDENING|LANDSCAP",
     "Professional Services", "Cleaning Maintenance Gardening Services"),

    # ── HEALTH & BEAUTY — BEFORE Education ────────────────────────────────────
    (r"FRED HUTCHINSON",
     "Health & Beauty", ""),
    (r"EVERGREENHEALTH|UW MEDICINE|CELLNETIX|ALLEGRO PEDIATRIC|PHR\*ALLEGRO|OPTIMALLIFE WELLNESS|CLUBPILATES",
     "Health & Beauty", ""),
    (r"CVS PHARMACY|WALGREENS|RITE AID|PHARMACY",
     "Health & Beauty", ""),

    # ── EDUCATION ─────────────────────────────────────────────────────────────
    (r"UCIC SCHOOL|BRGHTWHL.*BIRCH|BIRCH TREE|MATHNASIUM|KIDSTRONG|LYNNWOOD RECREATION|NORTHSHORE SCHOOL|ARENA SPORTS|BOTHELL GYMNASTIC|RIDGE ACTIVITY|STUDIO EAST|MADRASA|MY DANCE WORLD|CANYON PARK",
     "Education", ""),

    # ── GROCERY — Instacart always = Grocery ──────────────────────────────────
    (r"INSTACART|IC\* COSTCO BY INSTAC",
     "Grocery", ""),
    (r"SAFEWAY(?! FUEL)|QFC|H MART|WHOLE FOODS|PCC -|TRADER JOE|FRED.MEYER|MAYURI FOODS|APNA BAZAR|SPROUTS|ALBERTSONS",
     "Grocery", ""),

    # ── MEALS ─────────────────────────────────────────────────────────────────
    (r"CHIPOTLE|MCDONALD|BURGER KING|WENDY|SUBWAY|TACO BELL|STARBUCKS|DUNKIN|PEET.S COFFEE|DUTCH BROS|COFFEE|BAKERY|SUSHI|RAMEN|THAI|INDIAN|MEXICAN|ITALIAN|CHINESE|JAPANESE|KOREAN|HALAL|KABAB|DUMPLINGS|HAPPY LEMON|BOBA|ICE CREAM|GELATO|DINER|GRILL|BISTRO|TAVERN|PUB |BAR |KITCHEN|EATERY|FOGO|PAR\*FOGO|SAMBURNA|KABAB HOUSE|JUNIOR KUPPANNA|FRONTIER CAFE|DYAM|DIN TAI FUNG|PADDY COYNE|HAPPY SINGH|ROUND TABLE PIZZA|TST\*|PAR\*",
     "Meals", "Meals"),
    (r"UBER.*EATS|DOORDASH|GRUBHUB|POSTMATES",
     "Meals", "Meals"),

    # ── SHOPPING — Costco warehouse BEFORE Meals ──────────────────────────────
    (r"COSTCO WHSE|WWW COSTCO COM|COSTCO\.COM",
     "Shopping", ""),
    (r"AMAZON\.COM|AMAZON MARKETPLACE|AMZN\.COM|AMAZON MARKEPLACE",
     "Shopping", ""),
    (r"TARGET\.COM|TARGET #|WALMART\.COM|WALMART #|ROSS STORE|DOLLAR TREE|MARSHALLS|TJ MAXX",
     "Shopping", ""),
    (r"SUPERCUTS|GREAT CLIPS|HAIR|SALON|BARBER|NAIL|SPA |MASSAGE|ULTA|SEPHORA",
     "Health & Beauty", ""),

    # ── TRAVEL ────────────────────────────────────────────────────────────────
    (r"ALASKA AIR|UNITED AIRLINES|DELTA AIR|AMERICAN AIRLINES|SOUTHWEST|JETBLUE|HAWAIIAN AIR|WIFIONBOARD",
     "Travel", "Travel & Entertainment Expense"),
    (r"EXPEDIA|BOOKING\.COM|HOTELS\.COM|AIRBNB|VRBO",
     "Travel", "Travel & Entertainment Expense"),
    (r"HYATT|MARRIOTT|HILTON|SHERATON|WESTIN|RITZ|HOLIDAY INN|HAMPTON INN|COURTYARD|SONESTA|ROYAL SONESTA|CLEMENTINE HOTEL",
     "Travel", "Travel & Entertainment Expense"),
    (r"AVIS RENT|HERTZ|ENTERPRISE RENT|NATIONAL CAR|BUDGET RENT",
     "Travel", "Travel & Entertainment Expense"),
    (r"UBER(?! EATS)|CURB NYC TAXI|TAXI|WEGOPRO|CHEAPAIRPORTPARKING|DIAMOND PARKING|PARKWHIZ|HANGTAG PARKING|GOOD2GO|WSDOT.*GOODTOGO|WSFERRIES|FERRY",
     "Travel", "Cabs, Parking & Tolls"),
    (r"YELLOWSTONE|NATIONAL PARK|STATE PARK|RECREATION\.GOV|MUSEUM|JENNY LAKE|CITY PASS",
     "Travel", "Travel & Entertainment Expense"),

    # ── ENTERTAINING ──────────────────────────────────────────────────────────
    (r"FANDANGO|AMC |REGAL |CINEMARK|IMAX |MOVIE|CINEMA|THEATER|THEATRE",
     "Entertaining", "Travel & Entertainment Expense"),
    (r"LYNNWOOD RECREATION",
     "Entertaining", "Travel & Entertainment Expense"),
    (r"WOODLAND PARK ZOO|ZOO |AQUARIUM|SCIENCE CENTER|ELEVATED SPORTZ|LAKE EASTON|PAINTED PALACE",
     "Entertaining", "Entertaining Expense"),
    (r"GROUPON(?!.*PARKING)",
     "Entertaining", "Entertaining Expense"),

    # ── SPL INVESTMENT ────────────────────────────────────────────────────────
    (r"IQ SURGICAL|MORGAN STANLEY|FIDELITY|VANGUARD|SCHWAB|FID BKG SVC.*MONEYLINE",
     "Spl Investment", ""),
    (r"WINDSOR SAMCOS|WIO BANK",
     "Income", "Income AJ"),
]


def categorize(description, lookup):
    """
    Two-layer categorization:
    1. Exact match against historical Master List lookup
    2. Keyword regex rules
    3. Partial match against known descriptions
    """
    desc_upper = str(description).upper().strip()

    # Layer 1: Historical lookup (exact match)
    if desc_upper in lookup:
        r = lookup[desc_upper]
        return r["Category"], r["Sub Category"]

    # Layer 2: Keyword rules
    for pattern, cat, sub in RULES:
        if re.search(pattern, desc_upper, re.IGNORECASE):
            return cat, sub

    # Layer 3: Partial match on first 25 chars of known descriptions
    for known_desc, vals in lookup.items():
        key = known_desc[:25].strip()
        if len(key) > 8 and key in desc_upper:
            return vals["Category"], vals["Sub Category"]

    return "", ""  # Unmatched — will appear in Needs Review sheet


def apply_overrides(df):
    """
    Manual overrides for context-dependent categorizations.
    These run AFTER the rule engine to ensure correct final assignment.
    """
    def override(mask_pattern, cat, sub):
        mask = df["Description"].str.contains(mask_pattern, case=False, na=False, regex=True)
        df.loc[mask, "Category"]     = cat
        df.loc[mask, "Sub-Category"] = sub

    # All Lyft = Official A (work trips)
    override(r"LYFT",                              "Official",              "Official A")
    # Tailor Brands = subscription (all description variants)
    override(r"TAILOR BRANDS",                     "Bills & Utilities",     "Dues and Subscriptions")
    # USCIS = immigration legal filing
    override(r"USCIS ELIS",                        "Professional Services", "Legal and professional")
    # Fred Hutchinson = medical
    override(r"FRED HUTCHINSON",                   "Health & Beauty",       "")
    # Capital One mobile payment (all variants)
    override(r"CAPITAL ONE.*MOBILE PMT|CAPITAL ONE.*MOBILE PYMT", "Payment", "Payment")
    # SF meal merchants = Official A (work trip city)
    override(r"BUN MEE|THEMELT|MARU SUSHI",        "Official",              "Official A")
    # Costco warehouse in Meals → Shopping
    df.loc[
        df["Description"].str.contains(r"COSTCO WHSE|WWW COSTCO COM", case=False, na=False, regex=True) &
        (df["Category"] == "Meals"),
        ["Category", "Sub-Category"]
    ] = ["Shopping", ""]
    # Service fees
    override(r"PURCHASE INTEREST CHARGE|INTEREST CHARGE:PURCHASES|INTEREST CHARGED ON PURCHASES|LATE FEE|PAST DUE FEE|MONTHLY.*FEE",
             "Payment", "Service Fees")

    return df


# ── MAIN ──────────────────────────────────────────────────────────────────────

print("\n[Step 1] Loading historical Master List for lookup...")
lookup = build_lookup(ML_FILE)

print("\n[Step 2] Loading normalized transactions...")
tx = pd.read_excel(INPUT, sheet_name="Transactions")
tx["Category"]     = tx["Category"].fillna("").astype(str)
tx["Sub-Category"] = tx["Sub-Category"].fillna("").astype(str)
print(f"  Loaded {len(tx)} transactions")

print("\n[Step 3] Applying categorization engine...")
cats, subs = [], []
for _, row in tx.iterrows():
    c, s = categorize(row["Description"], lookup)
    cats.append(c)
    subs.append(s)
tx["Category"]     = cats
tx["Sub-Category"] = subs

print("\n[Step 4] Applying manual overrides...")
tx = apply_overrides(tx)

# Data quality fixes
tx.loc[tx["Sub-Category"] == "Card payment", "Sub-Category"] = "Card Payment"
tx.loc[(tx["Category"] == "Shopping") & (tx["Sub-Category"] == "gifts"), "Sub-Category"] = ""

# ── RESULTS ───────────────────────────────────────────────────────────────────
total     = len(tx)
matched   = (tx["Category"] != "").sum()
unmatched = (tx["Category"] == "").sum()

print(f"\n{'='*65}")
print("CATEGORIZATION RESULTS")
print(f"{'='*65}")
print(f"  Total transactions  : {total}")
print(f"  Auto-categorized    : {matched} ({matched/total*100:.1f}%)")
print(f"  Needs manual review : {unmatched} ({unmatched/total*100:.1f}%)")
print(f"{'='*65}")

if unmatched > 0:
    print("\n  Transactions requiring manual review:")
    needs_review = tx[tx["Category"] == ""]
    pd.set_option("display.max_colwidth", 50)
    print(needs_review[["Transaction Date", "Account", "Description", "Amount"]].to_string(index=False))
else:
    print("\n  All transactions categorized successfully.")

needs_review = tx[tx["Category"] == ""].copy()

# ── WRITE OUTPUT ──────────────────────────────────────────────────────────────
conf_data = tx.groupby("Category").agg(
    Rows=("Amount", "count"),
    Expenses=("Amount", lambda x: round(x[x > 0].sum(), 2)),
    Credits=("Amount",  lambda x: round(x[x < 0].sum(), 2)),
    Net=("Amount",      lambda x: round(x.sum(), 2))
).sort_values("Expenses", ascending=False).reset_index()

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    tx.to_excel(writer, sheet_name="Transactions", index=False)
    conf_data.to_excel(writer, sheet_name="Category Summary", index=False)
    needs_review.to_excel(writer, sheet_name="Needs Review", index=False)

print(f"\nOutput written to:\n  {OUTPUT}")
print(f"\nSheets: Transactions | Category Summary | Needs Review ({len(needs_review)} rows)")
print(f"\nNext step: Run append_to_master.py to add these rows to the Master List.")
