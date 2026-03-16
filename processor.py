import streamlit as st
import pandas as pd
import numpy as np
import html
from region_maps import classify_region  # your mapping helper

# ----------------------------- HELPERS -----------------------------
def sanitize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    """
    Decode HTML entities (&gt;,&lt;,&amp; etc.), replace non-breaking spaces, and trim.
    Ensures consistent, clean column names everywhere.
    """
    def _clean(name: str) -> str:
        s = html.unescape(str(name))           # &gt; -> >, &amp; -> &, ...
        s = s.replace("\u00A0", " ").strip()   # normalize NBSP
        return s
    return df.rename(columns={c: _clean(c) for c in df.columns})

def safe_to_datetime(series, column_name=""):
    """Safe to_datetime with row-level error display."""
    if series is None:
        return pd.to_datetime(pd.Series([], dtype="object"), errors="coerce")

    series_str = series.astype(str).str.strip().replace({"nan": np.nan, "": np.nan})
    dt = pd.to_datetime(series_str, errors="coerce")
    invalid_mask = dt.isna() & series_str.notna()
    if invalid_mask.any():
        st.error(f"⚠️ Invalid datetime values in column '{column_name}':")
        for idx, val in series_str[invalid_mask].items():
            st.write(f"Row {idx + 16} → {val}")  # header=15 -> +16
    return dt

# ----------------------------- MAIN PROCESSOR -----------------------------
def process_ar_file(file):
    excel = pd.ExcelFile(file, engine="openpyxl")
    sheet = excel.sheet_names[0]

    # ---------- READ B14 (As on Date) ----------
    header_block = pd.read_excel(excel, sheet_name=sheet, header=None, nrows=14)
    as_on_date = header_block.iloc[13, 1]
    if pd.isna(as_on_date):
        raise ValueError("Cell B14 must contain 'As on Date'")
    as_on_date = pd.to_datetime(str(as_on_date).strip(), errors="coerce")
    if pd.isna(as_on_date):
        raise ValueError("As on Date in B14 is not a valid date")

    # ---------- READ MAIN TABLE ----------
    df = pd.read_excel(excel, sheet_name=sheet, header=15, dtype=str)
    df = sanitize_colnames(df)  # normalize input header names

    # Keep original order for Solution C
    original_order = list(df.columns)

    # ---------- SAFE CONVERSIONS ----------
    ar_balance = pd.to_numeric(df.get("Ar Balance"), errors="coerce").fillna(0)
    doc_date_raw = df.get("Document Date")
    due_date_raw = df.get("Document Due Date")
    doc_date = safe_to_datetime(doc_date_raw, "Document Date")
    due_date = safe_to_datetime(due_date_raw, "Document Due Date")

    if len(doc_date) != len(df):
        doc_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")
    if len(due_date) != len(df):
        due_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")

    doc_date_filled = doc_date.fillna(as_on_date)
    due_date_filled = due_date.fillna(as_on_date)

    # ---------- AGEING / OVERDUE ----------
    ageing_days = (as_on_date - doc_date_filled).dt.days
    overdue_days = (as_on_date - due_date_filled).dt.days
    overdue_num = pd.to_numeric(overdue_days, errors="coerce").fillna(0)

    # ---------- REGION ----------
    if "Cust Region" in df.columns and "Cust Code" in df.columns:
        region_series = classify_region(df["Cust Region"], df["Cust Code"])
    elif "Cust Region" in df.columns:
        region_series = classify_region(df["Cust Region"])
    elif "Cust Code" in df.columns:
        empty_region = pd.Series([""] * len(df), index=df.index)
        region_series = classify_region(empty_region, df["Cust Code"])
    else:
        region_series = pd.Series(["KSA"] * len(df), index=df.index)

    # ---------- STATUS ----------
    if "Customer Status" in df.columns:
        updated_status = df["Customer Status"].fillna("").replace("", "SUBSTANDARD")
    else:
        updated_status = pd.Series(["SUBSTANDARD"] * len(df), index=df.index)

    # ---------- AGING BRACKET (label only) ----------
    conditions = [
        ar_balance < 0,                                # On account priority
        overdue_num < 0,                               # Not due
        (overdue_num >= 0) & (overdue_num <= 30),      # 0–30
        (overdue_num > 30) & (overdue_num <= 60),      # 31–60
        (overdue_num > 60) & (overdue_num <= 90),      # 61–90
        (overdue_num > 90) & (overdue_num <= 120),     # 91–120
        (overdue_num > 120) & (overdue_num <= 150),    # 121–150
        overdue_num > 150                              # >150
    ]
    choices = [
        "On account",
        "Not Due",
        "Aging 1 to 30",
        "Aging 31 to 60",
        "Aging 61 to 90",
        "Aging 91 to 120",
        "Aging 121 to 150",
        "Aging >=151"
    ]
    aging_bracket_label = np.select(conditions, choices, default="")

    # ---------- DERIVED AMOUNTS ----------
    invoice_value = ar_balance.clip(lower=0)               # IF(Ar Balance (Copy)>0, copy, 0)
    on_account_amount = ar_balance.clip(upper=0)           # IF(Ar Balance (Copy)<0, copy, 0)
    not_due_amount = np.where(overdue_num > 0, 0, invoice_value)  # IF(Overdue>0,0,InvoiceValue)

    # ---------- AMOUNT buckets (residual / waterfall) ----------
    BP = invoice_value
    BK = overdue_num
    amt_ge151   = np.where(BK > 150, BP, 0)
    amt_121_150 = np.where(BK > 120, BP, 0) - amt_ge151
    amt_91_120  = np.where(BK > 90,  BP, 0) - amt_121_150 - amt_ge151
    amt_61_90   = np.where(BK > 60,  BP, 0) - amt_91_120 - amt_121_150 - amt_ge151
    amt_31_60   = np.where(BK > 30,  BP, 0) - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151
    amt_1_30    = np.where(BK >= 0,  BP, 0) - amt_31_60 - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151

    for arr in [amt_1_30, amt_31_60, amt_61_90, amt_91_120, amt_121_150, amt_ge151]:
        np.maximum(arr, 0, out=arr)  # clamp tiny negatives

    # ---------- ASSIGN COLUMNS ----------
    df["Ageing (Days)"]            = ageing_days
    df["Overdue days (Days)"]      = overdue_num
    df["Region (Derived)"]         = region_series
    df["Aging Bracket (Label)"]    = aging_bracket_label
    df["Updated Status"]           = updated_status

    df["Invoice Value (Derived)"]  = invoice_value
    df["On Account (Derived)"]     = on_account_amount
    df["Not Due (Derived)"]        = not_due_amount

    # AMOUNT buckets (residual)
    df["Aging 1 to 30 (Amount)"]   = amt_1_30
    df["Aging 31 to 60 (Amount)"]  = amt_31_60
    df["Aging 61 to 90 (Amount)"]  = amt_61_90
    df["Aging 91 to 120 (Amount)"] = amt_91_120
    df["Aging 121 to 150 (Amount)"]= amt_121_150
    df["Aging >=151 (Amount)"]     = amt_ge151

    # Extra amount (used in summary)
    df["Ageing > 365 (Amt)"]       = np.where(df["Ageing (Days)"] > 365, ar_balance, 0)

    # Copy of Ar Balance for Excel formula references
    df["Ar Balance (Copy)"]        = ar_balance

    # ---------- ORDER ----------
    appended_block = [
        "Ageing (Days)",
        "Overdue days (Days)",
        "Region (Derived)",
        "Ar Balance (Copy)",
        "Aging Bracket (Label)",
        "Updated Status",
        "Invoice Value (Derived)",
        "On Account (Derived)",
        "Not Due (Derived)",
        "Aging 1 to 30 (Amount)",
        "Aging 31 to 60 (Amount)",
        "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)",
        "Aging 121 to 150 (Amount)",
        "Aging >=151 (Amount)",
        "Ageing > 365 (Amt)",
    ]
    for col in appended_block:
        if col not in df.columns:
            df[col] = ""

    df = df[list(original_order) + appended_block]
    df.attrs["as_on_date"] = as_on_date
    return df

def customer_summary(df):
    df = sanitize_colnames(df)
    out = df.copy()
    out = out.loc[:, ~out.columns.duplicated(keep="last")]  # dedupe

    # ---- Keys / descriptors ----
    if "Cust Code" in out.columns:
        out["Cust Code"] = out["Cust Code"].astype(str).str.strip()
    if "Main Ac" in out.columns:
        out["Main Ac"] = out["Main Ac"].fillna("").astype(str).str.strip()
    else:
        out["Main Ac"] = ""

    if "Region" not in out.columns:
        if "Region (Derived)" in out.columns:
            out["Region"] = out["Region (Derived)"]
        elif "Cust Region" in out.columns:
            out["Region"] = out["Cust Region"]
        else:
            out["Region"] = ""

    # ---- Require template column(s) (strict) ----
    required_cols = ["Not Due Amount"]
    missing = [c for c in required_cols if c not in out.columns]
    if missing:
        raise ValueError(f"Template error: missing required column(s): {missing}. "
                         "Please upload the standard AR Backlog template.")

    # ---- Numeric coercion for safe aggregation (values only; By_Customer has no formulas) ----
    numeric_cols = [
        "On Account (Derived)",
        "Aging 1 to 30 (Amount)",
        "Aging 31 to 60 (Amount)",
        "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)",
        "Aging 121 to 150 (Amount)",
        "Aging >=151 (Amount)",
        "Ageing > 365 (Amt)",
        "Overdue days (Days)",
        "Not Due Amount",
        "Ar Balance (Copy)",   # for AR Balance aggregation if original is missing
    ]
    for c in numeric_cols:
        if c not in out.columns:
            out[c] = 0
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)

    # ---- Quarter pivot YEAR: FIXED to 2026 ----
    yr = 2026

    # Parse Document Due Date: strip time, clean hidden chars, parse as ISO date
    if "Document Due Date" in out.columns:
        due_raw = out["Document Due Date"].astype(str)
        due_norm = (
            due_raw
            .str.replace("\u00A0", " ", regex=False)
            .str.replace(r"[^\x00-\x7F]", "", regex=True)
            .str.strip()
        )
        due_date_only = due_norm.str.replace(r"\s+\d{2}:\d{2}:\d{2}$", "", regex=True)
        due_dt = pd.to_datetime(due_date_only, format="%Y-%m-%d", errors="coerce")
    else:
        due_dt = pd.Series([pd.NaT] * len(out), index=out.index)

    # ---- VALUE SOURCE for quarters: Ar Balance or Ar Balance (Copy) ----
    if "Ar Balance" in out.columns:
        ar_bal_num = pd.to_numeric(out["Ar Balance"], errors="coerce").fillna(0)
    else:
        ar_bal_num = pd.to_numeric(out.get("Ar Balance (Copy)", 0), errors="coerce").fillna(0)

    # ---- Quarter model (inclusive) with your custom split ----
    quarters = {
        "Q1": (pd.Timestamp(year=yr, month=1,  day=1),  pd.Timestamp(year=yr, month=3,  day=15)),
        "Q2": (pd.Timestamp(year=yr, month=3,  day=16), pd.Timestamp(year=yr, month=6,  day=15)),
        "Q3": (pd.Timestamp(year=yr, month=6,  day=16), pd.Timestamp(year=yr, month=9,  day=15)),
        "Q4": (pd.Timestamp(year=yr, month=9,  day=16), pd.Timestamp(year=yr, month=12, day=15)),
    }

    # ---- Per-row quarter values (value: Ar Balance) ----
    quarter_cols = []
    for q, (start, end) in quarters.items():
        col = f"{q}-{yr} - pivot"
        mask = (due_dt >= start) & (due_dt <= end)
        out[col] = np.where(mask, ar_bal_num, 0.0)
        quarter_cols.append(col)

    # ---- Fixed future year buckets (values by Document Due Date year): 2027-2030 ----
    year_targets = [2027, 2028, 2029, 2030]
    due_dt_year = due_dt.dt.year

    year_cols = []
    for y in year_targets:
        col = f"{y}"
        out[col] = np.where(due_dt_year == y, ar_bal_num, 0.0)
        year_cols.append(col)

    # ---- Aggregation mapping (values only) ----
    agg_dict = {
        "Cust Name": ("Cust Name", "first"),
        "Cust Region": ("Cust Region", "first"),
        "Region": ("Region", "first"),
        "Updated Status": ("Updated Status", "first"),

        "On Account (Derived)": ("On Account (Derived)", "sum"),
        "Not Due Amount": ("Not Due Amount", "sum"),      # strict as requested
        "AR Balance": ("Ar Balance (Copy)", "sum"),

        "Overdue days (Days)": ("Overdue days (Days)", "sum"),
        "Aging 1 to 30 (Amount)": ("Aging 1 to 30 (Amount)", "sum"),
        "Aging 31 to 60 (Amount)": ("Aging 31 to 60 (Amount)", "sum"),
        "Aging 61 to 90 (Amount)": ("Aging 61 to 90 (Amount)", "sum"),
        "Aging 91 to 120 (Amount)": ("Aging 91 to 120 (Amount)", "sum"),
        "Aging 121 to 150 (Amount)": ("Aging 121 to 150 (Amount)", "sum"),
        "Aging >=151 (Amount)": ("Aging >=151 (Amount)", "sum"),

        "Ageing > 365 (Amt)": ("Ageing > 365 (Amt)", "sum"),
    }
    for qc in quarter_cols:
        agg_dict[qc] = (qc, "sum")
    for yc in year_cols:
        agg_dict[yc] = (yc, "sum")

    grouped = (
        out.groupby(["Cust Code", "Main Ac"], as_index=False)
           .agg(**agg_dict)
    )

    # ---- Overdue (Days) in summary = sum of six AMOUNT buckets (values) ----
    amount_buckets = [
        "Aging 1 to 30 (Amount)",
        "Aging 31 to 60 (Amount)",
        "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)",
        "Aging 121 to 150 (Amount)",
        "Aging >=151 (Amount)",
    ]
    present_amounts = [c for c in amount_buckets if c in grouped.columns]
    if present_amounts:
        grouped["Overdue days (Days)"] = grouped[present_amounts].sum(axis=1)
    else:
        if "Overdue days (Days)" not in grouped.columns:
            grouped["Overdue days (Days)"] = 0

    # ---- Column order ----
    base_order = [
        "Cust Code", "Main Ac", "Cust Name", "Cust Region", "Region", "Updated Status",
        "On Account (Derived)", "Not Due Amount", "AR Balance",
        "Overdue days (Days)",
        "Aging 1 to 30 (Amount)", "Aging 31 to 60 (Amount)", "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)", "Aging 121 to 150 (Amount)", "Aging >=151 (Amount)",
        "Ageing > 365 (Amt)",
    ] + quarter_cols + year_cols

    grouped = grouped[[c for c in base_order if c in grouped.columns] +
                      [c for c in grouped.columns if c not in base_order]]

    return grouped

# ----------------------------- NEW: INVOICE SHEET -----------------------------
def invoice_summary(df):
    """
    Build a per-invoice sheet with the exact columns requested.
    Output columns (in order):
      Cust Code, Cust Name, Main Ac, Cust Region, Region,
      Document Number, Document Date, Document Due Date,
      Ageing, Overdue days, Payment Terms,
      On Account, Not Due, Ar Balance,
      Aging 1 to 30, Aging 31 to 60, Aging 61 to 90, Aging 91 to 120, Aging 121 to 150, Aging >=151,
      Brand, Total Insurance Limit, LC & BG Guarantee, SO No, LPO No
    """
    work = sanitize_colnames(df).copy()
    work = work.loc[:, ~work.columns.duplicated(keep="last")]

    # Region fallback like in customer_summary
    if "Region" not in work.columns:
        if "Region (Derived)" in work.columns:
            work["Region"] = work["Region (Derived)"]
        elif "Cust Region" in work.columns:
            work["Region"] = work["Cust Region"]
        else:
            work["Region"] = ""

    # Choose the source for Ar Balance
    if "Ar Balance (Copy)" in work.columns:
        ar_source_col = "Ar Balance (Copy)"
    elif "Ar Balance" in work.columns:
        ar_source_col = "Ar Balance"
    else:
        work["Ar Balance (Copy)"] = 0
        ar_source_col = "Ar Balance (Copy)"

    # Numeric mappings (source -> target)
    numeric_map = {
        "Ageing (Days)": "Ageing",
        "Overdue days (Days)": "Overdue days",
        "On Account (Derived)": "On Account",
        "Not Due (Derived)": "Not Due",
        ar_source_col: "Ar Balance",
        "Aging 1 to 30 (Amount)": "Aging 1 to 30",
        "Aging 31 to 60 (Amount)": "Aging 31 to 60",
        "Aging 61 to 90 (Amount)": "Aging 61 to 90",
        "Aging 91 to 120 (Amount)": "Aging 91 to 120",
        "Aging 121 to 150 (Amount)": "Aging 121 to 150",
        "Aging >=151 (Amount)": "Aging >=151",
    }

    out = pd.DataFrame()

    # --- Direct, non-numeric mappings (as-is) ---
    def copy_if_exists(src, dst):
        out[dst] = work[src] if src in work.columns else ""

    copy_if_exists("Cust Code", "Cust Code")
    copy_if_exists("Cust Name", "Cust Name")
    copy_if_exists("Main Ac", "Main Ac")
    copy_if_exists("Cust Region", "Cust Region")
    copy_if_exists("Region", "Region")
    copy_if_exists("Document Number", "Document Number")
    copy_if_exists("Document Date", "Document Date")
    copy_if_exists("Document Due Date", "Document Due Date")
    copy_if_exists("Payment Terms", "Payment Terms")
    copy_if_exists("Brand", "Brand")
    copy_if_exists("Total Insurance Limit", "Total Insurance Limit")
    copy_if_exists("LC & BG Guarantee", "LC & BG Guarantee")
    copy_if_exists("SO No", "SO No")
    copy_if_exists("LPO No", "LPO No")

    # --- Numeric mappings with coercion ---
    for src, dst in numeric_map.items():
        if src not in work.columns:
            out[dst] = 0
        else:
            out[dst] = pd.to_numeric(work[src], errors="coerce").fillna(0)

    # --- Column order exactly as requested ---
    final_order = [
        "Cust Code", "Cust Name", "Main Ac", "Cust Region", "Region",
        "Document Number", "Document Date", "Document Due Date",
        "Ageing", "Overdue days", "Payment Terms",
        "On Account", "Not Due", "Ar Balance",
        "Aging 1 to 30", "Aging 31 to 60", "Aging 61 to 90", "Aging 91 to 120",
        "Aging 121 to 150", "Aging >=151",
        "Brand", "Total Insurance Limit", "LC & BG Guarantee", "SO No", "LPO No",
    ]

    for c in final_order:
        if c not in out.columns:
            out[c] = ""  # ensure presence

    out = out[final_order]
    return out