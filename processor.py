
import streamlit as st
import pandas as pd
import numpy as np
import html
from region_maps import classify_region  # mapping helper

# ----------------------------- HELPERS -----------------------------
def sanitize_colnames(df: pd.DataFrame) -> pd.DataFrame:
    """Decode HTML entities, remove NBSP, strip whitespace from column names."""
    def _clean(name: str) -> str:
        s = html.unescape(str(name))
        s = s.replace("\u00A0", " ").strip()
        return s
    return df.rename(columns={c: _clean(c) for c in df.columns})

def safe_to_datetime(series, column_name: str = "") -> pd.Series:
    """Parse dates with Streamlit error reporting for bad rows."""
    if series is None:
        return pd.to_datetime(pd.Series([], dtype="object"), errors="coerce")
    series_str = series.astype(str).str.strip().replace({"nan": np.nan, "": np.nan})
    dt = pd.to_datetime(series_str, errors="coerce")
    bad = dt.isna() & series_str.notna()
    if bad.any():
        st.error(f"⚠️ Invalid datetime values in column '{column_name}':")
        for idx, val in series_str[bad].items():
            st.write(f"Row {idx + 16} → {val}")
    return dt

# ----------------------------- MAIN PROCESSOR -----------------------------
def process_ar_file(file):
    excel = pd.ExcelFile(file, engine="openpyxl")
    sheet = excel.sheet_names[0]

    # --- As on Date from B14 ---
    header_block = pd.read_excel(excel, sheet_name=sheet, header=None, nrows=14)
    as_on_date = header_block.iloc[13, 1]
    if pd.isna(as_on_date):
        raise ValueError("Cell B14 must contain 'As on Date'")
    as_on_date = pd.to_datetime(str(as_on_date).strip(), errors="coerce")
    if pd.isna(as_on_date):
        raise ValueError("As on Date in B14 is not a valid date")

    # --- Main table starts header row 16 (index 15) ---
    df = pd.read_excel(excel, sheet_name=sheet, header=15, dtype=str)
    df = sanitize_colnames(df)
    original_order = list(df.columns)

    # --- Safe conversions ---
    ar_balance = pd.to_numeric(df.get("Ar Balance"), errors="coerce").fillna(0)
    doc_date = safe_to_datetime(df.get("Document Date"), "Document Date")
    due_date = safe_to_datetime(df.get("Document Due Date"), "Document Due Date")
    if len(doc_date) != len(df):
        doc_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")
    if len(due_date) != len(df):
        due_date = pd.to_datetime(pd.Series([pd.NaT] * len(df)), errors="coerce")

    doc_date_filled = doc_date.fillna(as_on_date)
    due_date_filled = due_date.fillna(as_on_date)

    # --- Ageing / overdue ---
    ageing_days = (as_on_date - doc_date_filled).dt.days
    overdue_days = (as_on_date - due_date_filled).dt.days
    overdue_num = pd.to_numeric(overdue_days, errors="coerce").fillna(0)

    # --- Region ---
    if "Cust Region" in df.columns and "Cust Code" in df.columns:
        region_series = classify_region(df["Cust Region"], df["Cust Code"])
    elif "Cust Region" in df.columns:
        region_series = classify_region(df["Cust Region"])
    elif "Cust Code" in df.columns:
        empty_region = pd.Series([""] * len(df), index=df.index)
        region_series = classify_region(empty_region, df["Cust Code"])
    else:
        region_series = pd.Series(["KSA"] * len(df), index=df.index)

    # --- Status ---
    if "Customer Status" in df.columns:
        updated_status = df["Customer Status"].fillna("").replace("", "SUBSTANDARD")
    else:
        updated_status = pd.Series(["SUBSTANDARD"] * len(df), index=df.index)

    # --- Ageing bracket label ---
    conditions = [
        ar_balance < 0,
        overdue_num < 0,
        (overdue_num >= 0) & (overdue_num <= 30),
        (overdue_num > 30) & (overdue_num <= 60),
        (overdue_num > 60) & (overdue_num <= 90),
        (overdue_num > 90) & (overdue_num <= 120),
        (overdue_num > 120) & (overdue_num <= 150),
        overdue_num > 150,
    ]
    choices = [
        "On account",
        "Not Due",
        "Aging 1 to 30",
        "Aging 31 to 60",
        "Aging 61 to 90",
        "Aging 91 to 120",
        "Aging 121 to 150",
        "Aging >=151",
    ]
    aging_bracket_label = np.select(conditions, choices, default="")

    # --- Derived amounts ---
    invoice_value = ar_balance.clip(lower=0)
    on_account_amount = ar_balance.clip(upper=0)
    not_due_amount = np.where(overdue_num > 0, 0, invoice_value)

    # --- Residual waterfall buckets ---
    BP, BK = invoice_value, overdue_num
    amt_ge151   = np.where(BK > 150, BP, 0)
    amt_121_150 = np.where(BK > 120, BP, 0) - amt_ge151
    amt_91_120  = np.where(BK > 90,  BP, 0) - amt_121_150 - amt_ge151
    amt_61_90   = np.where(BK > 60,  BP, 0) - amt_91_120 - amt_121_150 - amt_ge151
    amt_31_60   = np.where(BK > 30,  BP, 0) - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151
    amt_1_30    = np.where(BK >= 0,  BP, 0) - amt_31_60 - amt_61_90 - amt_91_120 - amt_121_150 - amt_ge151
    for arr in [amt_1_30, amt_31_60, amt_61_90, amt_91_120, amt_121_150, amt_ge151]:
        np.maximum(arr, 0, out=arr)

    # --- Assign columns ---
    df["Ageing (Days)"]             = ageing_days
    df["Overdue days (Days)"]       = overdue_num
    df["Region (Derived)"]          = region_series
    df["Aging Bracket (Label)"]     = aging_bracket_label
    df["Updated Status"]            = updated_status
    df["Invoice Value (Derived)"]   = invoice_value
    df["On Account (Derived)"]      = on_account_amount
    df["Not Due (Derived)"]         = not_due_amount
    df["Aging 1 to 30 (Amount)"]    = amt_1_30
    df["Aging 31 to 60 (Amount)"]   = amt_31_60
    df["Aging 61 to 90 (Amount)"]   = amt_61_90
    df["Aging 91 to 120 (Amount)"]  = amt_91_120
    df["Aging 121 to 150 (Amount)"] = amt_121_150
    df["Aging >=151 (Amount)"]      = amt_ge151
    df["Ageing > 365 (Amt)"]        = np.where(df["Ageing (Days)"] > 365, ar_balance, 0)
    df["Ar Balance (Copy)"]         = ar_balance

    # --- Order ---
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

# ----------------------------- CUSTOMER SUMMARY (FORECAST) -----------------------------
def customer_summary(df):
    df = sanitize_colnames(df)
    out = df.copy()
    out = out.loc[:, ~out.columns.duplicated(keep="last")]

    # Keys
    out["Cust Code"] = out.get("Cust Code", "").astype(str).str.strip()
    out["Main Ac"] = out.get("Main Ac", "").fillna("").astype(str).str.strip()

    if "Region" not in out.columns:
        if "Region (Derived)" in out.columns:
            out["Region"] = out["Region (Derived)"]
        elif "Cust Region" in out.columns:
            out["Region"] = out["Cust Region"]
        else:
            out["Region"] = ""

    # Required template column
    if "Not Due Amount" not in out.columns:
        raise ValueError("Template error: missing column 'Not Due Amount'. Upload the standard AR Backlog template.")

    # Numeric coercion
    for c in [
        "On Account (Derived)",
        "Aging 1 to 30 (Amount)", "Aging 31 to 60 (Amount)", "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)", "Aging 121 to 150 (Amount)", "Aging >=151 (Amount)",
        "Ageing > 365 (Amt)", "Overdue days (Days)", "Not Due Amount", "Ar Balance (Copy)"
    ]:
        out[c] = pd.to_numeric(out.get(c, 0), errors="coerce").fillna(0)

    # Due date parsing (as yyyy-mm-dd text or datetime)
    if "Document Due Date" in out.columns:
        due_raw = out["Document Due Date"].astype(str).str.strip()
        due_raw = due_raw.replace("\u00A0", " ", regex=False)
        due_raw = due_raw.str.replace(r"[^\x00-\x7F]", "", regex=True)
        due_raw = due_raw.str.replace(r"\s+\d{2}:\d{2}:\d{2}$", "", regex=True)
        due_dt = pd.to_datetime(due_raw, errors="coerce")
    else:
        due_dt = pd.Series([pd.NaT] * len(out))

    # Quarter windows (fixed 2026)
    yr = 2026
    Q1_START, Q1_END = pd.Timestamp(yr,1,1),  pd.Timestamp(yr,3,15)
    Q2_START, Q2_END = pd.Timestamp(yr,3,16), pd.Timestamp(yr,6,15)
    Q3_START, Q3_END = pd.Timestamp(yr,6,16), pd.Timestamp(yr,9,15)
    Q4_START, Q4_END = pd.Timestamp(yr,9,16), pd.Timestamp(yr,12,15)

    ar_val = out["Ar Balance (Copy)"]

    # Existing pivots (keep Q1-2026 - pivot and rename others later in export)
    out["Q1-2026 - pivot"] = np.where((due_dt >= Q1_START) & (due_dt <= Q1_END), ar_val, 0)
    out["Q2-2026 - pivot"] = np.where((due_dt >= Q2_START) & (due_dt <= Q2_END), ar_val, 0)
    out["Q3-2026 - pivot"] = np.where((due_dt >= Q3_START) & (due_dt <= Q3_END), ar_val, 0)
    out["Q4-2026 - pivot"] = np.where((due_dt >= Q4_START) & (due_dt <= Q4_END), ar_val, 0)

    # --- Forecast buckets (Python) ---
    out["Q1-2026"] = np.where(due_dt <= Q1_END, ar_val, 0)
    out["16.03.2026..31.03.2026"] = np.where((due_dt >= pd.Timestamp(yr,3,16)) & (due_dt <= pd.Timestamp(yr,3,31)), ar_val, 0)
    out["16.06.2026..30.06.2026"] = np.where((due_dt >= pd.Timestamp(yr,6,16)) & (due_dt <= pd.Timestamp(yr,6,30)), ar_val, 0)
    out["16.09.2026..30.09.2026"] = np.where((due_dt >= pd.Timestamp(yr,9,16)) & (due_dt <= pd.Timestamp(yr,9,30)), ar_val, 0)
    out["16.12.2026..31.12.2026"] = np.where((due_dt >= pd.Timestamp(yr,12,16)) & (due_dt <= pd.Timestamp(yr,12,31)), ar_val, 0)

    # Year buckets
    due_year = due_dt.dt.year
    for y in [2027, 2028, 2029, 2030]:
        out[str(y)] = np.where(due_year == y, ar_val, 0)

    # Aggregate by keys
    agg_map = {
        "Cust Name": ("Cust Name", "first"),
        "Cust Region": ("Cust Region", "first"),
        "Region": ("Region", "first"),
        "Updated Status": ("Updated Status", "first"),
        "On Account (Derived)": ("On Account (Derived)", "sum"),
        "Not Due Amount": ("Not Due Amount", "sum"),
        "AR Balance": ("Ar Balance (Copy)", "sum"),
        "Overdue days (Days)": ("Overdue days (Days)", "sum"),
        "Aging 1 to 30 (Amount)": ("Aging 1 to 30 (Amount)", "sum"),
        "Aging 31 to 60 (Amount)": ("Aging 31 to 60 (Amount)", "sum"),
        "Aging 61 to 90 (Amount)": ("Aging 61 to 90 (Amount)", "sum"),
        "Aging 91 to 120 (Amount)": ("Aging 91 to 120 (Amount)", "sum"),
        "Aging 121 to 150 (Amount)": ("Aging 121 to 150 (Amount)", "sum"),
        "Aging >=151 (Amount)": ("Aging >=151 (Amount)", "sum"),
        "Ageing > 365 (Amt)": ("Ageing > 365 (Amt)", "sum"),
        # forecast sums
        "Q1-2026": ("Q1-2026", "sum"),
        "16.03.2026..31.03.2026": ("16.03.2026..31.03.2026", "sum"),
        "16.06.2026..30.06.2026": ("16.06.2026..30.06.2026", "sum"),
        "16.09.2026..30.09.2026": ("16.09.2026..30.09.2026", "sum"),
        "16.12.2026..31.12.2026": ("16.12.2026..31.12.2026", "sum"),
        # pivots
        "Q1-2026 - pivot": ("Q1-2026 - pivot", "sum"),
        "Q2-2026 - pivot": ("Q2-2026 - pivot", "sum"),
        "Q3-2026 - pivot": ("Q3-2026 - pivot", "sum"),
        "Q4-2026 - pivot": ("Q4-2026 - pivot", "sum"),
        # years
        "2027": ("2027", "sum"),
        "2028": ("2028", "sum"),
        "2029": ("2029", "sum"),
        "2030": ("2030", "sum"),
    }

    grouped = out.groupby(["Cust Code", "Main Ac"], as_index=False).agg(**agg_map)

    # Overdue (Days) in summary = sum of amount buckets
    amount_buckets = [
        "Aging 1 to 30 (Amount)", "Aging 31 to 60 (Amount)", "Aging 61 to 90 (Amount)",
        "Aging 91 to 120 (Amount)", "Aging 121 to 150 (Amount)", "Aging >=151 (Amount)"
    ]
    present = [c for c in amount_buckets if c in grouped.columns]
    grouped["Overdue days (Days)"] = grouped[present].sum(axis=1) if present else 0

    # Rename for final display
    rename_final = {
        "On Account (Derived)": "On account",
        "Not Due Amount": "Not Due",
        "AR Balance": "Ar Balance",
        "Overdue days (Days)": "Overdue",
        "Aging 1 to 30 (Amount)": "Aging 1 to 30",
        "Aging 31 to 60 (Amount)": "Aging 31 to 60",
        "Aging 61 to 90 (Amount)": "Aging 61 to 90",
        "Aging 91 to 120 (Amount)": "Aging 91 to 120",
        "Aging 121 to 150 (Amount)": "Aging 121 to 150",
        "Aging >=151 (Amount)": "Aging >=151",
        "Ageing > 365 (Amt)": "Ageing > 365",
        "Q2-2026 - pivot": "Q2-2026",
        "Q3-2026 - pivot": "Q3-2026",
        "Q4-2026 - pivot": "Q4-2026",
    }
    grouped = grouped.rename(columns=rename_final)

    # Ensure manual/formula columns exist
    for col in ["% for Q1", "Actual Q1", "Remaining % from q1", "To add to Q2", "Forecast Q2"]:
        if col not in grouped.columns:
            grouped[col] = 0

    # Final order
    final_order = [
        "Cust Code", "Cust Name", "Main Ac", "Cust Region", "Region", "Updated Status",
        "On account", "Not Due", "Ar Balance", "Overdue",
        "Aging 1 to 30", "Aging 31 to 60", "Aging 61 to 90", "Aging 91 to 120",
        "Aging 121 to 150", "Aging >=151", "Ageing > 365",
        "Q1-2026 - pivot", "Q1-2026", "% for Q1", "Actual Q1",
        "16.03.2026..31.03.2026", "16.06.2026..30.06.2026", "16.09.2026..30.09.2026", "16.12.2026..31.12.2026",
        "Remaining % from q1", "To add to Q2", "Forecast Q2",
        "Q2-2026", "Q3-2026", "Q4-2026", "2027", "2028", "2029", "2030"
    ]
    for c in final_order:
        if c not in grouped.columns:
            grouped[c] = 0
    return grouped[final_order]

# ----------------------------- INVOICE SUMMARY -----------------------------
def invoice_summary(df):
    work = sanitize_colnames(df).copy()
    work = work.loc[:, ~work.columns.duplicated(keep="last")]

    if "Region" not in work.columns:
        if "Region (Derived)" in work.columns:
            work["Region"] = work["Region (Derived)"]
        elif "Cust Region" in work.columns:
            work["Region"] = work["Cust Region"]
        else:
            work["Region"] = ""

    if "Ar Balance (Copy)" in work.columns:
        ar_source = "Ar Balance (Copy)"
    elif "Ar Balance" in work.columns:
        ar_source = "Ar Balance"
    else:
        work["Ar Balance (Copy)"] = 0
        ar_source = "Ar Balance (Copy)"

    numeric_map = {
        "Ageing (Days)": "Ageing",
        "Overdue days (Days)": "Overdue days",
        "On Account (Derived)": "On Account",
        "Not Due (Derived)": "Not Due",
        ar_source: "Ar Balance",
        "Aging 1 to 30 (Amount)": "Aging 1 to 30",
        "Aging 31 to 60 (Amount)": "Aging 31 to 60",
        "Aging 61 to 90 (Amount)": "Aging 61 to 90",
        "Aging 91 to 120 (Amount)": "Aging 91 to 120",
        "Aging 121 to 150 (Amount)": "Aging 121 to 150",
        "Aging >=151 (Amount)": "Aging >=151",
    }

    out = pd.DataFrame()

    def copy_if_exists(src, dst):
        out[dst] = work[src] if src in work.columns else ""

    for src in [
        "Cust Code", "Cust Name", "Main Ac", "Cust Region", "Region",
        "Document Number", "Document Date", "Document Due Date",
        "Payment Terms", "Brand", "Total Insurance Limit", "LC & BG Guarantee",
        "SO No", "LPO No",
    ]:
        copy_if_exists(src, src)

    for src, dst in numeric_map.items():
        out[dst] = pd.to_numeric(work.get(src, 0), errors="coerce").fillna(0)

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
            out[c] = ""
    return out[final_order]
