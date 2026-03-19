# budg/bud2026_mapper.py
import pandas as pd
import numpy as np

try:
    from region_maps import classify_region
except Exception:
    classify_region = None

# ====================== HELPERS ======================

def _series_or_empty(df: pd.DataFrame, col: str) -> pd.Series:
    """Return column as Series if exists, else empty string Series"""
    if col in df.columns:
        return df[col]
    return pd.Series([""] * len(df), index=df.index)

def _num(df: pd.DataFrame, col: str) -> pd.Series:
    """Coerce to numeric; missing or invalid -> 0"""
    if not col or col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype="float64")
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

def _derive_sales_budget_region(df_cust: pd.DataFrame) -> pd.Series:
    """Derive 'Sales Budget region' robustly"""
    if "Region" in df_cust.columns:
        reg = df_cust["Region"].fillna("").astype(str)
        if reg.str.strip().any():
            return reg
    if classify_region is not None and "Cust Region" in df_cust.columns:
        cust_code = df_cust.get("Cust Code", None)
        derived = classify_region(df_cust["Cust Region"], cust_code)
        return derived.fillna("")
    return pd.Series([""] * len(df_cust), index=df_cust.index)

def _first_present(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first column name from candidates that exists"""
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ====================== MAIN MAPPER ======================

def map_by_customer_to_bud2026(df_customer: pd.DataFrame, ins_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Map input customer DataFrame to BUD2026 structure:
      - Identifiers
      - Insurance (from master)
      - AR / Aging columns
      - AR Balance
    """
    work = df_customer.copy()
    out = pd.DataFrame(index=work.index)

    # ---------------- Identifiers ----------------
    out["CustCode"]            = _series_or_empty(work, "Cust Code").astype(str).str.strip()
    out["Cust Name"]           = _series_or_empty(work, "Cust Name").astype(str)
    out["BT"]                  = ""
    out["Sales Budget region"] = _derive_sales_budget_region(work).astype(str)
    out["Cust Region"]         = _series_or_empty(work, "Cust Region").astype(str)
    status_col = "Updated Status" if "Updated Status" in work.columns else "Customer Status"
    out["Customer Status"]     = _series_or_empty(work, status_col).astype(str)
    out["Main Ac"]             = _series_or_empty(work, "Main Ac").astype(str).str.strip()
    out["Focus List"]          = ""

    # ---------------- Insurance ----------------
    out["Insurance"] = ""
    if ins_df is not None and not ins_df.empty:
        tmp = out[["CustCode", "Main Ac"]].copy()
        tmp["__CustCode"] = tmp["CustCode"].astype(str).str.strip()
        tmp["__MainAc"]   = tmp["Main Ac"].astype(str).str.strip()
        master = ins_df.copy()
        master["__CustCode"] = master["Customer Code"].astype(str).str.strip()
        master["__MainAc"]   = master["Main Account"].astype(str).str.strip()
        merged = tmp.merge(
            master[["__CustCode", "__MainAc", "Insurance Limit"]],
            how="left",
            left_on=["__CustCode", "__MainAc"],
            right_on=["__CustCode", "__MainAc"]
        )
        ins_val = pd.to_numeric(merged["Insurance Limit"], errors="coerce")
        out["Insurance"] = ins_val.where(ins_val.notna(), "").astype(object)

    # ---------------- AR / Aging Columns ----------------
    on_acc_src   = _first_present(work, ["On Account (Derived)", "On account"])
    not_due_src  = _first_present(work, ["Not Due Amount", "Not Due (Derived)", "Not Due"])
    a1_30_src    = _first_present(work, ["Aging 1 to 30"])
    a31_60_src   = _first_present(work, ["Aging 31 to 60"])
    a61_90_src   = _first_present(work, ["Aging 61 to 90"])
    a91_120_src  = _first_present(work, ["Aging 91 to 120"])
    a121_150_src = _first_present(work, ["Aging 121 to 150"])
    a_ge_151_src = _first_present(work, ["Aging >=151", "Aging ≥151 (Amount)"])

    on_acc   = _num(work, on_acc_src)
    not_due  = _num(work, not_due_src)
    a1_30    = _num(work, a1_30_src)
    a31_60   = _num(work, a31_60_src)
    a61_90   = _num(work, a61_90_src)
    a91_120  = _num(work, a91_120_src)
    a121_150 = _num(work, a121_150_src)
    a_ge_151 = _num(work, a_ge_151_src)

    # Compute Aging 1–60 as 1–30 + 31–60
    a1_60 = a1_30 + a31_60

    # AR Balance: use existing if available else sum components
    ar_balance_src = _first_present(work, ["AR Balance", "Ar Balance (Copy)"])
    if ar_balance_src:
        ar_bal = _num(work, ar_balance_src)
    else:
        ar_bal = on_acc + not_due + a1_60 + a61_90 + a91_120 + a121_150 + a_ge_151

    # ---------------- Map to BUD headers ----------------
    out["On\nAccount"]        = on_acc
    out["Not Due\nAmount"]    = not_due
    out["Aging\n1 to 60"]     = a1_60
    out["Aging\n61 to 90"]    = a61_90
    out["Aging\n91 to 120"]   = a91_120
    out["Aging\n121 to 150"]  = a121_150
    out["Aging\n>=151"]       = a_ge_151
    out[" AR\nBalance"]       = ar_bal

    return out