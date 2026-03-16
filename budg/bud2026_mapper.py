# budg/bud2026_mapper.py
import pandas as pd
import numpy as np

try:
    from region_maps import classify_region
except Exception:
    classify_region = None


def _series_or_empty(df: pd.DataFrame, col: str) -> pd.Series:
    return df[col] if col in df.columns else pd.Series([""] * len(df), index=df.index)


def _num(df: pd.DataFrame, col: str) -> pd.Series:
    """Coerce to numeric; missing -> 0.0"""
    if not col or col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype="float64")
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)


def _derive_sales_budget_region(df_cust: pd.DataFrame) -> pd.Series:
    if "Region" in df_cust.columns:
        reg = df_cust["Region"].fillna("").astype(str)
        if reg.str.strip().any():
            return reg
    if classify_region is not None and "Cust Region" in df_cust.columns:
        cust_code = df_cust["Cust Code"] if "Cust Code" in df_cust.columns else None
        derived = classify_region(df_cust["Cust Region"], cust_code)
        return derived.fillna("")
    return pd.Series([""] * len(df_cust), index=df_cust.index)


def _first_present(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first column name from candidates that exists in df.columns."""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def map_by_customer_to_bud2026(df_customer: pd.DataFrame, ins_df: pd.DataFrame = None) -> pd.DataFrame:
    """
    Build a DataFrame for BUD2026 with:
      - Identifiers (CustCode, Cust Name, BT, Sales Budget region, Cust Region, Customer Status, Main Ac, Focus List)
      - Insurance (from master if provided)
      - AR composition (On/Not Due/Aging buckets/AR Balance)
        with Aging 1–60 computed as (1–30 + 31–60).
    """
    work = df_customer.copy()

    # ---------------- Identifiers ----------------
    out = pd.DataFrame(index=work.index)
    out["CustCode"]             = _series_or_empty(work, "Cust Code").astype(str).str.strip()
    out["Cust Name"]            = _series_or_empty(work, "Cust Name").astype(str)
    out["BT"]                   = ""
    out["Sales Budget region"]  = _derive_sales_budget_region(work).astype(str)
    out["Cust Region"]          = _series_or_empty(work, "Cust Region").astype(str)

    status_col = "Updated Status" if "Updated Status" in work.columns else "Customer Status"
    out["Customer Status"]      = _series_or_empty(work, status_col).astype(str)
    out["Main Ac"]              = _series_or_empty(work, "Main Ac").astype(str).str.strip()
    out["Focus List"]           = ""

    # ---------------- Insurance (optional master merge) ----------------
    out["Insurance"]            = ""
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
            right_on=["__CustCode", "__MainAc"],
        )
        ins_val = pd.to_numeric(merged["Insurance Limit"], errors="coerce")
        out["Insurance"] = ins_val.where(ins_val.notna(), "").astype(object)

    # ---------------- AR composition (robust to label variants) ----------------
    # Accept both ASCII >= and Unicode ≥ in sources; accept derived/amount variants
    on_acc_src   = _first_present(work, ["On Account (Derived)", "On Account"])
    not_due_src  = _first_present(work, ["Not Due Amount", "Not Due (Derived)", "Not Due"])
    a1_30_src    = _first_present(work, ["Aging 1 to 30 (Amount)"])
    a31_60_src   = _first_present(work, ["Aging 31 to 60 (Amount)"])
    a61_90_src   = _first_present(work, ["Aging 61 to 90 (Amount)"])
    a91_120_src  = _first_present(work, ["Aging 91 to 120 (Amount)"])
    a121_150_src = _first_present(work, ["Aging 121 to 150 (Amount)"])
    a_ge_151_src = _first_present(work, ["Aging >=151 (Amount)", "Aging ≥151 (Amount)"])  # handle both

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

    # AR Balance from column if present; else recompute from parts
    ar_balance_src = _first_present(work, ["AR Balance", "Ar Balance (Copy)"])
    if ar_balance_src:
        ar_bal = _num(work, ar_balance_src)
    else:
        ar_bal = on_acc + not_due + a1_60 + a61_90 + a91_120 + a121_150 + a_ge_151

    # ---------------- Map to BUD headers (exact strings with spaces + \n) ----------------
    # NOTE: Ensure these match your HEADERS_BUD2026 1:1 (spaces and line breaks).
    # Map to EXACT BUD headers (match HEADERS_BUD2026)
    out["On\nAccount"]        = on_acc
    out["Not Due\nAmount"]    = not_due
    out["Aging\n1 to 60"]     = a1_60
    out["Aging\n61 to 90"]    = a61_90
    out["Aging\n91 to 120"]   = a91_120
    out["Aging\n121 to 150"]  = a121_150
    out["Aging\n>=151"]       = a_ge_151    # ASCII >=
    out[" AR\nBalance"]       = ar_bal  

    return out