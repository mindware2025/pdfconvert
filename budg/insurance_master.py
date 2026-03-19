# budg/insurance_master.py
import pandas as pd

_REQUIRED_COLS = {
    "Customer Code",
    "Customer name",
    "Main Account",
    "Buyer Insurance Code",
    "Region",
    "Region Name",
    "Euler ID",
    "Insurance Ref No",
    "Insurance Limit",
    "Currency",
    "Effective From",
    "Effective To",
    "Notification Status",
    "Created By",
    "Created Date",
    "Default Pay Term",
}


def _sanitize_cols(cols):
    """Trim, replace NBSP with space; keep original case."""
    fixed = []
    for c in cols:
        if c is None:
            fixed.append("")
            continue
        s = str(c).replace("\u00A0", " ").strip()
        fixed.append(s)
    return fixed


def load_insurance_master(xlsx_or_filelike) -> pd.DataFrame:
    """
    Read the Insurance Master with header at Excel row 8 (header=7), data from row 9.
    Returns a deduplicated DataFrame keyed by (Customer Code, Main Account),
    keeping only the most recent (Effective From, then Created Date).
    """
    df = pd.read_excel(
        xlsx_or_filelike,
        header=7,              # header at row 8 (0-based index = 7)
        dtype=str,
        engine="openpyxl"
    )
    # Clean column names
    df.columns = _sanitize_cols(df.columns)

    # Keep only required columns (if present)
    present = [c for c in _REQUIRED_COLS if c in df.columns]
    df = df[present].copy()

    # Strip whitespace on key fields
    df["Customer Code"] = df["Customer Code"].fillna("").astype(str).str.strip()
    df["Main Account"]  = df["Main Account"].fillna("").astype(str).str.strip()

    # Parse dates for sorting/tie-break (coerce errors -> NaT)
    for dc in ["Effective From", "Effective To", "Created Date"]:
        if dc in df.columns:
            df[dc] = pd.to_datetime(df[dc], errors="coerce", dayfirst=True)

    # Sort by latest Effective From, then latest Created Date
    sort_cols = []
    if "Effective From" in df.columns:
        sort_cols.append("Effective From")
    if "Created Date" in df.columns:
        sort_cols.append("Created Date")
    if sort_cols:
        df = df.sort_values(sort_cols, ascending=[False] * len(sort_cols))

    # Deduplicate: keep the first after sorting
    df = df.drop_duplicates(subset=["Customer Code", "Main Account"], keep="first")

    # Keep a small lookup footprint
    out_cols = ["Customer Code", "Main Account", "Insurance Limit"]
    if "Effective From" in df.columns:
        out_cols.append("Effective From")
    if "Effective To" in df.columns:
        out_cols.append("Effective To")

    # Normalize Insurance Limit (numeric if possible)
    df["Insurance Limit"] = pd.to_numeric(df["Insurance Limit"], errors="coerce")

    return df[out_cols].reset_index(drop=True)
