import pandas as pd

# ----------------------------- BUSINESS SETS -----------------------------
# QNAL: as per your lists (including CHAD, GHANA, MAURITIUS, EGYPT, PALESTINE, TANZANIA, TOGO, LATVIA)
QNAL = {
    "AFRICA", "ALGERIA", "BURKINA FASO", "CAMEROON", "CHAD", "CONGO",
    "EGYPT", "FRANCE", "GABON", "GHANA", "GUINEA", "IVORY COAST",
    "IRAQ", "JORDAN", "LEBANON", "LIBYA", "MALI", "MAURITANIA",
    "MAURITIUS", "MOROCCO", "NIGERIA", "PALESTINE", "QATAR",
    "SENEGAL", "SOUTH AFRICA", "SUDAN", "TANZANIA", "TOGO",
    "TUNISIA", "LATVIA",
}

# GCC: as per your lists (broad business group, not strictly political GCC)
GCC = {
    "AFGHANISTAN", "AUSTRALIA", "AUSTRIA", "BAHRAIN", "ETHIOPIA", "GERMANY",
    "GREAT BRITAIN (UNITED KINGDOM)", "GREECE", "HONG KONG", "INDIA",
    "IRELAND", "ITALY", "KENYA", "KUWAIT", "LUXEMBOURG", "NETHERLANDS",
    "NEW ZEALAND", "OMAN", "PAKISTAN", "POLAND", "SINGAPORE", "SWITZERLAND",
    "UAE", "UGANDA", "UNITED STATES OF AMERICA", "YEMEN",
}

# Cities/country tokens that should map to KSA directly
KSA_TOKENS = {"SAUDI", "RIYADH", "JEDDAH"}

# Optional normalization map for common variants/misspellings (extend as needed)
NORMALIZE_TOKENS = {
    "SAUDI ARABIA": "SAUDI",
    "KSA": "SAUDI",
    "U.A.E.": "UAE",
    "UNITED KINGDOM": "GREAT BRITAIN (UNITED KINGDOM)",
    "UK": "GREAT BRITAIN (UNITED KINGDOM)",
    "COTE DIVOIRE": "IVORY COAST",
    "CÔTE D’IVOIRE": "IVORY COAST",
    "COTE D'IVOIRE": "IVORY COAST",
}

# ----------------------------- HELPERS -----------------------------
def _normalize_str_series(s: pd.Series) -> pd.Series:
    """
    Uppercase, strip, remove NBSP and hidden non-ASCII.
    Apply token normalization for known variants.
    """
    if s is None:
        return pd.Series([], dtype="object")
    ser = (
        s.fillna("")
         .astype(str)
         .str.replace("\u00A0", " ", regex=False)    # NBSP -> space
         .str.replace(r"[^\x00-\x7F]", "", regex=True)  # remove hidden noise
         .str.strip()
         .str.upper()
    )
    # Apply normalization dict (vectorized)
    if NORMALIZE_TOKENS:
        norm_map_upper = {k.upper(): v.upper() for k, v in NORMALIZE_TOKENS.items()}
        ser = ser.replace(norm_map_upper)
    return ser

# ----------------------------- MAIN -----------------------------
def classify_region(series: pd.Series, customer_code_series: pd.Series = None) -> pd.Series:
    """
    Region rule:
      1) If Country in {SAUDI, RIYADH, JEDDAH} -> 'KSA'
      2) Else if Country in QNAL -> 'qnal'
      3) Else if Country in GCC  -> 'gcc'
      4) CK override: if Customer Code starts 'CK' -> 'KSA' (hard override)
      5) Else -> '' (blank)
    """
    region_src = _normalize_str_series(series)
    result = pd.Series("", index=region_src.index, dtype="object")

    # 1) KSA tokens first
    mask_ksa = region_src.isin(KSA_TOKENS)
    result[mask_ksa] = "KSA"

    # 2) QNAL mapping (where not yet decided)
    mask_qnal = region_src.isin(QNAL) & (result == "")
    result[mask_qnal] = "qnal"

    # 3) GCC mapping (where not yet decided)
    mask_gcc = region_src.isin(GCC) & (result == "")
    result[mask_gcc] = "gcc"

    # 4) CK override (wins over anything)
    if customer_code_series is not None:
        customer_code = _normalize_str_series(customer_code_series)
        result[customer_code.str.startswith("CK")] = "KSA"

    # 5) Remaining stay blank
    return result
