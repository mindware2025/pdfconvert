import io
import pandas as pd
import xlsxwriter


# ----------------- Helpers -----------------

def num_to_col_letters(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_col_map(df) -> dict:
    return {col: num_to_col_letters(i) for i, col in enumerate(df.columns)}


def normalize_all_date_strings(df: pd.DataFrame) -> pd.DataFrame:
    import re
    if df is None or df.empty:
        return df
    df = df.copy()
    iso = re.compile(r"^\s*\d{4}-\d{2}-\d{2}")
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            s = df[col].astype(str)
            if not s.head(50).apply(lambda x: bool(iso.match(x)) if x and x != "nan" else False).any():
                continue
            s = (
                s.replace("\u00A0", " ", regex=False)
                 .str.replace(r"[^\x00-\x7F]", "", regex=True)
                 .str.strip()
            )
            s = s.where(~s.str.match(r"^\d{4}-\d{2}-\d{2}"), s.str.slice(0, 10))
            df[col] = s
    return df


# ----------------- Exporter -----------------

def fast_excel_download_multiple_with_formulas(df_main, df_customer, df_invoice) -> io.BytesIO:
    """
    FINAL EXPORTER — NO SUMIFS.
    - Q1-2026 and all 4 tail buckets are Python-calculated.
    - Only 4 formulas are written:
        Actual Q1
        Remaining % from q1
        To add to Q2
        Forecast Q2
    """

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2"})

    # ---------------- Sheet 1: AR_Backlog ----------------
    ws_main = wb.add_worksheet("AR_Backlog")
    main_df = normalize_all_date_strings(df_main.copy()).fillna("")
    for c_idx, name in enumerate(main_df.columns):
        ws_main.write(0, c_idx, str(name), header_fmt)
    for r_idx, row in enumerate(main_df.to_numpy(), start=1):
        ws_main.write_row(r_idx, 0, row.tolist())

    # ---------------- Sheet 2: By_Customer ----------------
    ws_cust = wb.add_worksheet("By_Customer")
    cust_df = normalize_all_date_strings(df_customer.copy()).fillna("")
    if "Q1-2026" in cust_df.columns and "Q1-2026 - pivot" in cust_df.columns:
          cust_df["Q1-2026"] = cust_df["Q1-2026 - pivot"]
    # Headers
    for c_idx, name in enumerate(cust_df.columns):
        ws_cust.write(0, c_idx, str(name), header_fmt)

    col_map = build_col_map(cust_df)

    def idx(name):
        return list(cust_df.columns).index(name)

    # Write rows + ONLY the % block formulas
    for r_idx, row in enumerate(cust_df.to_numpy(), start=1):
        excel_row = r_idx + 1
        ws_cust.write_row(r_idx, 0, row.tolist())

        q1 = col_map.get("Q1-2026")
        pct = col_map.get("% for Q1")
        rem = col_map.get("Remaining % from q1")
        add = col_map.get("To add to Q2")
        q2 = col_map.get("Q2-2026")

        # Actual Q1 = % for Q1 * Q1-2026
        if q1 and pct:
            ws_cust.write_formula(r_idx, idx("Actual Q1"),
                                  f"=IFERROR(${q1}{excel_row}*${pct}{excel_row},0)")

        # Remaining % = 1 - % for Q1
        if pct:
            ws_cust.write_formula(r_idx, idx("Remaining % from q1"),
                                  f"=IFERROR(1-${pct}{excel_row},0)")

        # To add to Q2 = Remaining % * Q1-2026
        if rem and q1:
            ws_cust.write_formula(r_idx, idx("To add to Q2"),
                                  f"=IFERROR(${rem}{excel_row}*${q1}{excel_row},0)")

        # Forecast Q2 = Q2-2026 + To add to Q2
        if add and q2:
            ws_cust.write_formula(r_idx, idx("Forecast Q2"),
                                  f"=IFERROR(${q2}{excel_row}+${add}{excel_row},0)")

    ws_cust.freeze_panes(1, 0)
    ws_cust.autofilter(0, 0, max(1, len(cust_df)), len(cust_df.columns)-1)

    # ---------------- Sheet 3: Invoice ----------------
    ws_inv = wb.add_worksheet("Invoice")
    inv_df = normalize_all_date_strings(df_invoice.copy()).fillna("")
    for c_idx, name in enumerate(inv_df.columns):
        ws_inv.write(0, c_idx, str(name), header_fmt)
    for r_idx, row in enumerate(inv_df.to_numpy(), start=1):
        ws_inv.write_row(r_idx, 0, row.tolist())

    wb.close()
    output.seek(0)
    return output