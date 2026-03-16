
import io
import time
import pandas as pd
import streamlit as st
import xlsxwriter

from processor import process_ar_file, customer_summary, invoice_summary

# --- helpers (moved verbatim) ---
def num_to_col_letters(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def build_col_map(df) -> dict:
    return {col: num_to_col_letters(i) for i, col in enumerate(df.columns)}

def formula_for_main_cell(col_name, row_idx_1based, col_map):
    C = col_map
    r = str(row_idx_1based)
    def REF(name): return f"${C[name]}{r}"

    has_copy = "Ar Balance (Copy)" in C
    has_overdue = "Overdue days (Days)" in C
    has_bp = "Invoice Value (Derived)" in C

    if col_name == "Invoice Value (Derived)":
        return f"=IF({REF('Ar Balance (Copy)')}>0,{REF('Ar Balance (Copy)')},0)" if has_copy else None
    if col_name == "On Account (Derived)":
        return f"=IF({REF('Ar Balance (Copy)')}<0,{REF('Ar Balance (Copy)')},0)" if has_copy else None
    if col_name == "Not Due (Derived)":
        return f"=IF({REF('Overdue days (Days)')}>0,0,{REF('Invoice Value (Derived)')})" if (has_overdue and has_bp) else None

    if has_overdue and has_bp:
        if col_name == "Aging >=151 (Amount)":
            return f"=IF({REF('Overdue days (Days)')}>150,{REF('Invoice Value (Derived)')},0)"
        if col_name == "Aging 121 to 150 (Amount)":
            return f"=IF({REF('Overdue days (Days)')}>120,{REF('Invoice Value (Derived)')},0)-{REF('Aging >=151 (Amount)')}"
        if col_name == "Aging 91 to 120 (Amount)":
            return (f"=IF({REF('Overdue days (Days)')}>90,{REF('Invoice Value (Derived)')},0)"
                    f"-{REF('Aging 121 to 150 (Amount)')}-{REF('Aging >=151 (Amount)')}")
        if col_name == "Aging 61 to 90 (Amount)":
            return (f"=IF({REF('Overdue days (Days)')}>60,{REF('Invoice Value (Derived)')},0)"
                    f"-{REF('Aging 91 to 120 (Amount)')}-{REF('Aging 121 to 150 (Amount)')}-{REF('Aging >=151 (Amount)')}")
        if col_name == "Aging 31 to 60 (Amount)":
            return (f"=IF({REF('Overdue days (Days)')}>30,{REF('Invoice Value (Derived)')},0)"
                    f"-{REF('Aging 61 to 90 (Amount)')}-{REF('Aging 91 to 120 (Amount)')}"
                    f"-{REF('Aging 121 to 150 (Amount)')}-{REF('Aging >=151 (Amount)')}")
        if col_name == "Aging 1 to 30 (Amount)":
            return (f"=IF({REF('Overdue days (Days)') }>=0,{REF('Invoice Value (Derived)')},0)"
                    f"-{REF('Aging 31 to 60 (Amount)')}-{REF('Aging 61 to 90 (Amount)')}"
                    f"-{REF('Aging 91 to 120 (Amount)')}-{REF('Aging 121 to 150 (Amount)')}"
                    f"-{REF('Aging >=151 (Amount)')}")
    if col_name == "Ageing > 365 (Amt)":
        if "Ageing (Days)" in C and "Ar Balance (Copy)" in C:
            return f"=IF({REF('Ageing (Days)')}>365,{REF('Ar Balance (Copy)')},0)"
    return None

def normalize_all_date_strings(df: pd.DataFrame) -> pd.DataFrame:
    import re
    if df is None or df.empty: return df
    df = df.copy()
    iso = re.compile(r"^\s*\d{4}-\d{2}-\d{2}")
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            s = df[col].astype(str)
            if not s.head(50).apply(lambda x: bool(iso.match(x)) if x and x != "nan" else False).any():
                continue
            s = (s.replace("\u00A0"," ",regex=False)
                   .str.replace(r"[^\x00-\x7F]","",regex=True)
                   .str.strip())
            s = s.where(~s.str.match(r"^\d{4}-\d{2}-\d{2}"), s.str.slice(0,10))
            df[col] = s
    return df

def fast_excel_download_multiple_with_formulas(df_main, df_customer, df_invoice, export_with_formulas=True) -> io.BytesIO:
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"constant_memory": True})
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#F2F2F2"})

    # AR_Backlog
    ws_main = workbook.add_worksheet("AR_Backlog")
    main_df = normalize_all_date_strings(df_main.copy())
    for col in main_df.columns:
        if pd.api.types.is_datetime64_any_dtype(main_df[col]):
            main_df[col] = main_df[col].dt.strftime("%Y-%m-%d")
    main_df = main_df.fillna("")
    main_map = build_col_map(main_df)

    for c_idx, name in enumerate(main_df.columns):
        ws_main.write(0, c_idx, str(name), header_fmt)

    for r_idx, row in enumerate(main_df.to_numpy(), start=1):
        row_num_1b = r_idx + 1
        for c_idx, val in enumerate(row):
            col_name = main_df.columns[c_idx]
            if export_with_formulas:
                f = formula_for_main_cell(col_name, row_num_1b, main_map)
                if f:
                    ws_main.write_formula(r_idx, c_idx, f)
                    continue
            ws_main.write(r_idx, c_idx, val)

    # By_Customer
    ws_cust = workbook.add_worksheet("By_Customer")
    cust_df = normalize_all_date_strings(df_customer.copy()).fillna("")
    for col in cust_df.columns:
        if pd.api.types.is_datetime64_any_dtype(cust_df[col]):
            cust_df[col] = cust_df[col].dt.strftime("%Y-%m-%d")
    for c_idx, name in enumerate(cust_df.columns):
        ws_cust.write(0, c_idx, str(name), header_fmt)
    for r_idx, row in enumerate(cust_df.to_numpy(), start=1):
        ws_cust.write_row(r_idx, 0, row.tolist())

    # Invoice
    ws_inv = workbook.add_worksheet("Invoice")
    inv_df = normalize_all_date_strings(df_invoice.copy()).fillna("")
    for col in inv_df.columns:
        if pd.api.types.is_datetime64_any_dtype(inv_df[col]):
            inv_df[col] = inv_df[col].dt.strftime("%Y-%m-%d")
    for c_idx, name in enumerate(inv_df.columns):
        ws_inv.write(0, c_idx, str(name), header_fmt)
    for r_idx, row in enumerate(inv_df.to_numpy(), start=1):
        ws_inv.write_row(r_idx, 0, row.tolist())

    workbook.close()
    output.seek(0)
    return output

def render_old_tool():
    uploaded_file = st.file_uploader(
        "Upload AR Backlog Excel", type=["xlsx", "xls"], key="old_uploader",
        help="Full AR Backlog workbook."
    )
    if not uploaded_file:
        return
    try:
        total_start = time.perf_counter()
        process_start = time.perf_counter()
        with st.spinner("Processing file..."):
            df_main = process_ar_file(uploaded_file)
            df_customer = customer_summary(df_main)
            df_invoice  = invoice_summary(df_main)
        process_end = time.perf_counter()
        st.success("Processing completed")

        st.subheader("Export Options")
        export_with_formulas = st.checkbox("Export formulas on AR_Backlog", value=True, key="old_export_formulas")

        export_start = time.perf_counter()
        excel_file = fast_excel_download_multiple_with_formulas(
            df_main, df_customer, df_invoice, export_with_formulas=export_with_formulas
        )
        export_end = time.perf_counter()

        st.download_button(
            "Download Processed File",
            data=excel_file.getvalue(),
            file_name="processed_AR_backlog.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="old_download_btn",
        )

        st.subheader("Performance Metrics")
        c1, c2, c3 = st.columns(3)
        c1.metric("Processing Time", f"{(process_end - process_start):.2f} sec")
        c2.metric("Export Time", f"{(export_end - export_start):.2f} sec")
        total_end = time.perf_counter()
        c3.metric("Total Runtime", f"{(total_end - total_start):.2f} sec")

    except Exception as e:
        st.error(f"{e}\n\nIf this persists, expand 'Details' for traceback and share the top 10 lines.")
        st.exception(e)
