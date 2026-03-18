# ui_old_tool.py — FINAL VERSION (NO SUMIFS, NO FORMULAS for Q1/tails)

import io
import time
import pandas as pd
import streamlit as st

# ⬇️ USE THE NEW REGENERATED PYTHON FILES
from processor import process_ar_file, customer_summary, invoice_summary
from budg.updated_export import fast_excel_download_multiple_with_formulas


def render_old_tool():
    st.header("AR Backlog → By_Customer Forecast Tool")

    uploaded_file = st.file_uploader(
        "Upload AR Backlog Excel",
        type=["xlsx", "xls"],
        key="old_uploader",
        help="Upload the AR Backlog workbook (As on Date in B14, header row = 16)."
    )

    if not uploaded_file:
        return

    try:
        total_start = time.perf_counter()

        # ---------------- PROCESSING ----------------
        process_start = time.perf_counter()
        with st.spinner("Processing file..."):

            # df_main = AR_Backlog + derived values
            df_main = process_ar_file(uploaded_file)

            # df_customer = GROUPED FORECAST (Q1 & 4 tails ARE PYTHON)
            df_customer = customer_summary(df_main)

            # df_invoice = raw invoice sheet
            df_invoice = invoice_summary(df_main)

        process_end = time.perf_counter()
        st.success("Processing completed.")

        # ---------------- EXPORT ----------------
        st.subheader("Download")

        excel_file = fast_excel_download_multiple_with_formulas(
            df_main,
            df_customer,
            df_invoice
        )

        st.download_button(
            "Download Processed File",
            data=excel_file.getvalue(),
            file_name="processed_AR_backlog.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="old_download_btn"
        )

        # ---------------- METRICS ----------------
        st.subheader("Performance Metrics")
        c1, c2, c3 = st.columns(3)
        c1.metric("Processing Time", f"{process_end - process_start:.2f} sec")
        export_end = time.perf_counter()
        c2.metric("Export Time", f"{export_end - process_end:.2f} sec")
        total_end = time.perf_counter()
        c3.metric("Total Runtime", f"{total_end - total_start:.2f} sec")

    except Exception as e:
        st.error("An error occurred.")
        st.exception(e)