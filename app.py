import streamlit as st
import pandas as pd
import pdfplumber
import re
import tempfile
import os
import io
from datetime import datetime

from extractors.google_dnts import extract_invoice_info, extract_table_from_text, make_dnts_header_row, DNTS_HEADER_COLS, DNTS_ITEM_COLS
from utils.helpers import format_amount, format_invoice_date, format_month_year
from dotenv import load_dotenv
load_dotenv()
from extractors.google_invoice import extract_table_from_text as extract_invoice_table, extract_invoice_info as extract_invoice_info_invoice, GOOGLE_INVOICE_COLS
from extractors.cloud_invoice import process_cloud_invoice
from claims_automation import (
    build_output_rows_from_source1,
    write_output_excel,
    read_source1_rows,
    read_master1_map,
    read_source2_rows,
    build_debit_rows_from_source2,
    read_master2_entries,
    derive_defaults_from_source1,
)


st.set_page_config(
    page_title="Google DNTS upload file",
    layout="wide"
)

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Google Sans', sans-serif !important;
        background: #f6f8fa;
        color: #202124;
    }
    h1, h2, h3 {
        color: #1a73e8;
        font-weight: 700;
        letter-spacing: -1px;
        margin-bottom: 0.5em;
    }
    .stButton > button, .stDownloadButton > button {
        background: linear-gradient(90deg, #1a73e8, #188038);
        color: white;
        border-radius: 8px;
        font-weight: 600;
        font-size: 16px;
        border: none;
        padding: 12px 28px;
        margin-top: 10px;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(26, 115, 232, 0.08);
        transition: background 0.2s;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: linear-gradient(90deg, #188038, #1a73e8);
        color: #fff;
    }
    .stDataFrame {
        background: #fff;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(26, 115, 232, 0.06);
        margin-bottom: 2em;
    }
    </style>
""", unsafe_allow_html=True)

# ----------- Constants -----------
DEFAULTS = {
    "supp_code": "SDIG005",
    "curr_code": "USD",
    "form_code": 0,
    "doc_src_locn": "UJ000",
    "location_code": "UJ200"
}
#CORRECT_USERNAME = "admin"
#CORRECT_PASSWORD = "admin"

CORRECT_USERNAME = os.getenv("NAME")
CORRECT_PASSWORD = os.getenv("PASSWORD")


if "login_state" not in st.session_state:
    st.session_state.login_state = "login" 

def show_login():
    
    for _ in range(10):
        st.write("")
   
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.title("🔒 Login to PDF to Excel")
        username = st.text_input("Username", key="login_user")
        password = st.text_input("Password", type="password", key="login_pass")
        if st.button("Login", key="login_btn"):
            if username == CORRECT_USERNAME and password == CORRECT_PASSWORD:
                st.session_state.login_state = "success"
            else:
                st.session_state.login_state = "fail"

def show_fail():
    st.image("data:image/png;base64,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", width=350)
    st.error("Oops! Wrong credentials... Nice try, but no entry! 😜")
    if st.button("Back to Login", key="back_login"):
        st.session_state.login_state = "login"

if st.session_state.login_state == "login":
    show_login()
    st.stop()
elif st.session_state.login_state == "fail":
    show_fail()
    st.stop()

def extractor_workflow(
    extractor_name,
    extractor_info,
    file_uploader_label,
    extract_invoice_info_func,
    extract_table_func,
    table_columns,
    file_name_template,
    show_header_df_func=None,
    header_columns=None,
    item_row_builder=None,
    item_columns=None
):
    st.title(f"PDF TO EXCEL ({extractor_name})")
    st.write(extractor_info)
    uploaded_file = st.file_uploader(file_uploader_label, type=["pdf"], accept_multiple_files=False, key=f"uploader_{extractor_name}")
    if uploaded_file:
        invoice_num, invoice_date = extract_invoice_info_func(uploaded_file)
        if invoice_num and invoice_date:
            file_date = format_month_year(invoice_date)
            file_name = file_name_template.format(invoice_num=invoice_num, file_date=file_date)
        else:
            file_name = file_name_template.format(invoice_num='unknown', file_date='unknown')
        rows = extract_table_func(uploaded_file)
        if rows:
            df = pd.DataFrame(rows, columns=table_columns)
            if show_header_df_func and header_columns and item_row_builder and item_columns:
                today_str = datetime.today().strftime("%d/%m/%Y")
                remarks = f"GOOGLE INV-{invoice_num}" if invoice_num else "GOOGLE INV-UNKNOWN"
                header_df = pd.DataFrame([
                    show_header_df_func(invoice_num, invoice_date, today_str, remarks)
                ], columns=header_columns)
                st.subheader("DNTS Header Preview")
                st.dataframe(header_df, height=120)
                dnts_item_data = [item_row_builder(idx, *row, invoice_num) for idx, row in enumerate(rows, 1)]
                dnts_item_df = pd.DataFrame(dnts_item_data, columns=item_columns)
                st.subheader("DNTS Items Preview")
                st.dataframe(dnts_item_df)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    header_df.to_excel(writer, sheet_name='DNTS_HEADER', index=False)
                    dnts_item_df.to_excel(writer, sheet_name='DNTS_ITEM', index=False)
                output.seek(0)
                st.download_button(
                    label=f"⬇️ Download DNTS Excel",
                    data=output.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{extractor_name}"
                )
            else:
                st.subheader("Extracted Table")
                st.dataframe(df, height=300)
                towrite = io.BytesIO()
                df.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)
                st.download_button(
                    label="Download as Excel",
                    data=towrite,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("No table data found in the uploaded PDF.")
    else:
        st.info(f"Please upload a {extractor_name} PDF file to get started.")

# ----------- Tool Selector UI -----------
st.markdown("""
    <div style='text-align:center; margin-top:2rem; margin-bottom:1.5rem;'>
        <h2 style='color:#1a73e8; font-family:Google Sans, sans-serif; font-weight:700; letter-spacing:-1px;'>🛠️ Tool Selection</h2>
        <p style='font-size:1.2rem; color:#444;'>Choose the tool you want to use for your PDF extraction.</p>
    </div>
""", unsafe_allow_html=True)

TOOL_OPTIONS = [
    "-- Select a tool --",
    "🟦 Google DNTS Extractor",
    "🟩 Google Invoice Extractor",
    "📄 Claims Automation",
    "🧾 Cloud Invoice Tool",
    "💻 Dell Invoice Extractor", 
    "Other"
]
tool = st.selectbox(
    "Select a tool:",
    TOOL_OPTIONS,
    key="tool_selector"
)

if tool == "-- Select a tool --":
    st.info("Please select a tool above to get started.")
elif tool == "🟦 Google DNTS Extractor":
    def dnts_item_row(idx, domain, customer_id, amount, invoice_num):
        formatted_amount = format_amount(amount)
        item_name = (
            f"GOOGLE INV-{invoice_num} / DOMAIN NAME : {domain} / CUSTOMER ID : {customer_id} / AMOUNT - USD - {formatted_amount}"
        ).upper()
        return [
            idx, 1, "NS", item_name, "NA", "NA", "NOS", 1, 0, formatted_amount, 14401, "SDIG005", "PUHO", "GEN", "ZZ-COMM"
        ]
    extractor_workflow(
        extractor_name="Google DNTS Extractor",
        extractor_info="Upload one PDF containing a **'Summary of costs by domain'** table. The app will extract the table and let you download it as Excel.",
        file_uploader_label="Choose your Google DNTS Invoice PDF",
        extract_invoice_info_func=extract_invoice_info,
        extract_table_func=extract_table_from_text,
        table_columns=DNTS_ITEM_COLS[:3],
        file_name_template="{invoice_num}-{file_date}.xlsx",
        show_header_df_func=make_dnts_header_row,
        header_columns=DNTS_HEADER_COLS,
        item_row_builder=dnts_item_row,
        item_columns=DNTS_ITEM_COLS
    )
elif tool == "🟩 Google Invoice Extractor":
    extractor_workflow(
        extractor_name="Google Invoice Extractor",
        extractor_info="Upload a Google Invoice PDF. The app will extract the relevant data and let you download it as Excel.",
        file_uploader_label="Choose your Google Invoice PDF",
        extract_invoice_info_func=extract_invoice_info_invoice,
        extract_table_func=extract_invoice_table,
        table_columns=GOOGLE_INVOICE_COLS,
        file_name_template="{invoice_num}-{file_date}.xlsx"
    )
elif tool == "📄 Claims Automation":
    st.title("Claims Automation")
    
    st.header("📁 Upload Files")
    source1_file = st.file_uploader("JV Orion from SAP (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_source1")
    master1_file = st.file_uploader("User information (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_master1")
    source2_file = st.file_uploader("Employee benefits (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_source2")
    master2_file = st.file_uploader("Main acc file (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_master2")
    
    st.markdown("---")
    run_clicked = st.button("🚀 Generate Output", key="claims_run", use_container_width=True)

    if run_clicked:
        if not source1_file:
            st.error("Please upload Source File 1.")
            st.stop()
        try:
            src_rows = read_source1_rows(source1_file)
            master_map = read_master1_map(master1_file) if master1_file else None
            src2_rows = read_source2_rows(source2_file) if source2_file else None
            master2_entries = read_master2_entries(master2_file) if master2_file else None

            credit_rows = build_output_rows_from_source1(
                src_rows,
                master1_map=master_map,
                source2_rows=src2_rows,
                user_id_col="Sub Acct",
            )

            src1_defaults = derive_defaults_from_source1(src_rows)

            doc_ref = ""
            for r in src_rows:
                val = r.get("Doc Ref.", "")
                if val and str(val).strip():
                    doc_ref = str(val).strip()
                    break

            debit_rows = build_debit_rows_from_source2(
                src2_rows,
                master2_entries=master2_entries,
                master1_map=master_map,
                default_div=src1_defaults.get("Div", ""),
                default_dept=src1_defaults.get("Dept", ""),
                default_anly1=src1_defaults.get("Anly1", ""),
                default_anly2=src1_defaults.get("Anly2", ""),
                default_currency=src_rows[0].get("Currency", "") if src_rows else "",
                doc_ref=doc_ref,
            )

            out_rows = credit_rows + debit_rows

            output_buffer = io.BytesIO()
            write_output_excel(output_buffer, out_rows)
            output_buffer.seek(0)

            st.download_button(
                label="Download claims_output.xlsx",
                data=output_buffer,
                file_name="claims_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="claims_download"
            )
        except Exception as e:
            st.error(f"Error: {e}")

elif tool == "🧾 Cloud Invoice Tool":
    st.title("Cloud Invoice Tool")
    uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"], key="cloud_invoice_upload")
    if uploaded_file:
        df = pd.read_csv(uploaded_file)
        st.write("Preview of uploaded file:", df.head())

        processed_df = process_cloud_invoice(df)

        output_buffer = io.BytesIO()
        processed_df.to_excel(output_buffer, index=False, engine='openpyxl')
        output_buffer.seek(0)

        st.download_button(
            label="Download processed Excel",
            data=output_buffer,
            file_name="cloud_invoice_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif tool == "Other":
    st.warning("Need a different tool? Just let us know what you need and we'll build it for you! 🚀")
    st.info("Currently, only the Google DNTS Extractor tool is available. More tools can be added based on your requirements.")

st.markdown("""
<footer style='text-align:center; margin-top:3rem; color:#1a73e8; font-size:20px; font-weight:bold; font-family: Google Sans, sans-serif;'>
    Made with ❤️ by Mindware | © 2025
</footer>
""", unsafe_allow_html=True)
