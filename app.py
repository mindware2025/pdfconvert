
import base64
import csv
import json
import os
import zipfile
import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from google.oauth2 import service_account
import gspread
from google.oauth2.service_account import Credentials
from extractors.aws import AWS_OUTPUT_COLUMNS, build_dnts_cnts_rows, process_multiple_aws_pdfs
from extractors.google_dnts import extract_invoice_info, extract_table_from_text, make_dnts_header_row, DNTS_HEADER_COLS, DNTS_ITEM_COLS
from utils.helpers import format_amount, format_invoice_date, format_month_year
from dotenv import load_dotenv
load_dotenv()
from extractors.google_invoice import extract_table_from_text as extract_invoice_table, extract_invoice_info as extract_invoice_info_invoice, GOOGLE_INVOICE_COLS
from extractors.dell_invoice import (
    extract_invoice_info as extract_dell_invoice_info,
    extract_table_from_text as extract_dell_table,
    DELL_INVOICE_COLS,
    PRE_ALERT_HEADERS,
    build_pre_alert_rows,
    read_master_mapping,
)
from oauth2client.service_account import ServiceAccountCredentials
from extractors.cloud_invoice import create_summary_sheet, build_cloud_invoice_df, map_invoice_numbers
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
import plotly.express as px

SHEET_JSON = "tool-mindware-7596713f2b86.json"  # Path to your downloaded JSON
SHEET_NAME = "Mindware tool usage"
json_base64 = """
ewogICJ0eXBlIjogInNlcnZpY2VfYWNjb3VudCIsCiAgInByb2plY3RfaWQiOiAidG9vbC1taW5kd2FyZSIsCiAgInByaXZhdGVfa2V5X2lkIjogIjc1OTY3MTNmMmI4NjUzN2E3MjAyMTBjNjA0MWMyNDU2MzFlMTY1NGEiLAogICJwcml2YXRlX2tleSI6ICItLS0tLUJFR0lOIFBSSVZBVEUgS0VZLS0tLS1cbk1JSUV2Z0lCQURBTkJna3Foa2lHOXcwQkFRRUZBQVNDQktnd2dnU2tBZ0VBQW9JQkFRQzFMOEJMUHdwYml0V1VcbnRsdGV2ektKZ3Y4ZThlQk5NQndCd25iRVlOSko4Vk5pb2NrRStVSHhzU2xPNk85VzNpN2ZTSlpsSzUvR3Z3d0lcbkloZU5Nc1kwT291anBIWUk0bUU4dk41MzJQZmNFa3J1aTRrQmUvYXBhSGFFRTQ5Z0w3V3dRM0I1MGM2STNVaGhcbnJIMFdqSWJ1aGpmWEw0NkhLbGdqdVIwRGl6YTQyVTV3ZkR3cERLMFhxaGNvYWpQRDZFUFdhblhJV0NicXd0aVRcbm1MRzIxQ0luM1FZNzA5UFRMQ0FHdEVka2ExWEs1ejEvMjFROHNmRDFndWx1OU04ODhBYVptUEVHQW5aVnVBQjlcbkc4QVNWcGpXT0pJVUtKTjU0QVNqZllHYUk3NjVkcnJ4dWprU3QvbjhDT2l3TlFId0hOZGtlSnFub1pQLzlNcndcblI3a0h1Yy9aQWdNQkFBRUNnZ0VBQXA4L05lK1BmWlpOMXRFN3JRWUVFbU50ODJOMzNGaUJBNmZPQkVCNDgyY0tcbmU0SXo2OEdIeWpRTlhjK2FVQmtXeXVtVEx3SEg3NkdORzJZZjU1UXVEY1grQm1RY0xVaDh6alNqMEVHKzhGSW5cbldZQ29tR1M2ZldBVXFaaG0rSEwwVGlNODZTTE4rbU1HYXc3WG1hNnA0KzN3M3ErNDRFblZLY0ZpSCs2SGdCUkFcbkZ3K29HZ3ZKQnN2QTNZYW5FcU5GSVJtUFRFWFd6TWZBZEdab1IvQlRSc29Yb1NOSEZSaHNJZDJtUDZ1L2xGK2JcbjdTSmxJWHgzSWFiY2NDSjYxT1dZTVBYTzdVMklmUTJqVC9YWk1yU29sdHhSUG9EY3NDU2NOc01lM3ZZbjBEbCtcbmRSUER1SkRTRU1FbjF6anpJWDl0MXFteFRjaEFCKzQ0YmEweWZvdS9zUUtCZ1FENlhhaS9BVDVYT3BtT0crVlFcbjhYOW92bmJMQlVVMjRUMUZsWlRpVzZEYklldDZZbUQwWGNnSVlUMkhLazZvY1dlemNyQ3R0VHVlN3F3ZEE2dmdcbkVXUmJqRmtoc2gyZ1Q0a28yKzBVT0dMY0FOWE1IbDBwVmZjTFlzaE5hK2dqeHorOFZqUDBhVHhBSnI0TFJBQVdcbm1QUFhxd0VoYmlSNnJPZThtZm5CZ0lGc3FRS0JnUUM1UTQzMTRJUStJUWxDQ0EzRkowWlZ4aXpuaUdCYUJmVldcbjBCV2t0YkpJNDNGSnc3dE0vL1hpQUxCTkNGd045SXRDdWYzSzBZQ3F5U0ZkOHI4czhxY1RPN3JVZGI3N1h4VEZcblBDWGM0MkNYYVJreEJNNnBxSG0rZHZJb0ZoRlhwWXhveEwvQlV0L256cmJUOUZFcWFJRzBtZEVZcXNSTE44TDNcbml3aVNQN2FYc1FLQmdRRFNOeCsvdUppU2Z6WjlWcmpWbk9BZ240TjQ5YVRtN25vVzJnQ1hpdDNtQUhZS1hWNFFcbjhFbExsL0lrY29aMjhqbGpOOUpYR0F2R1o1b0dCcFlpM2hlSXNyQUlGZGpBU09mZWNjSi9MdFQ2Nm95WkJZbXRcbmNtdXFtTGVjSWhWWkxTdzd3NWwrQjNvNlZ3MU13anpjdkhKSlRHRDNvOVpuVnBTQkREdmptRFdUZVFLQmdBTC9cbkdMaTFYTzQwVXBZQzAxWXhBRzQ2dWxjMFdYcWJSaENWWlFRNC9CMDVzSWRrNXc2anhUSldtSU5tY3phMmtkb09cbmNCQnJ1dzBJRzhZTk94SmJDbURCUXBCVkp6V2hvQkJnbkt3cDhWSUJuU3F4elRYcFI2N1E5Ykc0U2FlRlFmUWZcbjJvb2g4UVVxenNJMjNXazJMNExnU2dXQUhaU3Azamxxd2tTN1N4VEJBb0dCQU5wcVltZEZFWFFMOWZvNTIwNzVcbmsyUTdVRmdCSnhsZUZXc1VmbG84YmdXbmtXazcweVMwSW1vWEFwcHhzbWZvRjNmMnlBcHJWWGRIbHZ5Rm03aG5cbkpGQUp0anZIc21xM25uWlN3WE4xTFFyaFlqQVJrdmxNdUpBNllKbDk2UndybDl2dER0Zk5DVmp5blBNNzR3dk9cbnhndy9aR2dMcUQwTU5BZHRNVlhnNitwUFxuLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLVxuIiwKICAiY2xpZW50X2VtYWlsIjogInN0cmVhbWxpdC11c2FnZUB0b29sLW1pbmR3YXJlLmlhbS5nc2VydmljZWFjY291bnQuY29tIiwKICAiY2xpZW50X2lkIjogIjExNTU1MjA4NzI5MzczMjA3NzgyOCIsCiAgImF1dGhfdXJpIjogImh0dHBzOi8vYWNjb3VudHMuZ29vZ2xlLmNvbS9vL29hdXRoMi9hdXRoIiwKICAidG9rZW5fdXJpIjogImh0dHBzOi8vb2F1dGgyLmdvb2dsZWFwaXMuY29tL3Rva2VuIiwKICAiYXV0aF9wcm92aWRlcl94NTA5X2NlcnRfdXJsIjogImh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL29hdXRoMi92MS9jZXJ0cyIsCiAgImNsaWVudF94NTA5X2NlcnRfdXJsIjogImh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL3JvYm90L3YxL21ldGFkYXRhL3g1MDkvc3RyZWFtbGl0LXVzYWdlJTQwdG9vbC1taW5kd2FyZS5pYW0uZ3NlcnZpY2VhY2NvdW50LmNvbSIsCiAgInVuaXZlcnNlX2RvbWFpbiI6ICJnb29nbGVhcGlzLmNvbSIKfQo=
"""  # paste the full Base64 string here

# -----------------------------
# 2️⃣ Decode and load JSON
# -----------------------------
creds_dict = json.loads(base64.b64decode(json_base64))

# -----------------------------
# 3️⃣ Define the scopes
# -----------------------------
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

# -----------------------------
# 4️⃣ Authorize gspread
# -----------------------------
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
gc = gspread.authorize(creds)

# -----------------------------
# 5️⃣ Open your sheet
# -----------------------------
SHEET_NAME = "Mindware tool usage"
sheet = gc.open(SHEET_NAME).worksheet("Sheet1")
def update_tool_usage(tool_name):
    month = datetime.today().strftime("%b-%Y")
    
    # Read all existing rows
    all_rows = sheet.get_all_records()
    
    # Check if entry exists
    found = False
    for i, row in enumerate(all_rows, start=2):  # skip header row
        if row["tool"] == tool_name and row["Month"] == month:
            current_count = row.get("Usage Count", 0) or 0
            sheet.update_cell(i, 3, current_count + 1)  # Usage Count column
            found = True
            break
    
    # If not found, append a new row
    if not found:
        sheet.append_row([tool_name, month, 1])


st.set_page_config(
    page_title="Mindware Tool",
    layout="wide",
    initial_sidebar_state="collapsed"
)
# Sidebar content
with st.sidebar:
    st.markdown("### ⚙️ Admin Panel")
    admin_mode = st.checkbox("Show Tool Usage Analytics", value=False)


# CSS: hide top-right icons but keep sidebar visible
st.markdown("""
    <style>
    /* Hide Share, GitHub, Settings icons on top-right */
    [data-testid="stToolbar"] {
        display: none !important;
    }
    /* Ensure sidebar is always visible */
    [data-testid="stSidebar"] {
        visibility: visible !important;
        min-width: 280px !important;
    }
    /* Optional: adjust sidebar content font */
    [data-testid="stSidebar"] * {
        font-family: 'Google Sans', sans-serif !important;
    }
    /* General styling */
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
else:
    if admin_mode:
        st.title("📊 Tool Usage Analytics (Admin Mode)")
        
        if os.path.exists("tool_usage.xlsx"):
            wb = load_workbook("tool_usage.xlsx")
            if "USAGE_REPORT" in wb.sheetnames:
                ws = wb["USAGE_REPORT"]
                
                # --- READ ONLY, no incrementing ---
                usage_data = [row for row in ws.iter_rows(min_row=2, values_only=True)]
                usage_df = pd.DataFrame(usage_data, columns=["tool", "Month", "Usage Count"])
                st.dataframe(usage_df)
                
                # --- DOWNLOAD BUTTON ---
                excel_io = io.BytesIO()
                with pd.ExcelWriter(excel_io, engine="openpyxl") as writer:
                    usage_df.to_excel(writer, index=False, sheet_name="Tool Usage")
                excel_io.seek(0)
                st.download_button(
                    "⬇️ Download Tool Usage Report",
                    data=excel_io.getvalue(),
                    file_name="tool_usage_report.xlsx"
                )
                
                # --- PLOTLY CHARTS ---
                import plotly.express as px
                
                st.subheader("📈 Tool Usage per Month")
                fig_bar = px.bar(
                    usage_df, 
                    x="Month", 
                    y="Usage Count", 
                    color="tool",
                    barmode="group",
                    text="Usage Count",
                    title="Tool Usage per Month"
                )
                fig_bar.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_bar, use_container_width=True)
                
                st.subheader("📊 Overall Tool Usage Share")
                pie_data = usage_df.groupby("tool")["Usage Count"].sum().reset_index()
                fig_pie = px.pie(
                    pie_data, 
                    values="Usage Count", 
                    names="tool", 
                    title="Overall Tool Usage Share", 
                    hole=0.3
                )
                st.plotly_chart(fig_pie, use_container_width=True)
                
                st.subheader("📉 Monthly Trend per Tool")
                trend_data = usage_df.pivot_table(
                    index="Month", columns="tool", values="Usage Count", fill_value=0
                ).reset_index()
                fig_line = px.line(
                    trend_data, 
                    x="Month", 
                    y=trend_data.columns[1:],  # all tools
                    markers=True,
                    title="Monthly Usage Trend per Tool"
                )
                st.plotly_chart(fig_line, use_container_width=True)
                
        else:
            st.info("No usage data found yet.")
        st.stop()
    else:
        team = st.radio("👥 Select your team:", ["Finance", "Operations"], horizontal=True)
# def update_tool_usage(tool_name):
#     month = datetime.today().strftime("%b-%Y")
#     if os.path.exists("tool_usage.xlsx"):
#         wb = load_workbook("tool_usage.xlsx")
#         if "USAGE_REPORT" not in wb.sheetnames:
#             ws = wb.create_sheet("USAGE_REPORT")
#             ws.append(["tool", "Month", "Usage Count"])
#         else:
#             ws = wb["USAGE_REPORT"]
#     else:
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "USAGE_REPORT"
#         ws.append(["tool", "Month", "Usage Count"])
#     # Check if entry exists
#     found = False
#     for row in ws.iter_rows(min_row=2):
#         if row[0].value == tool_name and row[1].value == month:
#             row[2].value = (row[2].value or 0) + 1
#             found = True
#             break
#     if not found: ws.append([tool_name, month, 1])
#     wb.save("tool_usage.xlsx")
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
                    on_click=lambda: update_tool_usage("google Automation"),
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
if team == "Finance":
    TOOL_OPTIONS = [
        "-- Select a tool --",
        "🟦 Google DNTS Extractor",
        "🟩 Google Invoice Extractor",
        "📄 Claims Automation",
        "🟨 AWS Invoice Tool"
    ]
elif team == "Operations":
    TOOL_OPTIONS = [
        "-- Select a tool --",
        "💻 Dell Invoice Extractor",
        "🧾 Cloud Invoice Tool"
    ]
else:
    TOOL_OPTIONS = ["-- Select a tool --"]
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
                on_click=lambda: update_tool_usage("Claims Automation"),
                key="claims_download"
            )
        except Exception as e:
            st.error(f"Error: {e}")
elif tool == "🧾 Cloud Invoice Tool":
    st.title("Cloud Invoice Tool")
    st.markdown(
        """
        <div style="
            padding: 18px 20px;
            background: linear-gradient(90deg, #fff3cd, #ffeeba);
            border: 2px solid #ffcc00;
            border-radius: 10px;
            font-weight: 700;
            color: #7a5a00;
            font-size: 16px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            margin-bottom: 12px;
        ">
            <span style="font-size: 18px;">🚨 IMPORTANT:</span>
            <span style="margin-left: 8px;">
            Please make sure to <b>open the CB file</b>, click <b>Convert</b>, then <b>upload the converted file here</b> and use this tool; otherwise <b>you will have missing invoices</b>.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div style="
            padding: 14px 16px;
            background: #fff;
            border: 1px dashed #ffcc00;
            border-radius: 10px;
            color: #4a4a4a;
            font-size: 15px;
            margin-bottom: 8px;
        ">
        <b>Follow these steps before uploading:</b>
        <ol style="margin-top: 6px;">
            <li>Open the <b>CB file</b>.</li>
            <li>Click <b>Convert</b> to generate the latest output.</li>
            <li>Upload the <b>converted file</b> here.</li>
        </ol>
        </div>
        """,
        unsafe_allow_html=True,
    )
    confirmed = st.checkbox("I confirm I opened the CB file and clicked Convert ✅", key="cloud_cb_confirm")
    if not confirmed:
        st.warning("Please confirm the IMPORTANT notice steps above to proceed.")
        st.stop()
    uploaded_file = st.file_uploader("Upload your CSV file", type=["csv", "xlsx"], key="cloud_invoice_upload")
    if uploaded_file:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            st.error("Unsupported file format. Please upload a CSV or Excel file.")
            st.stop()
        # Process invoice data
        final_df = build_cloud_invoice_df(df)
        final_df = map_invoice_numbers(final_df)
        sorted_df = final_df.sort_values(by=final_df.columns.tolist()).reset_index(drop=True)
        def highlight_row(row):
            end_user = str(row.get("End User", "")).strip()
            return not (" ; " in end_user)
        sorted_df["_highlight_end_user"] = sorted_df.apply(highlight_row, axis=1)
        # Create unique version rows based on Combined (D)
        unique_rows = sorted_df[["Invoice No.","LPO Number", "End User"]].copy()
        unique_rows["Combined (D)"] = (
            unique_rows["Invoice No."].astype(str) +
            unique_rows["LPO Number"].astype(str) +
            unique_rows["End User"].astype(str)
        )
        unique_rows = unique_rows.drop_duplicates(subset=["Combined (D)"]).reset_index(drop=True)
        # Versioning logic
        unique_rows["Version1 (E)"] = (unique_rows["Invoice No."].ne(unique_rows["Invoice No."].shift()).astype(int))
        v2 = []
        for i, v1 in enumerate(unique_rows["Version1 (E)"]):
            if v1 == 1:
                v2.append(1)
            else:
                prev_v2 = v2[-1]
                v2.append(prev_v2 + 1)
        unique_rows["Version2 (F)"] = v2
        unique_rows["Version3 (G)"] = unique_rows.apply(lambda row: f'-{row["Version2 (F)"]}', axis=1)
        unique_rows["Version4 (H)"] = unique_rows.apply(lambda row: f'{row["Invoice No."]}-{row["Version2 (F)"]}', axis=1)
        # --- MAP Version 4 back to main DataFrame ---
        version_map = dict(zip(unique_rows["Combined (D)"], unique_rows["Version4 (H)"]))
        sorted_df["Combined (D)"] = (
            sorted_df["Invoice No."].astype(str) +
            sorted_df["LPO Number"].astype(str) +
            sorted_df["End User"].astype(str)
        )
        sorted_df["Versioned Invoice No."] = sorted_df["Combined (D)"].map(version_map)
        cols = list(sorted_df.columns)
        cols.append(cols.pop(cols.index("Versioned Invoice No.")))
        sorted_df = sorted_df[cols]
        sorted_df = sorted_df.drop(columns=["Combined (D)"])
        # === ADD HIGHLIGHT FLAG HERE ===
        #sorted_df["_highlight_end_user"] = sorted_df["End User"].astype(str).str.strip() == ""
        # Display metrics
        pos_df = sorted_df[sorted_df["Gross Value"].astype(float) >= 0]
        neg_df = sorted_df[sorted_df["Gross Value"].astype(float) < 0]
        st.success(f"{len(pos_df)} positive, {len(neg_df)} negative, total: {len(sorted_df)}")
        c1, c2, c3 = st.columns(3)
        c1.metric("✅ Positive invoices", len(pos_df))
        c2.metric("❌ Negative invoices", len(neg_df))
        c3.metric("🧮 Total invoices", len(sorted_df))
        # DataFrame previews
        st.subheader("Processed Preview")
        st.dataframe(sorted_df.head(50))
        st.subheader("Versions Sheet Preview")
        st.dataframe(unique_rows.head(50))
        
        # Create Excel workbook with formulas
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        
        # Remove the internal helper column before saving
        if "_highlight_end_user" in sorted_df.columns:
            df_to_write = sorted_df.drop(columns=["_highlight_end_user"])
        else:
            df_to_write = sorted_df.copy()
        
        # Get index of 'End User' column (1-based for Excel)
        try:
            end_user_col_idx = df_to_write.columns.get_loc("End User") + 1
        except:
            end_user_col_idx = None
            
        try:
            item_code_col_idx = df_to_write.columns.get_loc("ITEM Code") + 1
        except:
            item_code_col_idx = None    
        
        # Create workbook and write rows
        wb = Workbook()
        ws_invoice = wb.active
        ws_invoice.title = "CLOUD INVOICE"
        
        for r_idx, row in enumerate(dataframe_to_rows(df_to_write, index=False, header=True), start=1):
            ws_invoice.append(row)
            
            # Skip header row
            if r_idx == 1:
              continue
        # Highlight End User
            if end_user_col_idx is not None:
                highlight = sorted_df.iloc[r_idx - 2].get("_highlight_end_user", False)
                if highlight:
                    col_letter = get_column_letter(end_user_col_idx)
                    ws_invoice[f"{col_letter}{r_idx}"].fill = red_fill
 
        # Highlight ITEM Code if empty
            if item_code_col_idx is not None:
                item_code_val = sorted_df.iloc[r_idx - 2].get("ITEM Code", "")
                if not item_code_val or str(item_code_val).strip().lower() in ["", "nan", "none"]:
                   col_letter = get_column_letter(item_code_col_idx)
                   ws_invoice[f"{col_letter}{r_idx}"].fill = red_fill
        
        # Create VERSIONS sheet with formulas
        ws_versions = wb.create_sheet(title="VERSIONS")
        headers = ["Invoice",  "LPO", "End User", "Combined (D)", "Version1 (E)", "Version2 (F)", "Version3 (G)", "Version4 (H)"]
        ws_versions.append(headers)
        for i, row in enumerate(unique_rows.itertuples(index=False, name=None), start=2):
            invoice, lpo, end_user, combined_d = row[:4]
            ws_versions.cell(row=i, column=1, value=invoice)
            ws_versions.cell(row=i, column=2, value=lpo)
            ws_versions.cell(row=i, column=3, value=end_user)
            ws_versions.cell(row=i, column=4, value=combined_d)
            ws_versions.cell(row=i, column=5, value=f'=IF(A{i}=A{i-1},"",1)')
            ws_versions.cell(row=i, column=6, value=f'=IFERROR(IF(E{i}="",E{i-1}+1,""),F{i-1}+1)')
            ws_versions.cell(row=i, column=7, value=f'="-"&E{i}&F{i}')
            ws_versions.cell(row=i, column=8, value=f'=A{i}&G{i}')
        
        # Save to buffer
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        st.download_button(
            label="⬇️ Download Cloud Invoice",
            data=output_buffer.getvalue(),
            file_name="cloud_invoice.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=lambda: update_tool_usage("Cloud Automation")
            
        )
        

elif tool == "💻 Dell Invoice Extractor":
    st.title("Dell Invoice Extractor (Pre-Alert Upload)")
    st.write("Upload one or more Dell invoice PDFs. We'll generate a single Excel with sheet 'PRE ALERT UPLOAD'.")
    uploaded_files = st.file_uploader("Choose Dell invoice PDF(s)", type=["pdf"], accept_multiple_files=True, key="dell_upload")
    master_file = st.file_uploader("Master Excel (starts header at row 9)", type=["xlsx"], key="dell_master")
    if uploaded_files:
        from datetime import datetime, timedelta
        tomorrow_date = (datetime.today() + timedelta(days=1)).strftime("%d/%m/%Y")
        all_rows = []
        master_lookup = None
        supplier_counts = None
        orion_counts = None
        supplier_index = None
        orion_index = None
        po_price_index = None
        if master_file is not None:
            try:
                master_lookup, supplier_counts, orion_counts, supplier_index, orion_index, po_price_index = read_master_mapping(master_file)
            except Exception as e:
                st.warning(f"Could not read master file: {e}")
        diag: list[dict] = []
        for f in uploaded_files:
            try:
                rows = build_pre_alert_rows(
                    f,
                    tomorrow_date,
                    master_lookup=master_lookup,
                    supplier_counts=supplier_counts,
                    orion_counts=orion_counts,
                    supplier_index=supplier_index,
                    orion_index=orion_index,
                    po_price_index=po_price_index,
                    diagnostics=diag,
                )
                all_rows.extend(rows)
            except Exception as e:
                st.warning(f"Failed to parse {getattr(f, 'name', 'file')}: {e}")
        if all_rows:
            df = pd.DataFrame(all_rows, columns=PRE_ALERT_HEADERS)
            st.subheader("PRE ALERT UPLOAD Preview")
            st.dataframe(df, height=300)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='PRE ALERT UPLOAD', index=False)
                ws = writer.sheets['PRE ALERT UPLOAD']
                from datetime import datetime as _dt
                for r in range(2, len(df) + 2):
                    for c in (4, 11, 12):  # D, K, L
                        cell = ws.cell(row=r, column=c)
                        val = cell.value
                        if isinstance(val, str):
                            try:
                                d = _dt.strptime(val, "%d/%m/%Y")
                                cell.value = d
                                cell.number_format = 'dd/mm/yyyy'
                            except Exception:
                                pass
                # ====== NEW: Highlight Unit Rate (Q) vs Orion Unit Price (U) in PRE ALERT UPLOAD ======
                from openpyxl.styles import PatternFill
                red_fill = PatternFill(fill_type='solid', start_color='FFFFC7CE', end_color='FFFFC7CE')  # light red
                green_fill = PatternFill(fill_type='solid', start_color='FFC6EFCE', end_color='FFC6EFCE')  # light green
                
                def safe_float(val):
                    try:
                        return float(str(val).replace(',', '').strip())
                    except Exception:
                        return None
                # Columns: Q = 17, U = 21
                for row_idx in range(2, len(df) + 2):
                    unit_rate_cell = ws.cell(row=row_idx, column=17)
                    orion_price_cell = ws.cell(row=row_idx, column=21)
                    unit_rate_val = safe_float(unit_rate_cell.value)
                    orion_price_val = safe_float(orion_price_cell.value)
                    if unit_rate_val is not None and orion_price_val is not None:
                        if abs(unit_rate_val - orion_price_val) < 1e-6:
                            # Match: highlight green
                            unit_rate_cell.fill = green_fill
                            orion_price_cell.fill = green_fill
                        else:
                            # Mismatch: highlight red
                            unit_rate_cell.fill = red_fill
                            orion_price_cell.fill = red_fill
                # ====== EXISTING color highlights on PRE ALERT UPLOAD based on diagnostics ======
                if diag:
                    red = PatternFill(fill_type='solid', start_color='FFFF0000', end_color='FFFF0000')
                    yellow = PatternFill(fill_type='solid', start_color='FFFFFF00', end_color='FFFFFF00')
                    green = PatternFill(fill_type='solid', start_color='FF00FF00', end_color='FF00FF00')
                    for i, d in enumerate(diag, start=2):
                        h = str(d.get('highlight', 'none')).lower()
                        fill = None
                        if h == 'red':
                            fill = red
                        elif h == 'yellow':
                            fill = yellow
                        elif h == 'green':
                            fill = green
                        if fill is not None:
                            ws.cell(row=i, column=13).fill = fill
                            ws.cell(row=i, column=14).fill = fill
                # Add an extremely user-friendly REVIEW sheet (emoji + clear actions)
                if diag:
                    def _status_texts(h: str, status: str) -> tuple[str, str, str, int]:
                        h = (h or 'none').lower()
                        status = status or ''
                        if h == 'red':
                            if status in ('C_orion_price_single', 'C_po_price_single'):
                                return (
                                    '❌ No match (price suggested)',
                                    'Use the suggested Orion code. Confirm Qty & Unit Price.',
                                    'We did not find a supplier match, but price uniquely matched one item.',
                                    0,
                                )
                            return (
                                '❌ No match',
                                'Ask to add mapping in MASTER or correct Supplier Item code in PDF.',
                                'We could not find this PO + Supplier Item in MASTER.',
                                0,
                            )
                        if h == 'yellow':
                            if status == 'B_price_single':
                                return (
                                    '✅ Price matched (needs quick check)',
                                    'Quick check the suggested Orion Code, Qty, and Unit Price.',
                                    'Multiple entries in MASTER; price selected exactly one.',
                                    1,
                                )
                            return (
                                '⚠️ Many matches',
                                'Pick the correct Orion code or ask to refine MASTER.',
                                'Multiple MASTER entries matched; price did not decide one.',
                                1,
                            )
                        return (
                            '✅ All good',
                            'No action. Just review and proceed.',
                            'Single clear match found in MASTER.',
                            2,
                        )
                    simple_rows = []
                    for idx0, d in enumerate(diag):
                        h = str(d.get('highlight', 'none')).lower()
                        status = str(d.get('status', '')).strip()
                        label, action, why, sort_key = _status_texts(h, status)
                        try:
                            row_df = df.iloc[idx0]
                        except Exception:
                            row_df = None
                        qty_val = row_df['Qty'] if row_df is not None and 'Qty' in row_df else ''
                        pdf_price = row_df['Unit Rate'] if row_df is not None and 'Unit Rate' in row_df else d.get('pdf_unit_price', '')
                        use_price = d.get('out_orion_unit_price', '') or pdf_price
                        orion_price_u = d.get('out_orion_unit_price', '')
                        use_code = d.get('out_orion_item_code', '') or d.get('mapped_item_code', '')
                        pre_alert_row_num = idx0 + 2  # header row is 1 in PRE ALERT
                        simple_rows.append({
                            'Row in PRE ALERT': pre_alert_row_num,
                            'Status': label,
                            'What to do now': action,
                            'Why this happened': why,
                            'PO': d.get('po', ''),
                            'Supplier Item': d.get('supplier_item_code', ''),
                            'Copy: Orion Code': use_code,
                            'Copy: Qty': qty_val,
                            'Copy: Unit Price': use_price,
                            'Orion Unit Price (col U)': orion_price_u,
                            '_sort': sort_key,
                            '_highlight': h,
                        })
                    simple_rows.sort(key=lambda x: x.get('_sort', 1))
                    for r in simple_rows:
                        r.pop('_sort', None)
                    review_columns = [
                        'Row in PRE ALERT',
                        'Status',
                        'What to do now',
                        'Why this happened',
                        'Copy: Orion Code',
                        'Copy: Qty',
                        'Copy: Unit Price',
                        'Orion Unit Price (col U)',
                        'PO',
                        'Supplier Item',
                    ]
                    review_df = pd.DataFrame(simple_rows, columns=review_columns + ['_highlight'])
                    review_df.to_excel(writer, sheet_name='REVIEW', index=False)
                    ws_review = writer.sheets['REVIEW']
                    ws_review.freeze_panes = 'A2'
                    widths = {
                        'A': 16, 'B': 30, 'C': 38, 'D': 42,
                        'E': 20, 'F': 12, 'G': 18, 'H': 12, 'I': 20,
                    }
                    for col, w in widths.items():
                        try:
                            ws_review.column_dimensions[col].width = w
                        except Exception:
                            pass
                    try:
                        from openpyxl.styles import Alignment
                        wrap_align = Alignment(wrap_text=True, vertical='top')
                        for r in ws_review.iter_rows(min_row=1, max_row=ws_review.max_row, min_col=1, max_col=ws_review.max_column):
                            for cell in r:
                                if cell.column in (2, 3, 4):  # B:Status, C:What to do, D:Why
                                    cell.alignment = wrap_align
                    except Exception:
                        pass
                    red_fill_row = PatternFill(fill_type='solid', start_color='FFFFE5E5', end_color='FFFFE5E5')
                    yellow_fill_row = PatternFill(fill_type='solid', start_color='FFFFFBE6', end_color='FFFFFBE6')
                    green_fill_row = PatternFill(fill_type='solid', start_color='FFE9FBE9', end_color='FFE9FBE9')
                    for idx1, row_data in enumerate(simple_rows, start=2):
                        hval = str(row_data.get('_highlight', '')).lower()
                        fill = None
                        if hval == 'red':
                            fill = red_fill_row
                        elif hval == 'yellow':
                            fill = yellow_fill_row
                        elif hval == 'green':
                            fill = green_fill_row
                        if fill is not None:
                            for c in range(1, ws_review.max_column + 1):
                                ws_review.cell(row=idx1, column=c).fill = fill
                    header_map = {str(ws_review.cell(row=1, column=c).value): c for c in range(1, ws_review.max_column + 1)}
                    orion_u_col = header_map.get('Orion Unit Price (col U)')
                    red_cell = PatternFill(fill_type='solid', start_color='FFFFC7CE', end_color='FFFFC7CE')
                    green_cell = PatternFill(fill_type='solid', start_color='FFC6EFCE', end_color='FFC6EFCE')
                    def _as_float(x):
                        try:
                            return float(str(x).replace(',', '').strip())
                        except Exception:
                            return None
                    if orion_u_col:
                        for idx0 in range(len(simple_rows)):
                            r_excel = idx0 + 2
                            try:
                                row_df = df.iloc[idx0]
                            except Exception:
                                row_df = None
                            pdf_price_val = _as_float(row_df['Unit Rate']) if row_df is not None and 'Unit Rate' in row_df else _as_float(simple_rows[idx0].get('Copy: Unit Price', ''))
                            orion_price_val = _as_float(simple_rows[idx0].get('Orion Unit Price (col U)', ''))
                            if orion_price_val is not None and pdf_price_val is not None and abs(orion_price_val - pdf_price_val) < 1e-6:
                                ws_review.cell(row=r_excel, column=orion_u_col).fill = green_cell
                            else:
                                ws_review.cell(row=r_excel, column=orion_u_col).fill = red_cell
                # Create COMPONENT UPLOAD sheet with only the header row
                component_headers = [
                    'PO Txn Code',
                    'PO Number',
                    'Parent Item Code',
                    'Component Item Code',
                    'UOM',
                    'Qty',
                    'Rate',
                ]
                pd.DataFrame(columns=component_headers).to_excel(
                    writer, sheet_name='COMPONENT UPLOAD', index=False
                )
                # Write master file content as sheet 2 if provided
                if master_file is not None:
                    try:
                        master_file.seek(0)
                    except Exception:
                        pass
                    try:
                        df_master = pd.read_excel(master_file, header=8)
                        df_master.to_excel(writer, sheet_name='MASTER', index=False)
                    except Exception:
                        pass
            output.seek(0)
            st.download_button(
                label="⬇️ Download PRE ALERT UPLOAD",
                data=output.getvalue(),
                file_name="pre_alert_upload.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_dell_pre_alert",
                on_click=lambda: update_tool_usage("DEll Automation")
            )
        else:
            st.warning("No items found in the uploaded PDF(s).")
elif tool == "🟨 AWS Invoice Tool":
        st.title("AWS Invoice Tool")
        st.write("Upload AWS invoice PDF(s) and download the extracted data as Excel.")
    
        uploaded_files = st.file_uploader(
            "Choose AWS invoice PDF(s)", type=["pdf"], key="aws_upload", accept_multiple_files=True
        )
    
        if uploaded_files:
            rows, template_map, text_map = process_multiple_aws_pdfs(uploaded_files)
            if rows:
                df = pd.DataFrame(rows, columns=AWS_OUTPUT_COLUMNS)
                st.subheader("Extracted AWS Invoice Data")
                st.dataframe(df, height=300)
    
                output_original = io.BytesIO()
                with pd.ExcelWriter(output_original, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='AWS_INVOICE', index=False)
                output_original.seek(0)
                st.download_button(
                    label="⬇️ Download Extracted AWS Invoice Data",
                    data=output_original.getvalue(),
                    file_name="aws_invoice_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    
                )
    
                output_files = build_dnts_cnts_rows(rows, template_map, text_map)
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_key, data in output_files.items():
                        bill_to, file_type = file_key.split("__")
                        safe_bill_to = bill_to.replace(" ", "_").replace(".", "").replace(",", "")
                        file_name = f"{file_type}_{safe_bill_to}.xlsx"
                
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            pd.DataFrame(data["header"], columns=[
                                "S.No", "Date - (dd/MM/yyyy)", "Supp_Code", "Curr_Code", "Form_Code",
                                "Doc_Src_Locn", "Location_Code", "Remarks", "Supplier_Ref", "Supplier_Ref_Date - (dd/MM/yyyy)"
                            ]).to_excel(writer, sheet_name=f"{file_type}_HEADER", index=False)
                
                            pd.DataFrame(data["item"], columns=[
                                "S.No", "Ref. Key", "Item_Code", "Item_Name", "Grade1", "Grade2", "UOM",
                                "Qty", "Qty_Ls", "Rate", "Main_Account", "Sub_Account", "Division", "Department", "Analysis2"
                            ]).to_excel(writer, sheet_name=f"{file_type}_ITEM", index=False)
                
                        output.seek(0)
                        zip_file.writestr(file_name, output.read())
                
                zip_buffer.seek(0)
                
                st.download_button(
                    label="⬇️ Download All DNTS/CNTS Files as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="aws_dnts_cnts_files.zip",
                    mime="application/zip",
                    on_click=lambda: update_tool_usage("AWS Automation")
                )
            else:
                st.warning("No data extracted from the uploaded AWS PDFs.")
        else:
            st.info("Please upload one or more AWS invoice PDFs to begin.")

elif tool == "Other":
    st.warning("Need a different tool? Just let us know what you need and we'll build it for you! 🚀")
    st.info("Currently, only the Google DNTS Extractor tool is available. More tools can be added based on your requirements.")



st.markdown("""
<footer style='text-align:center; margin-top:3rem; color:#1a73e8; font-size:20px; font-weight:bold; font-family: Google Sans, sans-serif;'>
    Made with ❤️ by Mindware | © 2025
</footer>
""", unsafe_allow_html=True)
