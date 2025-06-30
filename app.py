import streamlit as st
import pandas as pd
import pdfplumber
import re
import tempfile
import os
import io
from datetime import datetime

# ----------- Page config -----------
st.set_page_config(
    page_title="Google DNTS upload file",
    layout="wide"
)

# Hide Streamlit default menu, footer, and header
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ----------- Minimal, Modern CSS Styling -----------
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
DNTS_HEADER_COLS = [
    "S.No", "Date - (dd/MM/yyyy)", "Supp_Code", "Curr_Code", "Form_Code",
    "Doc_Src_Locn", "Location_Code", "Remarks", "Supplier_Ref", "Supplier_Ref_Date - (dd/MM/yyyy)"
]

DNTS_ITEM_COLS = [
    "S.No", "Ref. Key", "Item_Code", "Item_Name", "Grade1", "Grade2", "UOM", "Qty", "Qty_Ls",
    "Rate", "Main_Account", "Sub_Account", "Division", "Department", "Analysis-2"
]

DEFAULTS = {
    "supp_code": "SDIG005",
    "curr_code": "USD",
    "form_code": 0,
    "doc_src_locn": "UJ000",
    "location_code": "UJ200"
}

# ----------- Simple Login Page -----------
CORRECT_USERNAME = "S.Bhaskaran"
CORRECT_PASSWORD = "BHASAKHRAN@str2@25z#"

if "login_state" not in st.session_state:
    st.session_state.login_state = "login"  # can be 'login', 'fail', or 'success'

def show_login():
    # Add vertical space to help with vertical centering
    for _ in range(10):
        st.write("")
    # Use columns to center horizontally
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

# ----------- Utils -----------

def normalize_line(line):
    line = re.sub(r"[.]+", "", line)
    return re.sub(r"\s+", " ", line).strip()

def format_invoice_date(date_str):
    try:
        dt = datetime.strptime(date_str, "%d %b %Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        dt = datetime.strptime(date_str, "%d %B %Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str

def extract_invoice_info(pdf_path, debug_lines_callback=None):
    """Extract invoice number and invoice date from the section after 'Details' on the first page of the PDF. Optionally call debug_lines_callback with the lines after 'Details'."""
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        if not text:
            return None, None
        lines = text.splitlines()
        invoice_number = None
        invoice_date = None
        found_details = False
        details_lines = []
        for line in lines:
            if found_details:
                details_lines.append(line)
                norm_line = normalize_line(line)
                if invoice_number is None and "Invoice number" in norm_line:
                    match = re.search(r"Invoice number\s*:?\s*(\d{6,})", norm_line)
                    if not match:
                        match = re.search(r"Invoice number\s*:?\s*([0-9]+)", norm_line)
                    if match:
                        invoice_number = match.group(1)
                if invoice_date is None and "Invoice date" in norm_line:
                    match = re.search(r"Invoice date\s*:?\s*([0-9]{1,2} [A-Za-z]+ [0-9]{4}|[0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", norm_line)
                    if match:
                        invoice_date = match.group(1)
                if invoice_number and invoice_date:
                    break
            if 'Details' in line:
                found_details = True
        return invoice_number, invoice_date

def extract_table_from_text(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.splitlines()
            in_table = False
            for i, line in enumerate(lines):
                if "Summary of costs by domain" in line:
                    in_table = True
                    continue
                if in_table:
                    if re.match(r"\d{1,2} \w+ \d{4} - \d{1,2} \w+ \d{4}", line):
                        continue
                    if all(h in line for h in ["Domain name", "Customer ID", "Amount"]):
                        continue
                    m = re.match(r"^([\w\-.]+)\s+(C\w+)\s+([\d,]+\.\d{2})$", line.strip(), re.IGNORECASE)
                    if m:
                        domain, customer_id, amount = m.groups()
                        rows.append([domain, customer_id, amount])
                    elif line.strip() == '' or 'Subtotal' in line:
                        in_table = False
    return rows

def format_amount(amount_str):
    # Remove commas, convert to float, then format as int if possible, else as float
    try:
        amount = float(amount_str.replace(",", ""))
        if amount.is_integer():
            return str(int(amount))
        else:
            return str(amount).rstrip('0').rstrip('.') if '.' in str(amount) else str(amount)
    except Exception:
        return amount_str

def make_dnts_header_row(invoice_number, invoice_date, today_str, remarks):
    return [
        1,
        today_str,
        "SDIG005",
        "USD",
        0,
        "UJ000",
        "UJ200",
        remarks,
        remarks,
        format_invoice_date(invoice_date) if invoice_date else ""
    ]

# ----------- Streamlit UI -----------

st.title("PDF TO EXCEL")
st.write("Upload one PDF containing a **'Summary of costs by domain'** table. The app will extract the table and let you download it as Excel.")

uploaded_file = st.file_uploader("Choose your Google DNTS Invoice PDF", type=["pdf"], accept_multiple_files=False)

if uploaded_file:
    # Extract invoice info
    invoice_num, invoice_date = extract_invoice_info(uploaded_file)
    today_str = datetime.today().strftime("%d/%m/%Y")
    remarks = f"GOOGLE INV-{invoice_num}" if invoice_num else "GOOGLE INV-UNKNOWN"

    if invoice_num and invoice_date:
        st.success(f"Invoice Number: **{invoice_num}** | Invoice Date: **{format_invoice_date(invoice_date)}**")
    else:
        st.warning("Could not extract Invoice Number or Invoice Date from PDF header.")

    # Prepare DNTS header dataframe
    header_df = pd.DataFrame([make_dnts_header_row(invoice_num, invoice_date, today_str, remarks)], columns=DNTS_HEADER_COLS)
    st.subheader("DNTS Header Preview")
    st.dataframe(header_df, height=120)

    # Extract and show the table
    table_rows = extract_table_from_text(uploaded_file)
    dnts_item_data = []
    for idx, (domain, customer_id, amount) in enumerate(table_rows, 1):
        formatted_amount = format_amount(amount)
        item_name = (
            f"GOOGLE INV-{invoice_num} / DOMAIN NAME : {domain} / CUSTOMER ID : {customer_id} / AMOUNT - USD - {formatted_amount}"
        ).upper()
        dnts_item_data.append([
            idx,  # S.No
            1,    # Ref. Key
            "NS", # Item_Code
            item_name, # Item_Name
            "NA", # Grade1
            "NA", # Grade2
            "NOS", # UOM
            1,     # Qty
            0,     # Qty_Ls
            formatted_amount,   # Rate
            14401, # Main_Account
            "SDIG005", # Sub_Account
            "PUHO", # Division
            "GEN",  # Department
            "ZZ-COMM" # Analysis-2
        ])
    dnts_item_df = pd.DataFrame(dnts_item_data, columns=DNTS_ITEM_COLS)

    st.subheader("DNTS Items Preview")
    st.dataframe(dnts_item_df)

    # Download Excel (bottom button)
    st.markdown("## 📥 Download your Excel file below:")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        header_df.to_excel(writer, sheet_name='DNTS_HEADER', index=False)
        dnts_item_df.to_excel(writer, sheet_name='DNTS_ITEM', index=False)
    output.seek(0)
    st.download_button(
        label=f"⬇️ Download DNTS Excel",
        data=output.getvalue(),
        file_name=f"DNTS_Invoice_{invoice_num or 'unknown'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_bottom"
    )
else:
    st.info("Please upload a Google DNTS invoice PDF file to get started.")

st.markdown("""
<footer style='text-align:center; margin-top:3rem; color:#1a73e8; font-size:20px; font-weight:bold; font-family: Google Sans, sans-serif;'>
    Made with ❤️ by Mindware | © 2025
</footer>
""", unsafe_allow_html=True)
