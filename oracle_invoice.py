# oracle_tool.py
"""
Oracle Tool: PDF Invoice Extraction and Excel Export
"""

import pdfplumber
import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from typing import List, Dict
import concurrent.futures
import time

from typing import Tuple
REQUIRED_FIELDS = [
    "Billed To",
    "Invoice Number",
    "Invoice Date",
    "Due Date",
    "Purchase Order",
    "Invoice Amount",
    "Currency Code",
    "IBAN #",
    "ACCT #",
    "SWIFT Code"
]

def extract_text_from_pdf(pdf_path: str) -> str:
    doc = fitz.open(stream=pdf_path.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_fields(text: str) -> Dict[str, str]:
    lines = text.splitlines()
    result = {field: "" for field in REQUIRED_FIELDS}

    # Billed To: match from known list, search after 'Reference Invoice Number:' or in address blocks
    billed_to_options = [
        "Mindware for Computers",
        "Mindware Technology Trading LLC",
        "Mind Ware S.A",
        "Mindware Limited",
        "Mindware SPC",
        "Mindware FZ LLC"
    ]
    found_billed_to = False
    for idx, line in enumerate(lines):
        # Check after 'Reference Invoice Number:'
        if 'Reference Invoice Number:' in line:
            # Look ahead up to 2 lines for company name
            for look_ahead in range(1, 3):
                if idx + look_ahead < len(lines):
                    next_line = lines[idx + look_ahead].strip()
                    for option in billed_to_options:
                        if option.lower() in next_line.lower():
                            result["Billed To"] = option
                            found_billed_to = True
                            break
                    if found_billed_to:
                        break
        # Check in current and next line for multi-line company names
        for option in billed_to_options:
            if option.lower() in line.lower():
                result["Billed To"] = option
                found_billed_to = True
                break
            # Check next line for multi-line
            if idx + 1 < len(lines):
                combined = (line + ' ' + lines[idx + 1]).lower()
                if option.lower() in combined:
                    result["Billed To"] = option
                    found_billed_to = True
                    break
        if found_billed_to:
            break

    # Purchase Order: look for PO- pattern (prefer 'PO Number' line, else any PO-)
    for line in lines:
        if 'PO Number' in line:
            po_match = re.search(r'PO\s*-\s*([\w-]+)', line, re.IGNORECASE)
            if po_match:
                result["Purchase Order"] = f"PO-{po_match.group(1)}"
                break
    if not result["Purchase Order"]:
        po_match = re.search(r'PO\s*-\s*([\w-]+)', text, re.IGNORECASE)
        if po_match:
            result["Purchase Order"] = f"PO-{po_match.group(1)}"

    # Invoice Amount: find all 'Total' lines and use the highest value (should include VAT)
    total_amounts = []
    for line in lines:
        if line.strip().lower().startswith("total"):
            amt_match = re.search(r"([\d,.]+)", line)
            if amt_match:
                # Remove commas for float conversion
                amt_val = float(amt_match.group(1).replace(",", ""))
                total_amounts.append((amt_val, amt_match.group(1)))
    if total_amounts:
        # Use the highest total (should be with VAT)
        result["Invoice Amount"] = max(total_amounts, key=lambda x: x[0])[1]

    # IBAN and ACCT (extract full IBAN, clean spaces, stop at non-IBAN word)
    iban_match = re.search(r"IBAN[:# ]*[:\s]*([A-Z0-9 ]{10,})", text, re.IGNORECASE)
    acct_match = re.search(r"ACCT[:# ]*[:\s]*([0-9 ]+)", text, re.IGNORECASE)
    if iban_match:
        # Remove everything after a non-IBAN character (like a country name)
        iban_raw = iban_match.group(1)
        iban_clean = re.match(r"([A-Z]{2}\d{2}[A-Z0-9 ]+)", iban_raw)
        if iban_clean:
            result["IBAN #"] = iban_clean.group(1).replace(" ", "")
        else:
            result["IBAN #"] = iban_raw.replace(" ", "")
    if acct_match:
        acct_num = acct_match.group(1).replace(" ", "")
        result["ACCT #"] = acct_num
    # If IBAN is missing, use ACCT as IBAN fallback
    if not result["IBAN #"] and result["ACCT #"]:
        result["IBAN #"] = result["ACCT #"]


    # Look for lines with multiple values (e.g., 'TotalAmount DueDate InvoiceNumber' and the next line)
    for idx, line in enumerate(lines):
        if re.sub(r"\s+", "", line).lower() == "totalamountduedateinvoicenumber":
            if idx + 1 < len(lines):
                values = re.split(r"\s+", lines[idx + 1].strip())
                # Try to match 3 values: amount, due date, invoice number
                if len(values) >= 3:
                    result["Invoice Amount"] = values[0]
                    result["Due Date"] = values[1]
                    result["Invoice Number"] = values[2]
            break

    # Fallback: Invoice Number, Due Date, Invoice Date from their respective lines if not found
    if not result["Invoice Number"]:
        for line in lines:
            if 'Invoice Number' in line:
                match = re.search(r'Invoice Number[:\s]*([\w-]+)', line, re.IGNORECASE)
                if match:
                    result["Invoice Number"] = match.group(1)
                break

    if not result["Due Date"]:
        for line in lines:
            if 'Due Date' in line:
                match = re.search(r'Due Date[:\s]*([\dA-Z-]+)', line, re.IGNORECASE)
                if match:
                    result["Due Date"] = match.group(1)
                break

    if not result["Invoice Date"]:
        for line in lines:
            if 'Invoice Date' in line:
                match = re.search(r'Invoice Date[:\s]*([\dA-Z-]+)', line, re.IGNORECASE)
                if match:
                    result["Invoice Date"] = match.group(1)
                break

    # Invoice Date: look for 'Invoice Date' line, extract value
    for idx, line in enumerate(lines):
        if 'Invoice Date' in line:
            # Try to extract date from same line
            date_match = re.search(r'Invoice Date[:\s]*([\dA-Z-]+)', line, re.IGNORECASE)
            if date_match:
                result["Invoice Date"] = date_match.group(1)
            # If not, try next line
            elif idx + 1 < len(lines):
                next_line = lines[idx + 1].strip()
                if re.match(r'[\dA-Z-]+', next_line):
                    result["Invoice Date"] = next_line
            break

    # Currency Code: extract any 3-letter currency code from 'Total' line (not just at end)
    currency_candidates = set(["USD", "AED", "KES", "QAR", "OMR", "EUR", "GBP", "SAR", "BHD", "KWD"])
    found_currency = False
    for line in lines:
        if line.strip().lower().startswith("total"):
            # Look for any 3-letter currency code in the line
            matches = re.findall(r"\b([A-Z]{3})\b", line)
            for code in matches:
                if code in currency_candidates:
                    result["Currency Code"] = code
                    found_currency = True
                    break
        if found_currency:
            break
    # Fallback: search whole text for a likely currency code if not found
    if not result["Currency Code"]:
        matches = re.findall(r"\b([A-Z]{3})\b", text)
        for code in matches:
            if code in currency_candidates:
                result["Currency Code"] = code
                break

    # SWIFT Code: extract from 'SWIFT Code:' or 'SWIFT:'
    for line in lines:
        swift_match = re.search(r'SWIFT(?: Code)?[:\s-]*([A-Z0-9]+)', line, re.IGNORECASE)
        if swift_match:
            result["SWIFT Code"] = swift_match.group(1)
            break

    return result

import io

def process_pdfs(pdf_files: List) -> Tuple[pd.DataFrame, dict]:
    data = []
    logs = {}
    timings = []

    def process_single_pdf(pdf_file):
        start = time.time()
        with pdfplumber.open(pdf_file) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()
        fields = extract_fields(text)
        elapsed = time.time() - start
        return pdf_file.name, fields, text, elapsed

    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(executor.map(process_single_pdf, pdf_files))
    for fname, fields, text, elapsed in results:
        logs[fname] = text
        data.append(fields)
        timings.append((fname, elapsed))
    df = pd.DataFrame(data, columns=REQUIRED_FIELDS)
    return df, logs, timings

def show_oracle_tool():
    st.header("Oracle Invoice PDF Extractor")
    st.write("Upload one or more Oracle invoice PDFs to extract key fields and export to Excel.")
    uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)
    if uploaded_files:
        with st.spinner("Processing PDFs..."):
            df, _, timings = process_pdfs(uploaded_files)
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
        st.download_button(
            label="Download Excel",
            data=output,
            file_name="oracle_invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.subheader("PDF Processing Time Summary (seconds)")
        for fname, elapsed in timings:
            st.write(f"{fname}: {elapsed:.2f} seconds")
