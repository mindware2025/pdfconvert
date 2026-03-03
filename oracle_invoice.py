# oracle_tool.py
"""
Oracle Tool: PDF Invoice Extraction and Excel Export (PyMuPDF version)
- Fast first-page text extraction with PyMuPDF (fitz)
- Parallel processing for multiple PDFs
- Streamlit cache to avoid reprocessing on reruns (e.g., download clicks)
"""

from __future__ import annotations

import io
import re
from typing import List, Tuple, Dict

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from concurrent.futures import ThreadPoolExecutor, as_completed


# ----------------------------
# Fields and Parsing Utilities
# ----------------------------

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
    "SWIFT Code",
]


def extract_fields(text: str) -> Dict[str, str]:
    """
    Extract key invoice fields from the first-page text of an Oracle invoice PDF.
    This is your original logic, with minor robustness tweaks and clean Python operators.
    """
    lines = text.splitlines()
    result = {field: "" for field in REQUIRED_FIELDS}

    # ----- Billed To -----
    billed_to_options = [
        "Mindware for Computers",
        "Mindware Technology Trading LLC",
        "Mind Ware S.A",
        "Mindware Limited",
        "Mindware SPC",
        "Mindware FZ LLC",
    ]
    found_billed_to = False
    for idx, line in enumerate(lines):
        if "Reference Invoice Number:" in line:
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
        for option in billed_to_options:
            if option.lower() in line.lower():
                result["Billed To"] = option
                found_billed_to = True
                break
            if idx + 1 < len(lines):
                combined = (line + " " + lines[idx + 1]).lower()
                if option.lower() in combined:
                    result["Billed To"] = option
                    found_billed_to = True
                    break
        if found_billed_to:
            break

    # ----- Purchase Order -----
    for line in lines:
        if "PO Number" in line:
            po_match = re.search(r"PO\s*-\s*([\w-]+)", line, re.IGNORECASE)
            if po_match:
                result["Purchase Order"] = f"PO-{po_match.group(1)}"
                break
    if not result["Purchase Order"]:
        po_match = re.search(r"PO\s*-\s*([\w-]+)", text, re.IGNORECASE)
        if po_match:
            result["Purchase Order"] = f"PO-{po_match.group(1)}"

    # ----- Invoice Amount (pick highest 'Total' line) -----
    total_amounts = []
    for line in lines:
        if line.strip().lower().startswith("total"):
            amt_match = re.search(r"([\d.,]+)", line)
            if amt_match:
                try:
                    amt_val = float(amt_match.group(1).replace(",", ""))
                    total_amounts.append((amt_val, amt_match.group(1)))
                except ValueError:
                    pass
    if total_amounts:
        result["Invoice Amount"] = max(total_amounts, key=lambda x: x[0])[1]

    # ----- IBAN and ACCT -----
    iban_match = re.search(r"IBAN[:# ]*[:\s]*([A-Z0-9 ]{10,})", text, re.IGNORECASE)
    acct_match = re.search(r"ACCT[:# ]*[:\s]*([0-9 ]+)", text, re.IGNORECASE)
    if iban_match:
        iban_raw = iban_match.group(1)
        iban_clean = re.match(r"([A-Z]{2}\d{2}[A-Z0-9 ]+)", iban_raw)
        if iban_clean:
            result["IBAN #"] = iban_clean.group(1).replace(" ", "")
        else:
            result["IBAN #"] = iban_raw.replace(" ", "")
    if acct_match:
        acct_num = acct_match.group(1).replace(" ", "")
        result["ACCT #"] = acct_num
    if not result["IBAN #"] and result["ACCT #"]:
        result["IBAN #"] = result["ACCT #"]

    # ----- Multi-value header row ('TotalAmount DueDate InvoiceNumber') -----
    for idx, line in enumerate(lines):
        if re.sub(r"\s+", "", line).lower() == "totalamountduedateinvoicenumber":
            if idx + 1 < len(lines):
                values = re.split(r"\s+", lines[idx + 1].strip())
                if len(values) >= 3:
                    result["Invoice Amount"] = values[0]
                    result["Due Date"] = values[1]
                    result["Invoice Number"] = values[2]
            break

    # ----- Fallbacks: Invoice Number, Due Date, Invoice Date -----
    if not result["Invoice Number"]:
        for line in lines:
            if "Invoice Number" in line:
                match = re.search(r"Invoice Number[:\s]*([\w-]+)", line, re.IGNORECASE)
                if match:
                    result["Invoice Number"] = match.group(1)
                break

    if not result["Due Date"]:
        for line in lines:
            if "Due Date" in line:
                match = re.search(r"Due Date[:\s]*([\dA-Z-]+)", line, re.IGNORECASE)
                if match:
                    result["Due Date"] = match.group(1)
                break

    if not result["Invoice Date"]:
        for line in lines:
            if "Invoice Date" in line:
                match = re.search(r"Invoice Date[:\s]*([\dA-Z-]+)", line, re.IGNORECASE)
                if match:
                    result["Invoice Date"] = match.group(1)
                break

    # Explicit second pass for 'Invoice Date' (next line case)
    for idx, line in enumerate(lines):
        if "Invoice Date" in line:
            date_match = re.search(r"Invoice Date[:\s]*([\dA-Z-]+)", line, re.IGNORECASE)
            if date_match:
                result["Invoice Date"] = date_match.group(1)
            elif idx + 1 < len(lines):
                next_line = lines[idx + 1].strip()
                if re.match(r"[\dA-Z-]+", next_line):
                    result["Invoice Date"] = next_line
            break

    # ----- Currency Code -----
    currency_candidates = {"USD", "AED", "KES", "QAR", "OMR", "EUR", "GBP", "SAR", "BHD", "KWD"}
    found_currency = False
    for line in lines:
        if line.strip().lower().startswith("total"):
            matches = re.findall(r"\b([A-Z]{3})\b", line)
            for code in matches:
                if code in currency_candidates:
                    result["Currency Code"] = code
                    found_currency = True
                    break
        if found_currency:
            break
    if not result["Currency Code"]:
        matches = re.findall(r"\b([A-Z]{3})\b", text)
        for code in matches:
            if code in currency_candidates:
                result["Currency Code"] = code
                break

    # ----- SWIFT Code -----
    for line in lines:
        swift_match = re.search(r"SWIFT(?: Code)?[:\s-]*([A-Z0-9]+)", line, re.IGNORECASE)
        if swift_match:
            result["SWIFT Code"] = swift_match.group(1)
            break

    return result


# ----------------------------
# Extraction + Caching
# ----------------------------

def _extract_text_first_page_pymupdf(pdf_bytes: bytes) -> str:
    """
    Return first-page text using PyMuPDF (fast).
    """
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        if doc.page_count == 0:
            return ""
        page = doc.load_page(0)
        return page.get_text("text") or ""


def _process_single_pdf(name: str, pdf_bytes: bytes) -> Tuple[Dict[str, str], str, str]:
    """
    Process a single PDF: extract first-page text, parse fields.
    Returns: (fields_dict, raw_text, filename)
    """
    text = _extract_text_first_page_pymupdf(pdf_bytes)
    fields = extract_fields(text)
    return fields, text, name


@st.cache_data(show_spinner=False, max_entries=10)
def process_oracle_pdfs_cached(file_blobs: List[Tuple[str, bytes]]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Cached processing for Oracle PDFs.
    Args:
        file_blobs: list of (filename, file_bytes)
    Returns:
        df: DataFrame with REQUIRED_FIELDS columns
        text_map: {filename: first_page_text}
    """
    data: List[Dict[str, str]] = []
    text_map: Dict[str, str] = {}

    if not file_blobs:
        return pd.DataFrame(columns=REQUIRED_FIELDS), text_map

    # Process in parallel (I/O bound)
    with ThreadPoolExecutor(max_workers=min(8, len(file_blobs))) as ex:
        futures = [ex.submit(_process_single_pdf, name, b) for name, b in file_blobs]

        progress, total = 0, len(futures)
        bar = st.progress(0.0, text="Processing Oracle PDFs...")

        for fut in as_completed(futures):
            fields, text, name = fut.result()
            data.append(fields)
            text_map[name] = text
            progress += 1
            bar.progress(progress / max(total, 1), text=f"Processed {progress} of {total} PDFs")

    df = pd.DataFrame(data, columns=REQUIRED_FIELDS)
    return df, text_map


def prepare_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Convert a DataFrame to Excel bytes.
    """
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.getvalue()