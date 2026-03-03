# oracle_tool.py

"""
Oracle Invoice Extraction Engine
- Fast PDF text extraction via PyMuPDF
- Parallel processing
- Automatic logging
"""

import os
import re
import io
from datetime import datetime
from typing import Dict, List, Tuple

import fitz  # PyMuPDF
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

LOG_DIR = "logs/oracle"
os.makedirs(LOG_DIR, exist_ok=True)

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


def write_oracle_log(file_name: str, fields: Dict[str, str], raw_text: str):
    safe_name = file_name.replace("/", "_").replace("\\", "_")
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    path = os.path.join(LOG_DIR, f"{ts}_{safe_name}.log")

    with open(path, "w", encoding="utf-8") as f:
        f.write("=== Oracle Invoice Log ===\n")
        f.write(f"File: {file_name}\n")
        f.write(f"Timestamp: {ts}\n\n")

        f.write("--- Extracted Fields ---\n")
        for k, v in fields.items():
            f.write(f"{k}: {v}\n")

        f.write("\n--- Raw First Page Text ---\n")
        f.write(raw_text)

    return path


# -------------------------------
#  TEXT EXTRACTOR
# -------------------------------
def extract_fields(text: str) -> Dict[str, str]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    result = {k: "" for k in REQUIRED_FIELDS}

    # -----------------------------
    # 1) Stacked header block
    # -----------------------------
    for i in range(len(lines) - 5):
        if (
            lines[i].lower() == "total amount"
            and lines[i+1].lower() == "due date"
            and lines[i+2].lower() == "invoice number"
        ):
            result["Invoice Amount"] = lines[i+3]
            result["Due Date"] = lines[i+4]
            result["Invoice Number"] = lines[i+5]
            break

    # -----------------------------
    # 2) Invoice Date / PO Number
    # -----------------------------
    for i, ln in enumerate(lines):
        key = ln.lower()
        if key == "invoice date" and i+1 < len(lines):
            result["Invoice Date"] = lines[i+1]
        if key == "po number" and i+1 < len(lines):
            po = lines[i+1]
            m = re.search(r"PO\s*-\s*([\w-]+)", po, re.IGNORECASE)
            result["Purchase Order"] = f"PO-{m.group(1)}" if m else po

    # fallback PO
    if not result["Purchase Order"]:
        m = re.search(r"PO\s*-\s*([\w-]+)", text, re.IGNORECASE)
        if m:
            result["Purchase Order"] = f"PO-{m.group(1)}"

    # -----------------------------
    # 3) Highest Total Line
    # -----------------------------
    best_amount = None
    best_index = None

    # only match REAL total lines, not "Total Amount"
    for idx, ln in enumerate(lines):
        if ln.lower() == "total" or ln.lower().startswith("total "):

            # amount on same line
            m = re.search(r"([\d,]+\.\d+)", ln)
            if m:
                val = float(m.group(1).replace(",", ""))
                if best_amount is None or val > best_amount:
                    best_amount = val
                    best_index = idx

            # amount on next line
            if idx + 1 < len(lines):
                m2 = re.search(r"([\d,]+\.\d+)", lines[idx+1])
                if m2:
                    val2 = float(m2.group(1).replace(",", ""))
                    if best_amount is None or val2 > best_amount:
                        best_amount = val2
                        best_index = idx

    if best_amount is not None:
        result["Invoice Amount"] = f"{best_amount:,.2f}"

    # -----------------------------
    # 4) Currency detection
    # -----------------------------
    search_block = ""

    if best_index is not None:
        # include 3 lines: total line + amount line + currency line
        for j in range(best_index, min(best_index + 3, len(lines))):
            search_block += " " + lines[j]

    # DEBUG — THIS MUST APPEAR IN YOUR TERMINAL
    print("DEBUG CURRENCY BLOCK:", search_block)

    codes = re.findall(r"\b[A-Z]{3}\b", search_block)

    for c in ["AED", "USD", "EUR", "QAR", "OMR", "GBP", "SAR", "KWD", "BHD"]:
        if c in codes:
            result["Currency Code"] = c
            break

    # -----------------------------
    # 5) Billed To
    # -----------------------------
    billed_to_options = [
        "Mindware for Computers",
        "Mindware Technology Trading LLC",
        "Mind Ware S.A",
        "Mindware Limited",
        "Mindware SPC",
        "Mindware FZ LLC",
    ]

    for ln in lines:
        for b in billed_to_options:
            if b.lower() in ln.lower():
                result["Billed To"] = b
                break
        if result["Billed To"]:
            break

    # -----------------------------
    # 6) Bank details
    # -----------------------------
    m = re.search(r"IBAN[:# ]*[:\s]*([A-Z0-9 ]+)", text, re.IGNORECASE)
    if m:
        result["IBAN #"] = m.group(1).replace(" ", "")

    m = re.search(r"ACCT[:# ]*[:\s]*([0-9 ]+)", text, re.IGNORECASE)
    if m:
        acc = m.group(1).replace(" ", "")
        result["ACCT #"] = acc
        if not result["IBAN #"]:
            result["IBAN #"] = acc

    for ln in lines:
        m = re.search(r"SWIFT(?: Code)?[:\s-]*([A-Z0-9]+)", ln, re.IGNORECASE)
        if m:
            result["SWIFT Code"] = m.group(1)
            break

    return result

# -------------------------------
# PDF Handling
# -------------------------------
def _extract_text_first_page(pdf_bytes: bytes) -> str:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        if doc.page_count == 0:
            return ""
        return doc.load_page(0).get_text("text") or ""


def _process_single_pdf(name: str, pdf_bytes: bytes):
    text = _extract_text_first_page(pdf_bytes)
    fields = extract_fields(text)
    write_oracle_log(name, fields, text)
    return fields, text, name


def process_oracle_pdfs_cached(files: List[Tuple[str, bytes]]):
    data = []
    text_map = {}

    with ThreadPoolExecutor(max_workers=min(8, len(files))) as pool:
        futures = [pool.submit(_process_single_pdf, name, blob) for name, blob in files]
        for future in as_completed(futures):
            fields, raw, fname = future.result()
            data.append(fields)
            text_map[fname] = raw

    df = pd.DataFrame(data, columns=REQUIRED_FIELDS)
    return df, text_map


def prepare_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.getvalue()