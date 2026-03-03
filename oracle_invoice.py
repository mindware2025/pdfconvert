# oracle_tool.py (excerpt)

import io
import hashlib
from typing import List, Tuple, Dict
import pdfplumber
import pandas as pd
import streamlit as st
from concurrent.futures import ThreadPoolExecutor, as_completed

REQUIRED_FIELDS = [
    "Billed To", "Invoice Number", "Invoice Date", "Due Date",
    "Purchase Order", "Invoice Amount", "Currency Code",
    "IBAN #", "ACCT #", "SWIFT Code"
]

def _hash_bytes(b: bytes) -> str:
    # Stable hash to cache processing results per file content
    h = hashlib.sha256()
    h.update(b)
    return h.hexdigest()

def _extract_first_page_text(pdf_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_page = pdf.pages[0]
        return first_page.extract_text() or ""

def _process_single_pdf(pdf_name: str, pdf_bytes: bytes) -> Tuple[Dict[str, str], str, str]:
    """
    Returns: (fields, raw_text, name)
    """
    text = _extract_first_page_text(pdf_bytes)
    fields = extract_fields(text)
    return fields, text, pdf_name

@st.cache_data(show_spinner=False, max_entries=8)
def process_pdfs_cached(file_blobs: List[Tuple[str, bytes]]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Cached processing for a set of PDFs.
    Input is a list of tuples: (filename, file_bytes)
    Returns DataFrame and logs dict[filename] -> first page text
    """
    data = []
    logs: Dict[str, str] = {}

    # Parallel extraction for speed
    # Note: pdfplumber is I/O-bound; threads work well for many small PDFs.
    with ThreadPoolExecutor(max_workers=min(8, len(file_blobs) or 1)) as executor:
        futures = {
            executor.submit(_process_single_pdf, name, b): (name, b)
            for name, b in file_blobs
        }

        progress = 0
        total = len(futures)
        progress_bar = st.progress(0.0, text="Processing PDFs...")

        for future in as_completed(futures):
            fields, text, name = future.result()
            data.append(fields)
            logs[name] = text
            progress += 1
            progress_bar.progress(progress / max(1, total), text=f"Processed {progress} of {total} PDFs")

    df = pd.DataFrame(data, columns=REQUIRED_FIELDS)
    return df, logs

def prepare_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    # OpenPyXL is used implicitly by pandas for .xlsx
    df.to_excel(output, index=False)
    output.seek(0)
    return output.read()