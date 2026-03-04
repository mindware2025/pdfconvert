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
    "Invoice Amount",   # document currency amount only
    "Currency Code",    # document currency code only
    "IBAN #",
    "ACCT #",
    "SWIFT Code",
]


def _safe_filename(name: str) -> str:
    return name.replace("/", "_").replace("\\", "_").replace("..", "_")


def write_oracle_log(file_name: str, fields: Dict[str, str], raw_text: str) -> str:
    safe_name = _safe_filename(file_name)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    path = os.path.join(LOG_DIR, f"{ts}_{safe_name}.log")

    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("=== Oracle Invoice Log ===\n")
            f.write(f"File: {file_name}\n")
            f.write(f"Timestamp: {ts}\n\n")

            f.write("--- Extracted Fields ---\n")
            for k in REQUIRED_FIELDS:
                v = fields.get(k, "")
                f.write(f"{k}: {v}\n")

            f.write("\n--- Raw First Page Text ---\n")
            f.write(raw_text or "")
    except Exception as e:
        print(f"[WARN] Failed to write log for {file_name}: {e}")

    return path


# -------------------------------
#  TEXT EXTRACTOR
# -------------------------------
def extract_fields(text: str) -> Dict[str, str]:
    """
    Parse the first page text of an Oracle invoice/credit memo and return the required fields.

    Rules:
      - Always pick the document currency line in the Subtotal/Tax/Total block:
          "Total 16,926.38 USD"
        *or* if "Total" and "<amount> <CUR>" are on separate lines, handle that too.
      - Ignore alternate totals like "Total (AED) 62,162.13".
      - Support Credit Memos:
          - Map "Credit Memo Number" -> Invoice Number
          - Map "Credit Memo Date"   -> Invoice Date
      - Allow negative totals (e.g., "-801,566.14 AED").
      - Do not copy ACCT into IBAN; IBAN only if a real IBAN exists.
      - Only capture PO Number if the next line actually looks like a PO value.
    """
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    result = {k: "" for k in REQUIRED_FIELDS}

    # -----------------------------
    # 0) Helpers
    # -----------------------------
    def _looks_like_label(s: str) -> bool:
        """Heuristic: is this line a field label rather than a value?"""
        labels = {
            "invoice date", "credit memo date", "due date", "invoice number",
            "credit memo number", "po number", "payment terms", "order number",
            "customer vat no.", "agreement", "end user", "subtotal", "tax",
            "total", "payment instructions", "wire funds to", "bill to", "ship to",
            "special instructions", "currency (aed)", "currency (usd)"
        }
        s_low = s.lower()
        if s_low in labels:
            return True
        # short label-like lines with colon
        if (":" in s and len(s) <= 40) and not re.search(r"\S+@\S+", s):
            return True
        return False

    # -----------------------------
    # 1) Stacked header block (works for Invoice or Credit Memo)
    # Pattern (three stacked labels, then three values):
    #   Total Amount
    #   Due Date
    #   Invoice Number | Credit Memo Number
    #   <amount>
    #   <due date>
    #   <doc number>
    # -----------------------------
    for i in range(len(lines) - 5):
        a = lines[i].lower()
        b = lines[i + 1].lower()
        c = lines[i + 2].lower()
        if a == "total amount" and b == "due date" and c in ("invoice number", "credit memo number"):
            result["Invoice Amount"] = lines[i + 3]              # preliminary; may be negative
            result["Due Date"] = lines[i + 4]
            result["Invoice Number"] = lines[i + 5]
            break

    # -----------------------------
    # 2) Invoice/Credit Memo Date + PO Number (validated)
    # -----------------------------
    for i, ln in enumerate(lines):
        key = ln.lower()

        # Invoice Date or Credit Memo Date -> Invoice Date
        if key in ("invoice date", "credit memo date"):
            if i + 1 < len(lines) and not _looks_like_label(lines[i + 1]):
                result["Invoice Date"] = lines[i + 1]

        # PO Number (only if next line looks like a PO value)
        if key == "po number":
            if i + 1 < len(lines):
                nxt = lines[i + 1]
                # Accept "PO - XXXXX" or any non-label token (avoid capturing labels like "Credit Memo Date")
                m = re.search(r"PO\s*-\s*([\w-]+)", nxt, re.IGNORECASE)
                if m:
                    result["Purchase Order"] = f"PO-{m.group(1)}"
                elif not _looks_like_label(nxt):
                    # Fallback: if it's clearly not a label, accept the line as-is
                    result["Purchase Order"] = nxt

    # Extra fallback: scan whole text for "PO - XXXXX"
    if not result["Purchase Order"]:
        m = re.search(r"PO\s*-\s*([\w-]+)", text or "", re.IGNORECASE)
        if m:
            result["Purchase Order"] = f"PO-{m.group(1)}"

    # -----------------------------
    # 3) Totals (DOCUMENT CURRENCY ONLY)
    # Prefer the Subtotal region and handle line breaks and negatives.
    # Ignore any "Total (CUR) ..." alternates.
    # -----------------------------
    def _to_float(s: str) -> float:
        return float(s.replace(",", ""))

    doc_total_amount = None
    doc_total_currency = None

    # parse "<amount> <CUR>" with optional leading '-'
    def _parse_amount_currency(s: str):
        m = re.search(r"(-?[\d,]+\.\d+)\s+([A-Z]{3})\b", s)
        if m:
            try:
                return _to_float(m.group(1)), m.group(2)
            except ValueError:
                pass
        return None, None

    def _scan_window(start_idx: int, max_ahead: int) -> bool:
        """Look for:
           A) 'Total <amount> <CUR>' on same line (amount can be negative)
           B) 'Total' then '<amount> <CUR>' on next 1–2 lines
        """
        nonlocal doc_total_amount, doc_total_currency
        end_idx = min(start_idx + max_ahead, len(lines))
        i = start_idx
        while i < end_idx:
            ln = lines[i]

            # Skip alternates like "Total (AED) 62,162.13"
            if re.search(r"\bTotal\s*\([A-Z]{3}\)", ln):
                i += 1
                continue

            # Same line
            m_same = re.search(r"\bTotal\b[^\d-]*(-?[\d,]+\.\d+)\s+([A-Z]{3})\b", ln)
            if m_same:
                try:
                    doc_total_amount = _to_float(m_same.group(1))
                    doc_total_currency = m_same.group(2)
                    return True
                except ValueError:
                    pass

            # Split across lines
            low = ln.lower()
            if low == "total" or low.startswith("total "):
                for k in (i + 1, i + 2):
                    if k < end_idx:
                        amt, cur = _parse_amount_currency(lines[k])
                        if amt is not None and cur is not None:
                            doc_total_amount = amt
                            doc_total_currency = cur
                            return True
            i += 1
        return False

    sub_idx = next((idx for idx, l in enumerate(lines) if l.lower().startswith("subtotal")), None)
    found = False
    if sub_idx is not None:
        found = _scan_window(sub_idx, max_ahead=12)
    if not found:
        _scan_window(0, max_ahead=len(lines))

    # Finalize
    if doc_total_amount is not None:
        result["Invoice Amount"] = f"{doc_total_amount:,.2f}"
    if doc_total_currency:
        result["Currency Code"] = doc_total_currency

    # -----------------------------
    # 4) Billed To
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
    # 5) Bank details (IBAN / ACCT / SWIFT)
    # -----------------------------
    compact = (text or "").replace(" ", "")
    iban_match = re.search(r"\b([A-Z]{2}\d{2}[A-Z0-9]{11,30})\b", compact, re.IGNORECASE)
    if iban_match:
        result["IBAN #"] = iban_match.group(1).upper()

    m = re.search(r"ACCT[:# ]*[:\s]*([0-9 ]+)", text or "", re.IGNORECASE)
    if m:
        result["ACCT #"] = m.group(1).replace(" ", "")

    for ln in lines:
        m = re.search(r"SWIFT(?: Code)?[:\s-]*([A-Z0-9]+)", ln, re.IGNORECASE)
        if m:
            result["SWIFT Code"] = m.group(1).upper()
            break

    return result


# -------------------------------
# PDF Handling
# -------------------------------
def _extract_text_first_page(pdf_bytes: bytes) -> str:
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            if doc.page_count == 0:
                return ""
            return doc.load_page(0).get_text("text") or ""
    except Exception as e:
        print(f"[ERROR] Failed to read PDF bytes: {e}")
        return ""


def _process_single_pdf(name: str, pdf_bytes: bytes):
    text = _extract_text_first_page(pdf_bytes)
    fields = extract_fields(text)
    write_oracle_log(name, fields, text)
    return fields, text, name


def process_oracle_pdfs_cached(files: List[Tuple[str, bytes]]):
    """
    files: list of (file_name, pdf_bytes)
    returns: (df, text_map)
      - df: DataFrame with REQUIRED_FIELDS as columns
      - text_map: {file_name: raw_first_page_text}
    """
    if not files:
        return pd.DataFrame(columns=REQUIRED_FIELDS), {}

    data: List[Dict[str, str]] = []
    text_map: Dict[str, str] = {}

    with ThreadPoolExecutor(max_workers=min(8, len(files))) as pool:
        futures = [pool.submit(_process_single_pdf, name, blob) for name, blob in files]
        for future in as_completed(futures):
            try:
                fields, raw, fname = future.result()
                data.append(fields)
                text_map[fname] = raw
            except Exception as e:
                print(f"[ERROR] Processing failed for one PDF: {e}")

    df = pd.DataFrame(data, columns=REQUIRED_FIELDS)
    return df, text_map


def prepare_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Writes the DataFrame to an in-memory .xlsx using openpyxl.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf.getvalue()