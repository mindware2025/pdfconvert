# extractors/lenovo_credit_note.py
"""
Lenovo Credit Note Extractor
- Parses Lenovo credit note PDFs
- Emits two rows per PDF (D and C) for CN upload format
"""

import io
import re
from datetime import datetime
from typing import Dict, List, Tuple

import fitz  # PyMuPDF
import pandas as pd

# === Output header (exactly as in CN EXCEL sample) ===
CN_HEADERS = [
    "Page", "Doc No", "Doc Dt", "Seq No", "Ref Seq No", "Manual Entry Y/N",
    "Main A/C", "Sub A/C", "Div", "Dept", "Anly1", "Anly2", "Acty1", "Acty2",
    "Currency", "FC Amt", "LC Amt", "Dr/Cr", "Detail Narration", "Header Narration",
    "Paym Mode", "Chq Book Id", "Chq No", "Chq Dt", "Payee Name", "Val Date",
    "Doc Ref", "TH Doc ref", "Due Dt",
    # FLEX_01 ... FLEX_50 (leave blank)
] + [f"FLEX_{i:02d}" for i in range(1, 51)] + [
    "Party Code", "NOP/NOR", "Tax Code", "Expense Code", "DISC Code",
    # TH_FLEX_01 ... TH_FLEX_50 (blank)
] + [f"TH_FLEX_{i:02d}" for i in range(1, 51)]

# === Constants / defaults per business rule ===
SUB_AC = "SDIL006"
DIV = "PUHO"
DEPT = "GEN"
NARR_PREFIX = "LENOVO(PCG) SELLOUT REBATE RECEIPT / AGREEMENT # "
MAIN_AC_D = "14301"   # when Dr/Cr = D
MAIN_AC_C = "12741"   # when Dr/Cr = C

def _read_pdf_text_first_two_pages(pdf_bytes: bytes) -> str:
    """Read first two pages as text (Lenovo CN generally fits there)."""
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        out = []
        for i in range(min(doc.page_count, 2)):
            out.append(doc.load_page(i).get_text("text") or "")
        return "\n".join(out)

def _normalize_date(d: str) -> str:
    """Convert dates like 12-FEB-2026 to dd/mm/yyyy."""
    try:
        return datetime.strptime(d.strip(), "%d-%b-%Y").strftime("%d/%m/%Y")
    except Exception:
        # Try other variants if needed (uppercasing month)
        try:
            return datetime.strptime(d.strip().upper(), "%d-%b-%Y").strftime("%d/%m/%Y")
        except Exception:
            return d  # return as-is if parse fails

def _extract_fields(text: str) -> Dict[str, str]:
    """
    Pull: Credit No, Credit Date, Currency, Total Amount, Program/Agreement ID
    """
    fields = {
        "credit_no": "",
        "credit_date": "",
        "currency": "",
        "total_amount": "",
        "program_id": "",
    }

    compact = " ".join(text.split())

    # Credit No / Credit Date (samples show 'Credit No. :6415169196   Credit Date :12-FEB-2026')
    m = re.search(r"Credit\s*No\.?\s*:\s*([A-Z0-9/-]+)\s+Credit\s*Date\s*:\s*([0-9]{2}-[A-Za-z]{3}-[0-9]{4})", compact, re.IGNORECASE)
    if m:
        fields["credit_no"] = m.group(1).strip()
        fields["credit_date"] = _normalize_date(m.group(2))

    # Currency and Total — lines show "Total of Products/Services USD 126.00" or "USD 126.00"
    m = re.search(r"Total\s+of\s+Products/Services\s+([A-Z]{3})\s+([0-9,]+\.\d{2})", compact, re.IGNORECASE)
    if m:
        fields["currency"] = m.group(1)
        fields["total_amount"] = m.group(2).replace(",", "")
    else:
        # Fallback: "USD 126.00" near "Total" or "Summary"
        m = re.search(r"(?:Total|Summary)[:\s]+([A-Z]{3})\s+([0-9,]+\.\d{2})", compact, re.IGNORECASE)
        if m:
            fields["currency"] = m.group(1)
            fields["total_amount"] = m.group(2).replace(",", "")

    # Program/Agreement — appears as Program ID: SM-xxxxxx, also in line like "Claim ref ID: SM-10088208"
    m = re.search(r"\bProgram\s*ID\s*:\s*(SM-[0-9-]+)", compact, re.IGNORECASE)
    if m:
        fields["program_id"] = m.group(1).strip()
    else:
        m = re.search(r"Claim\s*ref\s*ID\s*:\s*(SM-[0-9-]+)", compact, re.IGNORECASE)
        if m:
            fields["program_id"] = m.group(1).strip()

    return fields

def _build_two_rows(fields: Dict[str, str], seq: int) -> List[List]:
    """
    Build the two ledger rows (D & C) in the exact header order.
    - LC Amt mirrors FC Amt (USD base).
    - All FLEX and TH_FLEX left blank, also Party/Tax/Expense/DISC blank unless you request values.
    """
    doc_dt = fields["credit_date"]
    doc_ref = fields["credit_no"]
    currency = fields["currency"] or "USD"
    amt = fields["total_amount"] or "0.00"
    narration = f"{NARR_PREFIX}{fields['program_id']}".strip()

    def base_row(main_ac: str, drcr: str, seq_no: int) -> List:
        # Columns up to Due Dt (29 columns)
        row = [
            1,               # Page
            "",              # Doc No (left blank for system to assign, unless you want a value)
            doc_dt,          # Doc Dt
            seq_no,          # Seq No
            seq_no,          # Ref Seq No (mirror)
            "",              # Manual Entry Y/N
            main_ac,         # Main A/C
            SUB_AC,          # Sub A/C
            DIV,             # Div
            DEPT,            # Dept
            "", "", "", "",  # Anly1, Anly2, Acty1, Acty2
            currency,        # Currency
            amt,             # FC Amt
            "",             # LC Amt (per confirmation)
            drcr,            # Dr/Cr
            narration,       # Detail Narration
            narration,       # Header Narration
            "", "", "",      # Paym Mode, Chq Book Id, Chq No
            doc_dt,          # Chq Dt
            "",              # Payee Name
            doc_dt,          # Val Date
            doc_ref,         # Doc Ref
            doc_ref,         # TH Doc ref
            doc_dt,          # Due Dt
        ]
        # FLEX_01 ... FLEX_50 (blank)
        row += [""] * 50
        # Party, NOP/NOR, Tax, Expense, DISC (blank)
        row += ["", "", "", "", ""]
        # TH_FLEX_01 ... TH_FLEX_50 (blank)
        row += [""] * 50
        return row

    row_d = base_row(MAIN_AC_D, "D", seq)
    row_c = base_row(MAIN_AC_C, "C", seq)

    return [row_d, row_c]

def process_lenovo_credit_pdfs(files: List[Tuple[str, bytes]]) -> pd.DataFrame:
    """
    Input: list of (filename, pdf_bytes)
    Output: DataFrame with CN_HEADERS and rows = 2 per PDF (D & C)
    """
    all_rows: List[List] = []
    seq = 1
    for name, blob in files:
        text = _read_pdf_text_first_two_pages(blob)
        fields = _extract_fields(text)
        rows = _build_two_rows(fields, seq)
        all_rows.extend(rows)
        seq += 1
    df = pd.DataFrame(all_rows, columns=CN_HEADERS)
    return df

def prepare_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf.getvalue()