"""Joint Line invoices - NEW format (Joint Line software change, Aug 2026).

Reads the new Joint Line tax invoices (INV0xxxxx) and customer job credit
notes (CCN0xxxxx) and builds the same consolidated Excel as the old tool
(extractors/joint_line_processor.py, which is kept for the old format).

Rules (same base data as before):
- Base Value for Vat = sum of "Pre Tax Amt" for lines with a tax rate > 0 (SR5%)
- W/o Vat            = sum of "Pre Tax Amt" for lines with 0% (DS0% disbursements)
- Narration          = the charge line descriptions, comma separated
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import pdfplumber
from openpyxl import Workbook

from extractors.joint_line_processor import (
    AMOUNT_FORMAT,
    CURRENCY_FORMAT,
    DATE_FORMAT,
    JOB_LABEL,
    OUTPUT_HEADERS,
)

INV_LABEL = "JOINTLINE INV #"
CN_LABEL = "JOINTLINE CREDIT NOTE #"

AMOUNT = r"-?[\d,]+\.\d{2}"
TAX_CODE = r"[A-Z]{1,3}\d+(?:\.\d+)?%"
# "... 150.00 SR5% 7.50 157.50" at the end of an item row
ROW_TAIL = re.compile(rf"(?P<pretax>{AMOUNT})\s+(?P<code>{TAX_CODE})\s+{AMOUNT}\s+{AMOUNT}\s*$")
DESC_X_MIN, DESC_X_MAX = 40, 195  # "Charge Line Item" column on the invoice


def _to_float(value: str) -> float:
    return float(value.replace(",", ""))


def _tax_rate(code: str) -> float:
    match = re.search(r"(\d+(?:\.\d+)?)%", code)
    return float(match.group(1)) if match else 0.0


def _parse_date(raw: str) -> datetime:
    return datetime.strptime(raw.strip(), "%d/%b/%Y")


def _search(pattern: str, text: str) -> str:
    match = re.search(pattern, text, re.MULTILINE)
    return match.group(1).strip() if match else ""


def _line_after(text: str, marker: str) -> str:
    lines = [line.strip() for line in text.splitlines()]
    for i, line in enumerate(lines):
        if line.startswith(marker) and i + 1 < len(lines):
            return lines[i + 1]
    return ""


def _invoice_descriptions(page) -> dict[int, str]:
    """Charge descriptions per line number, from word positions.

    A long description wraps over several lines and the PDF centres it
    vertically on the row, so text can sit above or below the row's number.
    Each word in the description column is given to the nearest row.
    """
    words = page.extract_words(keep_blank_chars=False)
    header = next((w for w in words if w["text"] == "Sr."), None)
    footer = next(
        (w for w in words if w["text"] == "Pre" and header and w["top"] > header["top"] + 5 and w["x0"] > 300),
        None,
    )
    if not header:
        return {}
    top_limit = header["bottom"]
    bottom_limit = footer["top"] if footer else page.height

    zone = [w for w in words if top_limit < w["top"] < bottom_limit]
    rows = [w for w in zone if w["x0"] < DESC_X_MIN and w["text"].isdigit()]
    if not rows:
        return {}

    parts: dict[int, list[tuple[float, float, str]]] = {int(r["text"]): [] for r in rows}
    for w in zone:
        if DESC_X_MIN <= w["x0"] < DESC_X_MAX:
            nearest = min(rows, key=lambda r: abs(r["top"] - w["top"]))
            parts[int(nearest["text"])].append((round(w["top"]), w["x0"], w["text"]))

    return {
        number: " ".join(text for _, _, text in sorted(items))
        for number, items in parts.items()
    }


def _parse_invoice(file_name: str, text: str, page) -> tuple[dict | None, list[str]]:
    errors: list[str] = []
    entity = _line_after(text, "TAX INVOICE").split(" Quote Ref No")[0].strip()
    doc_no = _search(r"Invoice No\s+(INV\d+)", text)
    date_raw = _search(r"Invoice Date\s+(\d{2}/[A-Za-z]{3}/\d{4})", text)
    job_no = _search(r"Job No\s+(\S+)", text)

    descriptions = _invoice_descriptions(page)
    items: list[tuple[str, float, float]] = []
    in_items = False
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("Sr. Charge"):
            in_items = True
            continue
        if not in_items:
            continue
        if line.startswith("Pre Tax Amt"):
            break
        number = re.match(r"^(\d+)\s", line)
        tail = ROW_TAIL.search(line)
        if number and tail:
            n = int(number.group(1))
            items.append((descriptions.get(n, ""), _tax_rate(tail.group("code")), _to_float(tail.group("pretax"))))

    return _finish(file_name, "invoice", entity, doc_no, date_raw, job_no, items, errors)


def _parse_credit_note(file_name: str, text: str) -> tuple[dict | None, list[str]]:
    errors: list[str] = []
    entity = _line_after(text, "Customer Job Credit Note")
    doc_no = _search(r"Tranx No\s+(CCN\d+)", text)
    date_raw = _search(r"Tranx Date\s+(\d{2}/[A-Za-z]{3}/\d{4})", text)

    items: list[tuple[str, float, float]] = []
    job_no = ""
    in_items = False
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("Sr. Debit Account"):
            in_items = True
            continue
        if not in_items:
            continue
        if line.startswith("Amount ") or line.startswith("Against the settlement"):
            break
        tail = ROW_TAIL.search(line)
        if re.match(r"^\d+\s", line) and tail:
            job = re.search(r"(\S+)\s+" + AMOUNT + r"\s+" + TAX_CODE, line)
            if job and not job_no:
                job_no = job.group(1)
            # Narration is left blank for credit notes, as in Finance's template.
            items.append(("", _tax_rate(tail.group("code")), _to_float(tail.group("pretax"))))

    return _finish(file_name, "credit_note", entity, doc_no, date_raw, job_no, items, errors)


def _finish(file_name, doc_type, entity, doc_no, date_raw, job_no, items, errors):
    if not entity:
        errors.append(f"{file_name}: could not find customer/entity name.")
    if not doc_no:
        errors.append(f"{file_name}: could not find invoice / credit note number.")
    if not date_raw:
        errors.append(f"{file_name}: could not find the date.")
    if not job_no:
        errors.append(f"{file_name}: could not find job number.")
    if not items:
        errors.append(f"{file_name}: no charge lines found.")
    if errors:
        return None, errors

    return {
        "file_name": file_name,
        "doc_type": doc_type,
        "invoice_no": doc_no,
        "invoice_date": _parse_date(date_raw),
        "job_no": job_no,
        "base_value_vat": round(sum(amt for _, rate, amt in items if rate > 0), 2),
        "wo_vat": round(sum(amt for _, rate, amt in items if rate == 0), 2),
        "narration": ", ".join(desc for desc, _, _ in items if desc),
        "entity": entity,
    }, []


def extract_document(file_name: str, file_bytes: bytes) -> tuple[dict | None, list[str]]:
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            page = pdf.pages[0]
            text = page.extract_text() or ""
            if "Customer Job Credit Note" in text:
                return _parse_credit_note(file_name, text)
            if "TAX INVOICE" in text:
                return _parse_invoice(file_name, text, page)
    except Exception as exc:  # unreadable / not a PDF
        return None, [f"{file_name}: could not read the PDF ({exc})."]
    return None, [f"{file_name}: not a new-format Joint Line invoice or credit note."]


def process_joint_line_v2_pdfs(uploaded_files: list) -> tuple[list[dict], list[str]]:
    docs: list[dict] = []
    errors: list[str] = []
    for uploaded_file in uploaded_files:
        doc, parse_errors = extract_document(uploaded_file.name, uploaded_file.getvalue())
        if parse_errors:
            errors.extend(parse_errors)
        else:
            docs.append(doc)
    # Invoices first, then credit notes; each by number.
    docs.sort(key=lambda d: (d["doc_type"] != "invoice", d["invoice_no"]))
    return docs, errors


def _label(doc: dict) -> str:
    return CN_LABEL if doc["doc_type"] == "credit_note" else INV_LABEL


def build_preview_dataframe(docs: list[dict]) -> pd.DataFrame:
    rows = []
    for doc in docs:
        vat = round(doc["base_value_vat"] * 0.05, 2)
        rows.append(
            {
                "INV NUMBER": doc["invoice_no"],
                "INV DATE": doc["invoice_date"].date(),
                "Job No": doc["job_no"],
                "Base Value for Vat": doc["base_value_vat"],
                "W/o Vat": doc["wo_vat"],
                "NARRATION": doc["narration"],
                "Due date": (doc["invoice_date"] + timedelta(days=30)).date(),
                "Vat": vat,
                "Inv ": round(vat + doc["wo_vat"] + doc["base_value_vat"], 2),
                "Description of Charges": f'{_label(doc)}{doc["invoice_no"]} {doc["narration"]}  {JOB_LABEL}{doc["job_no"]}',
                "Entity": doc["entity"],
            }
        )
    return pd.DataFrame(rows)


def create_joint_line_v2_excel_file(docs: list[dict]) -> BytesIO:
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(OUTPUT_HEADERS)

    for idx, doc in enumerate(docs, start=2):
        ws.cell(row=idx, column=1, value=doc["invoice_no"])
        ws.cell(row=idx, column=2, value=doc["invoice_date"]).number_format = DATE_FORMAT
        ws.cell(row=idx, column=3, value=doc["job_no"])
        ws.cell(row=idx, column=4, value=doc["base_value_vat"])
        ws.cell(row=idx, column=5, value=doc["wo_vat"]).number_format = CURRENCY_FORMAT
        ws.cell(row=idx, column=6, value=doc["narration"])
        ws.cell(row=idx, column=7, value=f"=B{idx}+30").number_format = DATE_FORMAT
        ws.cell(row=idx, column=9, value=f"=D{idx}*5%").number_format = AMOUNT_FORMAT
        ws.cell(row=idx, column=10, value=f"=I{idx}+E{idx}+D{idx}").number_format = AMOUNT_FORMAT
        ws.cell(row=idx, column=12, value=_label(doc))
        ws.cell(row=idx, column=13, value=JOB_LABEL)
        ws.cell(row=idx, column=15, value=f'=L{idx}&A{idx}&" "&F{idx}&" "&N{idx}&" "&M{idx}&C{idx}')
        ws.cell(row=idx, column=16, value=doc["entity"])

    for column_cells in ws.columns:
        max_length = max((len(str(cell.value or "")) for cell in column_cells), default=0)
        ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 60)

    wb.save(output)
    output.seek(0)
    return output
