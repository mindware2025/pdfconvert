from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re
from typing import BinaryIO

import pandas as pd
from openpyxl import Workbook
from PyPDF2 import PdfReader


OUTPUT_HEADERS = [
    "Doc No",
    "Doc Dt",
    "Seq No ",
    "Ref Seq No",
    "Manual Entry Y/N",
    "Main A/C",
    "Sub A/C",
    "Div",
    "Dept",
    "Anly1",
    "Anly2",
    "Acty1",
    "Acty2",
    "Currency",
    "FC Amt",
    "LC Amt",
    "Dr/Cr",
    "Detail Narration",
    "Header Narration",
    "Paym Mode",
    "Chq Book Id",
    "Chq No",
    "Chq Dt",
    "Payee Name",
    "Val Date",
    "Doc Ref",
    "TH Doc ref",
    "Due Dt",
    *[f"FLEX_{idx:02d}" for idx in range(1, 51)],
    "Party Code",
    "NOP/NOR",
    "Tax Code",
    "Expense Code",
    "DISC Code",
    *[f"TH_FLEX_{idx:02d}" for idx in range(1, 51)],
]


@dataclass
class JVConfig:
    debit_main_account: str = "14779"
    credit_main_account: str = "14310"
    credit_sub_account: str = "SDFE001"
    division: str = "PUHO"
    department: str = "GEN"
    manual_entry_flag: str = ""
    payment_mode: str = ""
    payee_name: str = ""
    party_code: str = ""
    nop_nor: str = ""
    tax_code: str = ""
    expense_code: str = ""
    discount_code: str = ""
    narration_suffix: str = ""
    due_days: int = 60


def _extract_text(file_obj: BinaryIO) -> str:
    reader = PdfReader(file_obj)
    if not reader.pages:
        return ""
    return reader.pages[0].extract_text() or ""


def _search(pattern: str, text: str, flags: int = 0) -> str:
    match = re.search(pattern, text, flags)
    return match.group(1).strip() if match else ""


def _parse_amount(value: str) -> Decimal:
    cleaned = value.replace(",", "").strip()
    return Decimal(cleaned)


def _format_amount(value: Decimal) -> str:
    return f"{value:,.2f}"


def _parse_date(date_text: str) -> datetime.date | None:
    if not date_text:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_text.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _normalize_invoice_number(value: str) -> str:
    if not value:
        return ""
    match = re.match(r"[A-Z0-9-]+", value.strip())
    return match.group(0) if match else value.strip()


def _format_date(date_value: datetime.date | None) -> str:
    return date_value.strftime("%d/%m/%Y") if date_value else ""


def _build_narration(invoice_no: str, awb_no: str, hawb_no: str, suffix: str) -> str:
    parts = [f"EXPEDITORS INV # {invoice_no}"]
    parts.append("SHIPMENT CHARGE")
    if awb_no:
        parts.append(f"AWB # {awb_no}")
    if hawb_no:
        parts.append(f"HAWB # {hawb_no}")
    if suffix.strip():
        parts.append(suffix.strip())
    return " / ".join(parts)


def extract_invoice_data(file_name: str, file_bytes: bytes) -> tuple[dict, list[str]]:
    errors: list[str] = []
    pdf_text = _extract_text(BytesIO(file_bytes))

    invoice_match = re.search(
        r"(?i:INVOICE\s*NUMBER)\s*([A-Z0-9-]+?)(?=(?:[A-Z][a-z])|\s|$)",
        pdf_text,
    )
    if not invoice_match:
        invoice_match = re.search(
            r"(?i:INVOICE\s*NO\.?)\s*([A-Z0-9-]+?)(?=(?:[A-Z][a-z])|\s|$)",
            pdf_text,
        )
    invoice_no = _normalize_invoice_number(invoice_match.group(1) if invoice_match else "")

    invoice_date = _parse_date(_search(r"INVOICE\s*DATE\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", pdf_text, re.IGNORECASE))
    awb_no = _search(r"AWB/BL\s+([A-Z0-9-]+)", pdf_text, re.IGNORECASE)
    hawb_no = _search(r"HAWB/HBL:\s*([A-Z0-9-]+)", pdf_text, re.IGNORECASE)
    payee_name = _search(r"NAME:\s*(.+?)\s+INVOICE", pdf_text, re.IGNORECASE | re.DOTALL)
    if payee_name:
        payee_name = " ".join(payee_name.split())

    total_match = re.search(
        r"INVOICE\s*TOTAL\s*:?\s*([0-9,]+\.\d{2})\s*([A-Z]{3})",
        pdf_text,
        re.IGNORECASE,
    )
    if not total_match:
        total_match = re.search(
            r"([A-Z]{3})\s*[\r\n ]+INVOICE\s*TOTAL\s*:?\s*([0-9,]+\.\d{2})",
            pdf_text,
            re.IGNORECASE,
        )

    currency = ""
    amount = Decimal("0.00")
    if total_match:
        if total_match.lastindex == 2 and total_match.group(1).isalpha():
            if len(total_match.group(1)) == 3 and total_match.group(1).isupper():
                currency = total_match.group(2).upper()
                amount = _parse_amount(total_match.group(1))
        if not currency:
            amount = _parse_amount(total_match.group(1))
            currency = total_match.group(2).upper()
    else:
        errors.append(f"{file_name}: could not find invoice total and currency.")

    if not invoice_no:
        errors.append(f"{file_name}: could not find invoice number.")
    if not invoice_date:
        errors.append(f"{file_name}: could not find invoice date.")

    return {
        "file_name": file_name,
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "awb_no": awb_no,
        "hawb_no": hawb_no,
        "payee_name": payee_name,
        "currency": currency,
        "amount": amount,
        "text": pdf_text,
    }, errors


def build_jv_rows(invoice_data: dict, doc_no: int, config: JVConfig) -> list[dict]:
    today = datetime.today().date()
    doc_date = today
    due_date = doc_date + timedelta(days=config.due_days) if doc_date else None
    narration = _build_narration(
        invoice_data["invoice_no"],
        invoice_data["awb_no"],
        invoice_data["hawb_no"],
        config.narration_suffix,
    )
    amount = _format_amount(invoice_data["amount"])
    base_row = {header: "" for header in OUTPUT_HEADERS}
    base_row.update(
        {
            "Doc No": str(doc_no),
            "Doc Dt": _format_date(doc_date),
            "Seq No ": str(doc_no),
            "Ref Seq No": "",
            "Manual Entry Y/N": config.manual_entry_flag,
            "Div": config.division,
            "Dept": config.department,
            "Currency": invoice_data["currency"],
            "FC Amt": amount,
            "LC Amt": "",
            "Detail Narration": narration,
            "Header Narration": narration,
            "Paym Mode": config.payment_mode,
            "Payee Name": config.payee_name,
            "Val Date": _format_date(doc_date),
            "Doc Ref": invoice_data["invoice_no"],
            "TH Doc ref": invoice_data["invoice_no"],
            "Due Dt": _format_date(due_date),
            "Party Code": config.party_code,
            "NOP/NOR": config.nop_nor,
            "Tax Code": config.tax_code,
            "Expense Code": config.expense_code,
            "DISC Code": config.discount_code,
        }
    )

    debit_row = dict(base_row)
    debit_row["Main A/C"] = config.debit_main_account
    debit_row["Dr/Cr"] = "D"

    credit_row = dict(base_row)
    credit_row["Main A/C"] = config.credit_main_account
    credit_row["Sub A/C"] = config.credit_sub_account
    credit_row["Dr/Cr"] = "C"

    return [debit_row, credit_row]


def process_freight_forwarder_pdfs(uploaded_files: list, config: JVConfig) -> tuple[pd.DataFrame, list[dict], list[str]]:
    all_rows: list[dict] = []
    parsed_invoices: list[dict] = []
    errors: list[str] = []

    for doc_no, uploaded_file in enumerate(uploaded_files, start=1):
        file_bytes = uploaded_file.getvalue()
        invoice_data, parse_errors = extract_invoice_data(uploaded_file.name, file_bytes)
        parsed_invoices.append(invoice_data)
        errors.extend(parse_errors)

        if parse_errors:
            continue

        all_rows.extend(build_jv_rows(invoice_data, doc_no, config))

    df = pd.DataFrame(all_rows, columns=OUTPUT_HEADERS)
    return df, parsed_invoices, errors


def create_excel_file(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "JV Upload"

    ws.append(OUTPUT_HEADERS)
    for _, row in df.iterrows():
        ws.append([row.get(col, "") for col in OUTPUT_HEADERS])

    for column_cells in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 40)

    wb.save(output)
    output.seek(0)
    return output
