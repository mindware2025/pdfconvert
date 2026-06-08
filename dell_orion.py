from datetime import datetime
from io import BytesIO
import re

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from dell import (
    CURRENCY_CONVERSION_RATES,
    CURRENCY_NUMBER_FORMATS,
    _extract_compact_quote_items_and_config,
    _extract_grouped_template_items_and_config,
    _extract_pdf_quote_data,
    _extract_product_detail_headings,
    _extract_quote_metadata,
    _extract_config_rows_from_configuration_sheet,
    _try_extract_items_from_pricing_summary,
)

ORION_HEADERS = [
    "Vendor Item Code",
    "P/N - Orion Item code",
    "Description",
    "Qty",
    "MSRP",
    "Unit Cost",
    "Unit Selling",
]


def _cell_to_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()


def _sanitize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _extract_part_number_from_description(text: str) -> str:
    text = _sanitize_text(text)
    if not text:
        return ""

    matches = re.findall(r"\(([^()]+)\)", text)
    for candidate in reversed(matches):
        normalized = candidate.strip().replace("–", "-").replace("—", "-")
        normalized = re.sub(r"\s*-\s*", "-", normalized)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        if re.fullmatch(r"(?:\d{3}-[A-Z0-9]{4,5}|[A-Z]{2}\d{6,})", normalized, re.I):
            return normalized

    fallback = re.search(r"\b(?:[A-Z]{2}\d{6,}|\d{3}-[A-Z0-9]{4,5})\b", text, re.I)
    if fallback:
        return fallback.group(0).upper()
    return ""


def _best_description(desc: str, config_rows: list, idx: int) -> str:
    base = _sanitize_text(desc)
    if base:
        return base

    for row in config_rows:
        if len(row) >= 5 and str(row[0]).strip() == str(idx + 1):
            candidate = _sanitize_text(row[3] or row[4] or row[2] or row[1])
            if candidate:
                return candidate
    return base


def _extract_items_and_metadata(input_excel_bytes: bytes):
    is_pdf = input_excel_bytes.lstrip().startswith(b"%PDF")
    if is_pdf:
        items, quote_meta, config_rows, quote_ref_text, date_text, _ = _extract_pdf_quote_data(input_excel_bytes)
        return items, quote_meta, config_rows, {}, quote_ref_text, date_text, is_pdf

    wb = openpyxl.load_workbook(BytesIO(input_excel_bytes), data_only=True)
    ws = wb.active
    quote_meta = _extract_quote_metadata(ws)
    item_headings_by_item = _extract_product_detail_headings(ws)

    items = _try_extract_items_from_pricing_summary(ws)
    config_rows = []
    if not items:
        items, config_rows = _extract_compact_quote_items_and_config(ws)
    if not items:
        items, config_rows = _extract_grouped_template_items_and_config(ws)

    if not items:
        items = []

    config_sheet = None
    for sheet_name in wb.sheetnames:
        if "config" in sheet_name.lower():
            config_sheet = wb[sheet_name]
            break

    if config_sheet is not None:
        config_rows = _extract_config_rows_from_configuration_sheet(config_sheet)

    quote_ref_text = ""
    date_text = ""
    for r in range(1, min(ws.max_row, 60) + 1):
        label = _cell_to_text(ws.cell(r, 2).value).lower()
        if "quote" in label and "ref" in label:
            quote_ref_text = _cell_to_text(ws.cell(r, 5).value)
        if label.startswith("date"):
            date_text = _cell_to_text(ws.cell(r, 5).value)

    return items, quote_meta, config_rows, item_headings_by_item, quote_ref_text, date_text, is_pdf


def build_dell_orion_output_filename(input_excel_bytes: bytes) -> str:
    _, quote_meta, _, _, quote_ref_text, _, _ = _extract_items_and_metadata(input_excel_bytes)
    partner_name = _sanitize_text(
        quote_meta.get("reseller") or quote_meta.get("company name") or quote_meta.get("end user") or "Dell"
    )
    quote_ref = _sanitize_text(quote_ref_text) or "Quote"
    safe_partner = re.sub(r"[^A-Za-z0-9._-]+", " ", partner_name).strip() or "Dell"
    safe_ref = re.sub(r"[^A-Za-z0-9._-]+", " ", quote_ref).strip() or "Quote"
    return f"{safe_partner}-{safe_ref}-{datetime.now().strftime('%Y%m%d')}.xlsx"


def generate_orion_quote(input_excel_bytes: bytes, currency_code: str = "USD") -> bytes:
    """
    Generate a basic Orion quotation workbook from the same Dell quotation input.
    The output is intentionally lightweight and keeps the existing Dell generator untouched.
    """
    items, _, config_rows, item_headings_by_item, _, _, _ = _extract_items_and_metadata(input_excel_bytes)
    conversion_rate = CURRENCY_CONVERSION_RATES.get((currency_code or "USD").upper(), 1.0)
    number_format = CURRENCY_NUMBER_FORMATS.get((currency_code or "USD").upper(), "#,##0.00")

    wb = Workbook()
    ws = wb.active
    ws.title = "Orion_Quote"

    ws.append(ORION_HEADERS)

    item_descs_order = [item[0] for item in items]
    part_numbers_by_item = {}
    for row_data in config_rows:
        if len(row_data) >= 7:
            item_no, _heading, _module, _desc, sku, _tax, _qty = row_data
        else:
            item_no, _heading, _module, _desc, sku, _tax = row_data
        if sku and item_no not in part_numbers_by_item:
            part_numbers_by_item[item_no] = sku

    heading_part_numbers_by_item = {}
    for item_key, heading in item_headings_by_item.items():
        part_number = _extract_part_number_from_description(heading)
        if part_number:
            heading_part_numbers_by_item[item_key] = part_number

    for idx, desc in enumerate(item_descs_order, start=1):
        item_key = str(idx)
        if item_key not in heading_part_numbers_by_item:
            part_number = _extract_part_number_from_description(desc)
            if part_number:
                heading_part_numbers_by_item[item_key] = part_number

    for item_key, part_number in heading_part_numbers_by_item.items():
        part_numbers_by_item.setdefault(item_key, part_number)

    for idx, (desc, qty, unit_price, total_price) in enumerate(items, start=1):
        qty_value = int(qty) if qty not in (None, "") else 0
        unit_value = float(unit_price or 0.0)
        total_value = float(total_price) if total_price is not None else (qty_value * unit_value)
        unit_value *= conversion_rate
        total_value *= conversion_rate

        vendor_code = part_numbers_by_item.get(str(idx), "") or _extract_part_number_from_description(desc)

        description = _best_description(desc, config_rows, idx)

        ws.append([
            vendor_code,
            "",
            description,
            qty_value,
            total_value,
            total_value,
            "",
        ])

    # Formatting
    header_fill = PatternFill(fill_type="solid", start_color="D9EAF7", end_color="D9EAF7")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row[4:7]:
            cell.number_format = number_format
        row[3].number_format = "#,##0"

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
