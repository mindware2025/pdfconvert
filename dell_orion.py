from datetime import datetime
from io import BytesIO
from typing import List, Tuple
import re

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from dell import _extract_dell_quote_items, CURRENCY_NUMBER_FORMATS


def _build_orion_description(
    item_key: str,
    pricing_desc: str,
    config_rows: List[Tuple[str, str, str, str, str, str]],
) -> str:
    if not config_rows:
        return pricing_desc

    processor = ""
    graphics = ""
    operating_system = ""
    memory = ""
    storage_parts: List[str] = []

    storage_tokens = ("storage", "m.2", "hard drive", "ssd", "flex bay", "pci")

    for row_item, _heading, module, desc, _sku, _qty in config_rows:
        if row_item != item_key:
            continue
        mod = module.strip().lower()
        value = desc.strip() or module.strip()
        if not value:
            continue

        normalized = re.sub(r"\s+", " ", mod)
        if normalized == "processor" or normalized.startswith("processor"):
            processor = value
            continue
        if normalized == "graphics" or normalized.startswith("graphics"):
            if "holder" not in normalized:
                graphics = value
            continue
        if "operating system" in normalized:
            operating_system = value
            continue
        if normalized == "memory" or normalized.startswith("memory"):
            memory = value
            continue
        if any(token in normalized for token in storage_tokens):
            if value not in storage_parts:
                storage_parts.append(value)
            continue

    description_parts: List[str] = []
    if processor:
        description_parts.append(processor)
    if graphics:
        description_parts.append(graphics)
    if operating_system:
        description_parts.append(operating_system)
    if memory:
        description_parts.append(memory)
    if storage_parts:
        storage_text = "; ".join(storage_parts)
        description_parts.append(storage_text)

    if description_parts:
        return " | ".join(description_parts)
    return pricing_desc


def build_dell_orion_output_filename(input_excel_bytes: bytes) -> str:
    quote_date = datetime.now().strftime('%Y%m%d_%H%M')
    return f"Dell_Orion_Quotation_{quote_date}.xlsx"


def generate_orion_quote(
    input_excel_bytes: bytes,
    currency_code: str = "USD",
) -> bytes:
    items, consolidation_fee, part_numbers_by_item, config_rows = _extract_dell_quote_items(
        input_excel_bytes,
        currency_code=currency_code,
    )

    total_original_value = sum((qty_val * (unit_val or 0.0)) for _, qty_val, unit_val, _ in items)
    consolidation_ratio = (consolidation_fee / total_original_value) if total_original_value else 0.0

    wb = Workbook()
    ws = wb.active
    ws.title = "Orion Quote"

    headers = [
        "Vendor Item Code",
        "P/N - Orion Item code",
        "Description",
        "Qty",
        "MSRP",
        "Unit Cost",
        "Unit Selling",
    ]
    for idx, label in enumerate(headers, start=1):
        ws.cell(row=1, column=idx).value = label
        ws.cell(row=1, column=idx).font = Font(bold=True)
        ws.cell(row=1, column=idx).alignment = Alignment(horizontal="center", vertical="center")

    currency_fmt = CURRENCY_NUMBER_FORMATS.get(currency_code, '#,##0.00')

    for row_idx, (desc_text, qty_val, unit_val, _subtotal_val) in enumerate(items, start=2):
        item_key = str(row_idx - 1)
        part_number = part_numbers_by_item.get(item_key, "")
        orion_description = _build_orion_description(item_key, desc_text, config_rows)

        ws.cell(row=row_idx, column=1).value = part_number
        ws.cell(row=row_idx, column=2).value = part_number
        ws.cell(row=row_idx, column=3).value = orion_description
        ws.cell(row=row_idx, column=4).value = qty_val
        # MSRP and Unit Cost should include consolidation/shipping costs
        consolidated_unit = round((unit_val or 0.0) * (1.0 + consolidation_ratio), 2)
        ws.cell(row=row_idx, column=5).value = consolidated_unit
        ws.cell(row=row_idx, column=6).value = consolidated_unit

        # Per request: leave Unit Selling blank
        ws.cell(row=row_idx, column=7).value = ""

        ws.cell(row=row_idx, column=5).number_format = currency_fmt
        ws.cell(row=row_idx, column=6).number_format = currency_fmt
        ws.cell(row=row_idx, column=7).number_format = currency_fmt

    for col_idx, width in enumerate([20, 20, 60, 10, 15, 15, 15], start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
