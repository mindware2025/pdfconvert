import re
from typing import Dict, Tuple


def _cell_to_text(v, fallback=""):
    if v is None:
        return fallback
    if isinstance(v, str):
        return v
    return str(v)


def find_label_value(ws, labels: Tuple[str, ...], max_rows: int = 40, max_cols: int = 10) -> str:
    """Find the first value next to a matching label in a Dell quote sheet."""
    for r in range(1, min(ws.max_row, max_rows) + 1):
        for c in range(1, min(ws.max_column, max_cols) + 1):
            cell_text = _cell_to_text(ws.cell(r, c).value).strip().lower()
            if not cell_text:
                continue

            if any(label in cell_text for label in labels):
                for next_c in range(c + 1, min(ws.max_column, max_cols) + 1):
                    candidate = _cell_to_text(ws.cell(r, next_c).value).strip()
                    if candidate:
                        return candidate
                for next_r in range(r + 1, min(ws.max_row, max_rows) + 1):
                    candidate = _cell_to_text(ws.cell(next_r, c).value).strip()
                    if candidate:
                        return candidate
    return ""


def is_configuration_sheet_name(name: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]", "", (name or "").lower().strip())
    return normalized in (
        "configuration",
        "config",
        "configsheet",
        "configurationsheet",
        "configurationdetails",
        "configdetails",
        "productdetails",
        "productdetailssheet",
        "productdetailsconfiguration",
        "configdetailsheet",
    )


def find_grouped_config_header(ws):
    """Detect newer Dell grouped-config quote layouts used by EUR-style exports."""
    for r in range(1, min(ws.max_row, 40) + 1):
        row_values = [_cell_to_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        normalized = [re.sub(r"\s+", " ", _cell_to_text(v).strip().lower()) for v in row_values]

        has_description = any("description" in name for name in normalized)
        has_sku = any("sku" in name or "part number" in name or "part no" in name for name in normalized)
        has_qty = any(name in ("qty", "quantity", "q-ty") for name in normalized)
        has_unit = any("unit selling price" in name or "unit price" in name for name in normalized)
        has_total = any("total selling price" in name or "total price" in name for name in normalized)

        if has_description and has_sku and has_qty and has_unit and has_total:
            columns = {}
            for c, name in enumerate(normalized, start=1):
                if "description" in name and "description" not in columns:
                    columns["description"] = c
                if "sku" in name or "part number" in name or "part no" in name:
                    columns.setdefault("sku", c)
                if name in ("qty", "quantity", "q-ty"):
                    columns.setdefault("qty", c)
                if "unit selling price" in name or "unit price" in name:
                    columns.setdefault("unit", c)
                if "total selling price" in name or "total price" in name:
                    columns.setdefault("total", c)
            if "description" in columns and "sku" in columns:
                return r, columns
    return None


def find_compact_quote_header(ws):
    """Detect compact Dell quote layouts used by newer quote exports."""
    search_end = min(ws.max_row, 40)
    for r in range(1, search_end + 1):
        columns = {}
        for c in range(1, ws.max_column + 1):
            name = _cell_to_text(ws.cell(r, c).value).strip().lower()
            if not name:
                continue
            if name == "#" and "item" not in columns:
                columns["item"] = c
            if "sku" in name and "sku" not in columns:
                columns["sku"] = c
            if "description" in name and "description" not in columns:
                columns["description"] = c
            if name in ("q-ty", "qty", "quantity") and "qty" not in columns:
                columns["qty"] = c
            if "unit selling price" in name and "unit" not in columns:
                columns["unit"] = c
            if "unit price" in name and "unit" not in columns:
                columns["unit"] = c
            if "total selling price" in name and "total" not in columns:
                columns["total"] = c
            if "total price" in name and "total" not in columns:
                columns["total"] = c

        has_description = "description" in columns
        has_qty = "qty" in columns
        has_total = "total" in columns
        has_sku_or_item = "sku" in columns or "item" in columns

        if has_description and has_qty and has_total and has_sku_or_item:
            return r, columns
    return None
