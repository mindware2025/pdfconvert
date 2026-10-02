# southcomp_engine.py
# Standalone EUR-only quote engine for Dell Quotation Southcomp Polaris.
# No imports from dell.py or any other dell_* module.

from datetime import datetime, timedelta
from io import BytesIO
from typing import Dict, List, Optional, Tuple
import os
import re
import xml.etree.ElementTree as ET
import zipfile

import openpyxl
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string as colidx

try:
    from PIL import Image as PILImage, ImageChops
except ImportError:
    PILImage = ImageChops = None


# ==================== CONSTANTS ====================

CURRENCY_FORMATS = {
    "EUR": '"€"#,##0.00',
    "USD": '"$"#,##0.00',
}

# EUR is the base; USD conversion keeps original USD prices from the BOQ
CONVERSION_RATES: Dict[str, float] = {
    "EUR": 0.92,
    "USD": 1.0,
}


# ==================== HELPERS ====================

def _parse_money(val) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = re.sub(r"[^\d,.\-]", "", str(val).strip())
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _cell_to_text(v, fallback: str = "") -> str:
    if v is None:
        return fallback
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    return _sanitize_excel_text(str(v).strip())


def _sanitize_excel_text(value: str) -> str:
    if value is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(value))[:32767]


def _normalize_text(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower()) if s else ""


def _strip_trailing_asterisk(value: str) -> str:
    if value is None:
        return ""
    text = _cell_to_text(value).split("|")[0]
    return re.sub(r"\s*\*+", "", text).strip()


def _row_text(ws, r: int, c1: int = 1, c2: Optional[int] = None) -> str:
    if c2 is None:
        c2 = ws.max_column
    return " ".join(
        _cell_to_text(ws.cell(r, c).value)
        for c in range(c1, c2 + 1)
        if ws.cell(r, c).value
    ).strip()


def _is_price_or_qty_line(text: str) -> bool:
    t = text.lower()
    return any(tok in t for tok in [
        "qty", "quantity", "unit price", "subtotal", "total", "price",
        "amount", "discount", "tax", "grand total", "msrp", "usd", "aed", "eur", "sar",
    ]) or bool(re.search(r"(\$|€|£|aed|usd|eur|sar)", t, re.IGNORECASE))


def _sanitize_filename_part(value: str) -> str:
    text = re.sub(r"\s+", " ", _cell_to_text(value)).strip()
    if not text:
        return ""
    text = re.sub(r'[<>:"/\\|?*]', "", text)
    return text.rstrip(". ")


_MONTH_NAMES = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}
_MONTH_DATE_PAT = re.compile(
    r"([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})"
)


def _parse_month_date(text: str) -> str:
    """Parse 'July 15, 2026' → '15/07/2026'. Returns '' on failure."""
    m = _MONTH_DATE_PAT.search(text)
    if not m:
        return ""
    month_str, day_str, year_str = m.groups()
    month_num = _MONTH_NAMES.get(month_str.lower())
    if not month_num:
        return ""
    return f"{int(day_str):02d}/{month_num:02d}/{year_str}"


# ==================== LOGO ====================

def _trim_logo(pil_img):
    if PILImage is None:
        return pil_img
    img = pil_img.convert("RGBA")
    alpha = img.getchannel("A")
    bbox = alpha.getbbox()
    if bbox:
        img = img.crop(bbox)
    if ImageChops is not None:
        bg = PILImage.new("RGBA", img.size, (255, 255, 255, 0))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        if bbox:
            img = img.crop(bbox)
    return img


def _pil_to_xl(pil_img):
    buf = BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    return XLImage(buf)


def _logo_path() -> Optional[str]:
    base = os.path.dirname(os.path.abspath(__file__))
    for name in ["ims.png","spc.png","dell spc.png", "dell copy.png", "dell.png"]:
        for directory in [base, os.path.dirname(base)]:
            p = os.path.join(directory, name)
            if os.path.exists(p):
                return p
    return None


def _add_logo(ws, anchor: str = "A1", width: int = 780, height: int = 52) -> None:
    path = _logo_path()
    if not path:
        return
    if PILImage is not None:
        try:
            img = _pil_to_xl(_trim_logo(PILImage.open(path)))
            img.width = width
            img.height = height
            ws.add_image(img, anchor)
            return
        except Exception:
            pass
    try:
        img = XLImage(path)
        img.width = width
        img.height = height
        ws.add_image(img, anchor)
    except Exception:
        pass


# ==================== TEMPLATE DETECTION ====================

def detect_template_type(input_bytes: bytes) -> str:
    """Return 'extended_services' or 'standard_quote'."""
    if input_bytes.lstrip().startswith(b"%PDF"):
        return "standard_quote"
    try:
        wb = openpyxl.load_workbook(BytesIO(input_bytes), data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=80, max_col=10):
            for cell in row:
                if isinstance(cell.value, str) and "dell extended services details" in cell.value.lower():
                    return "extended_services"
    except Exception:
        pass
    return "standard_quote"


# ==================== HEADER / COLUMN DETECTION ====================

def _is_qar_report(ws) -> bool:
    return "quote analysis report" in _cell_to_text(ws.cell(1, 1).value).lower()


def _find_compact_header(ws) -> Optional[Tuple[int, Dict[str, int]]]:
    for r in range(1, min(ws.max_row, 40) + 1):
        cols: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            name = _cell_to_text(ws.cell(r, c).value).strip().lower()
            if not name:
                continue
            if name == "#" and "item" not in cols:
                cols["item"] = c
            if "sku" in name and "sku" not in cols:
                cols["sku"] = c
            if "description" in name and "description" not in cols:
                cols["description"] = c
            if name in ("q-ty", "qty", "quantity") and "qty" not in cols:
                cols["qty"] = c
            if ("unit selling price" in name or "unit price" in name) and "unit" not in cols:
                cols["unit"] = c
            if ("total selling price" in name or "total price" in name) and "total" not in cols:
                cols["total"] = c
        if all(k in cols for k in ("description", "qty", "total")) and ("sku" in cols or "item" in cols):
            return r, cols
    return None


def _find_grouped_header(ws) -> Optional[Tuple[int, Dict[str, int]]]:
    for r in range(1, min(ws.max_row, 40) + 1):
        row_values = [_cell_to_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        normalized = [re.sub(r"\s+", " ", v.strip().lower()) for v in row_values]
        has_desc = any("description" in n for n in normalized)
        has_sku = any("sku" in n or "part number" in n or "part no" in n for n in normalized)
        has_qty = any(n in ("qty", "quantity", "q-ty") for n in normalized)
        has_unit = any("unit selling price" in n or "unit price" in n for n in normalized)
        has_total = any("total selling price" in n or "total price" in n for n in normalized)
        if has_desc and has_sku and has_qty and has_unit and has_total:
            cols: Dict[str, int] = {}
            for c, n in enumerate(normalized, start=1):
                if "description" in n and "description" not in cols:
                    cols["description"] = c
                if ("sku" in n or "part number" in n or "part no" in n) and "sku" not in cols:
                    cols["sku"] = c
                if n in ("qty", "quantity", "q-ty") and "qty" not in cols:
                    cols["qty"] = c
                if ("unit selling price" in n or "unit price" in n) and "unit" not in cols:
                    cols["unit"] = c
                if ("total selling price" in n or "total price" in n) and "total" not in cols:
                    cols["total"] = c
            if "description" in cols and "sku" in cols:
                return r, cols
    return None


def _find_generic_header(ws) -> Tuple[int, int, int, int]:
    """Return (first_data_row, desc_col, qty_col, unit_col)."""
    for r in range(1, min(ws.max_row, 40) + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(row_vals):
            continue
        texts = [_cell_to_text(v).lower() for v in row_vals]
        if any("description" in t for t in texts) and any("qty" in t or "quantity" in t for t in texts):
            desc_idx = qty_idx = unit_idx = None
            for i, v in enumerate(row_vals, start=1):
                name = _cell_to_text(v).lower()
                if desc_idx is None and "description" in name:
                    desc_idx = i
                if qty_idx is None and ("qty" in name or "quantity" in name):
                    qty_idx = i
                if unit_idx is None and ("unit price" in name or "unitprice" in name or name == "price"):
                    unit_idx = i
            return r + 1, desc_idx or 3, qty_idx or 4, unit_idx or 5
    return 8, 3, 4, 5


# ==================== METADATA EXTRACTION ====================

def _scan_all_quote_refs(ws, max_rows: int = 80) -> List[str]:
    pat = r"\b\d{6,}(?:\.[A-Za-z0-9]+)?[A-Za-z0-9\-]*\b"
    refs = []
    for r in range(1, min(ws.max_row, max_rows) + 1):
        for c in range(1, min(ws.max_column, 10) + 1):
            text = _cell_to_text(ws.cell(r, c).value)
            low = text.lower()
            if low.startswith("quote") and "quoted on" not in low:
                m = re.search(pat, text)
                if m:
                    refs.append(m.group(0))
            elif any(tok in low for tok in ("quote no", "quote number", "quote ref")):
                m = re.search(pat, text)
                if m:
                    refs.append(m.group(0))
                else:
                    row_texts = [_cell_to_text(ws.cell(r, cc).value) for cc in range(1, ws.max_column + 1)]
                    for t in row_texts:
                        m = re.search(pat, t)
                        if m:
                            refs.append(m.group(0))
                            break
    seen = []
    for ref in refs:
        if ref not in seen:
            seen.append(ref)
    return seen


def _find_label_value(ws, labels: Tuple[str, ...], max_rows: int = 60, max_cols: int = 10) -> str:
    for r in range(1, min(ws.max_row, max_rows) + 1):
        for c in range(1, min(ws.max_column, max_cols) + 1):
            text = _cell_to_text(ws.cell(r, c).value).strip().lower()
            if not text:
                continue
            if any(label in text for label in labels):
                for nc in range(c + 1, min(ws.max_column, max_cols) + 1):
                    candidate = _cell_to_text(ws.cell(r, nc).value).strip()
                    if candidate:
                        return candidate
                for nr in range(r + 1, min(ws.max_row, max_rows) + 1):
                    candidate = _cell_to_text(ws.cell(nr, c).value).strip()
                    if candidate:
                        return candidate
    return ""


def _extract_metadata(ws) -> Tuple[str, str]:
    """Return (quote_ref, date)."""
    ref = _find_label_value(ws, ("quote no", "quote number", "quote ref", "quotation no"))
    if not ref:
        raw = ws["E15"].value
        ref = "" if raw is None else (raw.strftime("%d/%m/%Y") if isinstance(raw, datetime) else str(raw).strip())

    date = _find_label_value(ws, ("quote date", "quoted on", "date"))
    if not date:
        raw_d = ws["E18"].value
        if isinstance(raw_d, datetime):
            date = raw_d.strftime("%d/%m/%Y")
        else:
            date = "" if raw_d is None else str(raw_d).strip()

    all_refs = _scan_all_quote_refs(ws)
    if all_refs:
        combined = []
        if ref:
            combined.append(ref)
        for r in all_refs:
            if r not in combined:
                combined.append(r)
        ref = ", ".join(combined)

    # Fallback: scan for quoted-on date
    if not date:
        for row in ws.iter_rows(min_row=1, max_row=80, max_col=10):
            row_text = " ".join(_cell_to_text(c.value) for c in row).lower()
            if "quoted on" in row_text or "quote date" in row_text:
                for cell in row:
                    m = re.search(r"\d{2}/\d{2}/\d{4}", str(cell.value))
                    if m:
                        date = m.group(0)
                        break

    return ref, date


def _extract_expiry(ws) -> str:
    def _fmt(value) -> str:
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        text = str(value or "").strip()
        m = re.search(r"\d{2}/\d{2}/\d{4}", text)
        return m.group(0) if m else text

    def _adjust(value: str) -> str:
        if not value:
            return ""
        try:
            return (datetime.strptime(value, "%d/%m/%Y") - timedelta(days=2)).strftime("%d/%m/%Y")
        except Exception:
            return value

    # Try strict position first
    direct = _fmt(ws["E19"].value)
    if direct:
        return _adjust(direct)

    for r in range(1, min(ws.max_row, 80) + 1):
        row_text = " ".join(_cell_to_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 10) + 1)).lower()
        if "expires by" not in row_text:
            continue
        for c in range(1, min(ws.max_column, 10) + 1):
            if "expires by" in _cell_to_text(ws.cell(r, c).value).lower():
                for nc in range(c + 1, min(ws.max_column, 10) + 1):
                    candidate = _fmt(ws.cell(r, nc).value)
                    if candidate:
                        return _adjust(candidate)
        m = re.search(r"\d{2}/\d{2}/\d{4}", row_text)
        if m:
            return _adjust(m.group(0))
    return ""


def _extract_quote_metadata(ws) -> Dict[str, str]:
    keys = {"company name", "customer name", "customer number", "end user", "reseller"}
    out = {k: "" for k in keys}
    max_row = min(ws.max_row, 120)
    for r in range(1, max_row + 1):
        label = _cell_to_text(ws.cell(r, 2).value).strip().lower().rstrip(":")
        if label in keys:
            out[label] = _cell_to_text(ws.cell(r, 5).value)
    # Shipping information block (PDF-style Excel)
    for r in range(1, max_row + 1):
        row_values = [_cell_to_text(ws.cell(r, c).value) for c in range(1, 11)]
        if any("shipping information" in v.lower() for v in row_values):
            for idx, v in enumerate(row_values):
                if "shipping information" in v.lower():
                    col = idx + 1
                    lines = []
                    for nr in range(r + 1, min(ws.max_row, r + 12) + 1):
                        cv = _cell_to_text(ws.cell(nr, col).value)
                        if not cv:
                            break
                        if any(m in cv.lower() for m in ("quote summary", "payment details", "product details")):
                            break
                        lines.append(cv.strip())
                    if lines:
                        out["end user"] = "\n".join(lines)
                    break
            break
    return out


def _extract_grouped_metadata(ws) -> Tuple[str, str]:
    refs = _scan_all_quote_refs(ws, max_rows=200)
    ref = ", ".join(refs)
    date = ""
    for r in range(1, min(ws.max_row, 200) + 1):
        first = _cell_to_text(ws.cell(r, 1).value).strip().lower()
        if first.startswith("date"):
            v = ws.cell(r, 2).value
            date = v.strftime("%d/%m/%Y") if isinstance(v, datetime) else _cell_to_text(v)
    return ref, date


_QAR_CUSTOMER_LABELS = {
    "bill to customer": "bill_to",
    "sold to customer": "sold_to",
    "end user customer": "end_user",
}


def _find_qar_customer_block(ws, max_rows: int = 20) -> Optional[Tuple[int, Dict[int, str]]]:
    for r in range(1, min(ws.max_row, max_rows) + 1):
        col_map: Dict[int, str] = {}
        for c in range(1, min(ws.max_column, 6) + 1):
            text = _cell_to_text(ws.cell(r, c).value).strip().lower().rstrip(":")
            if text in _QAR_CUSTOMER_LABELS:
                col_map[c] = _QAR_CUSTOMER_LABELS[text]
        if col_map:
            return r, col_map
    return None


def _extract_qar_metadata(ws) -> Tuple[str, Dict[str, str]]:
    quote_ref = _find_label_value(ws, ("quote number",), max_rows=10)
    quote_name = _find_label_value(ws, ("quote name",), max_rows=10)

    bill_to = sold_to = end_user = end_user_id = ""
    block = _find_qar_customer_block(ws)
    if block:
        label_row, col_map = block
        for col, key in col_map.items():
            name = _cell_to_text(ws.cell(label_row + 1, col).value)
            ident = _cell_to_text(ws.cell(label_row + 2, col).value)
            if key == "bill_to":
                bill_to = name
            elif key == "sold_to":
                sold_to = name
            elif key == "end_user":
                end_user, end_user_id = name, ident

    quote_meta = {
        "company name": bill_to or quote_name,
        "customer name": sold_to,
        "end user": f"{end_user}\n{end_user_id}" if end_user and end_user_id else end_user,
        "reseller": "",
    }
    return quote_ref, quote_meta


# ==================== ITEMS EXTRACTION ====================

def _extract_items_compact(ws) -> Tuple[List, List]:
    header_info = _find_compact_header(ws)
    if not header_info:
        return [], []
    header_row, cols = header_info
    items: List[Tuple] = []
    config_rows: List[Tuple] = []
    current_item: Optional[str] = None
    blank_streak = 0
    for r in range(header_row + 1, ws.max_row + 1):
        row_text = _row_text(ws, r, 1, ws.max_column)
        if not row_text:
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0
        first_cell = _cell_to_text(ws.cell(r, cols.get("item", 1)).value).strip()
        sku = _cell_to_text(ws.cell(r, cols["sku"]).value) if "sku" in cols else ""
        desc = _cell_to_text(ws.cell(r, cols["description"]).value)
        qty_raw = _cell_to_text(ws.cell(r, cols["qty"]).value)
        unit_col = cols.get("unit")
        unit_price = (_parse_money(ws.cell(r, unit_col).value) or 0.0) if unit_col else None
        total_price = _parse_money(ws.cell(r, cols["total"]).value)
        if not any([first_cell, sku, desc, qty_raw, unit_price, total_price]):
            continue
        if total_price is None and any(t in row_text.lower() for t in ("total", "subtotal", "quote number", "solution id")):
            continue
        try:
            qty_val = int(qty_raw) if qty_raw not in (None, "") else 0
        except Exception:
            qty_val = int(_parse_money(qty_raw) or 0)
        if unit_col is None and total_price is not None:
            unit_price = total_price
        elif unit_price == 0.0 and qty_val > 0 and total_price is not None:
            unit_price = total_price / qty_val
        if first_cell:
            if desc and qty_val > 0:
                items.append((desc, qty_val, unit_price, total_price))
                current_item = str(len(items))
            continue
        if current_item and (sku or desc):
            config_rows.append((current_item, "", "", desc, "", qty_raw))
    return items, config_rows


def _is_grouped_summary_row(ws, r: int, cols: Dict) -> bool:
    row_text = _row_text(ws, r, 1, ws.max_column).lower()
    if not row_text:
        return False
    first = _cell_to_text(ws.cell(r, 1).value).strip().lower()
    if first.startswith(("quote", "name")):
        return True
    if "consolidation fee" in row_text:
        return not _cell_to_text(ws.cell(r, cols.get("sku", 0)).value)
    return "total" in row_text and "total selling price" not in row_text


_QUOTE_REF_PAT = re.compile(r"\b\d{6,}(?:\.[A-Za-z0-9]+)?[A-Za-z0-9\-]*\b")


def _extract_items_grouped(ws, item_quote_refs: Optional[Dict[str, str]] = None) -> Tuple[List, List]:
    """item_quote_refs, when given, is filled with {item_no: quote ref} — each
    item's ref comes from the "Quote | <ref>" row above it (a grouped BOQ can
    hold several quotes), falling back to the item row's own "Config" cell."""
    header_info = _find_grouped_header(ws)
    if not header_info:
        return [], []
    header_row, cols = header_info
    config_col = next(
        (c for c in range(1, ws.max_column + 1)
         if _cell_to_text(ws.cell(header_row, c).value).strip().lower() == "config"),
        None,
    )
    current_quote_ref = ""
    items: List[Tuple] = []
    config_rows: List[Tuple] = []
    current_item: Optional[str] = None
    blank_streak = 0
    for r in range(header_row + 1, ws.max_row + 1):
        row_text = _row_text(ws, r, 1, ws.max_column)
        if not row_text:
            blank_streak += 1
            if blank_streak >= 4:
                break
            continue
        blank_streak = 0
        if _cell_to_text(ws.cell(r, 1).value).strip().lower().startswith("quote"):
            m = _QUOTE_REF_PAT.search(_row_text(ws, r, 2, ws.max_column))
            if m:
                current_quote_ref = m.group(0)
        if _is_grouped_summary_row(ws, r, cols):
            continue
        first_cell = _cell_to_text(ws.cell(r, 1).value).strip()
        desc = _cell_to_text(ws.cell(r, cols["description"]).value)
        sku = _cell_to_text(ws.cell(r, cols["sku"]).value)
        qty_raw = _cell_to_text(ws.cell(r, cols.get("qty", 0)).value)
        try:
            qty_val = int(qty_raw) if qty_raw else 0
        except Exception:
            qty_val = int(_parse_money(qty_raw) or 0)
        unit_price = _parse_money(ws.cell(r, cols.get("unit", 0)).value) or 0.0
        total_price = _parse_money(ws.cell(r, cols.get("total", 0)).value) or 0.0
        if first_cell.lower().startswith("quote"):
            continue
        if first_cell:
            if desc:
                items.append((desc, qty_val, unit_price, total_price))
                current_item = str(len(items))
                if item_quote_refs is not None:
                    ref = current_quote_ref
                    if not ref and config_col:
                        m = _QUOTE_REF_PAT.search(_cell_to_text(ws.cell(r, config_col).value))
                        ref = m.group(0) if m else ""
                    if ref:
                        item_quote_refs[current_item] = ref
            continue
        if current_item and desc:
            config_rows.append((current_item, "", "", desc, sku, qty_raw))
    return items, config_rows


def _find_qar_table_header(ws) -> Optional[int]:
    for r in range(1, ws.max_row + 1):
        if _cell_to_text(ws.cell(r, 1).value).strip().lower() == "order code":
            return r
    return None


def _extract_items_qar(ws) -> Tuple[List, List]:
    header_row = _find_qar_table_header(ws)
    if header_row is None:
        return [], []
    items: List[Tuple] = []
    config_rows: List[Tuple] = []
    current_item: Optional[str] = None
    for r in range(header_row + 1, ws.max_row + 1):
        order_code = _cell_to_text(ws.cell(r, 1).value).strip()
        category = _cell_to_text(ws.cell(r, 2).value).strip()
        qty_raw = _cell_to_text(ws.cell(r, 3).value).strip()
        sku = _cell_to_text(ws.cell(r, 4).value).strip()
        desc = _cell_to_text(ws.cell(r, 5).value).strip()
        price = _parse_money(ws.cell(r, 6).value)
        if not any([order_code, category, qty_raw, sku, desc, price]):
            continue
        try:
            qty_val = int(qty_raw) if qty_raw else 0
        except Exception:
            qty_val = int(_parse_money(qty_raw) or 0)
        if order_code:
            if desc:
                total_price = price or 0.0
                unit_price = (total_price / qty_val) if qty_val else 0.0
                items.append((desc, qty_val, unit_price, total_price))
                current_item = str(len(items))
            continue
        if current_item and (category or desc):
            config_rows.append((current_item, "", category, desc, sku, qty_raw))
    return items, config_rows


def _locate_pricing_summary(ws) -> Optional[Tuple[int, int]]:
    B = colidx("B")
    for r in range(30, min(ws.max_row, 120) + 1):
        v = ws.cell(r, B).value
        if v and "pricing" in str(v).lower() and "summary" in str(v).lower():
            return r + 1, r + 3
    return None


def _extract_items_pricing_summary(ws) -> Optional[List[Tuple]]:
    located = _locate_pricing_summary(ws)
    if not located:
        return None
    _, start_row = located
    A, B, K, L, N = colidx("A"), colidx("B"), colidx("K"), colidx("L"), colidx("N")
    items = []
    r = start_row
    while r <= ws.max_row:
        sr = ws.cell(r, A).value
        if sr is None or not re.match(r"^\d+", str(sr).strip()):
            break
        desc = _cell_to_text(ws.cell(r, B).value)
        if not desc:
            break
        qty_val = int(_parse_money(ws.cell(r, K).value) or 0)
        unit_val = _parse_money(ws.cell(r, L).value) or 0.0
        sub_val = _parse_money(ws.cell(r, N).value)
        if sub_val is None:
            sub_val = qty_val * unit_val
        if qty_val <= 0 and unit_val == 0.0 and (sub_val is None or sub_val == 0.0):
            break
        items.append((desc, qty_val, unit_val, sub_val))
        r += 1
    return items if items else None


def _extract_pdf_metadata_by_position(pdf_bytes: bytes) -> Dict[str, str]:
    """
    Extract customer metadata from page 1 of a Dell portal PDF using word X-positions.
    The PDF uses a 2-column layout; the right column (x >= ~200) holds the actual values.
    Returns keys: quote_creator, end_user (shipping address), quote_name.
    """
    out = {"quote_creator": "", "end_user": "", "quote_name": ""}
    try:
        import pdfplumber
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[0]
            words = page.extract_words(use_text_flow=True)
    except Exception:
        return out

    # Group words by y position
    rows: Dict[int, List] = {}
    for w in words:
        y = round(w.get("top", 0))
        rows.setdefault(y, []).append(w)

    # Detect the right-column x boundary from the "Quote Creator:" label
    col2_x = 200.0
    for y in sorted(rows):
        row_words = sorted(rows[y], key=lambda w: w.get("x0", 0))
        line = " ".join(w["text"] for w in row_words).lower()
        if "quote creator" in line and "quote name" in line:
            for w in row_words:
                if "quote" in w["text"].lower() and w.get("x0", 0) > 100:
                    col2_x = w.get("x0", 200.0)
                    break
            break

    # State machine over sorted rows
    next_row_is_quote_name_creator = False
    next_row_is_reseller = False
    in_shipping = False

    for y in sorted(rows):
        row_words = sorted(rows[y], key=lambda w: w.get("x0", 0))
        line = " ".join(w["text"] for w in row_words).strip()
        low = line.lower()

        # Stop at "Quote Summary" or "Custom Fields"
        if any(stop in low for stop in ("quote summary", "custom fields")):
            break

        left_words = [w["text"] for w in row_words if w.get("x0", 0) < col2_x]
        right_words = [w["text"] for w in row_words if w.get("x0", 0) >= col2_x]
        left_text = " ".join(left_words).strip().lower().rstrip(":")
        right_text = " ".join(right_words).strip()

        # "Quote Name:" (left) / "Quote Creator:" (right) — label row
        if "quote name" in left_text and "quote creator" in right_text.lower():
            next_row_is_quote_name_creator = True
            in_shipping = False
            continue

        if next_row_is_quote_name_creator:
            next_row_is_quote_name_creator = False
            out["quote_name"] = " ".join(left_words).strip()
            out["quote_creator"] = right_text
            continue

        if next_row_is_reseller:
            next_row_is_reseller = False
            out["reseller"] = right_text.strip()
            continue

        # "Page Name:" (left) / "Authorized Partner:" (right) — label row → next row has reseller
        if "authorized partner" in right_text.lower() and not out.get("reseller"):
            next_row_is_reseller = True
            in_shipping = False
            continue

        # "Shipping Information:" label row — left column may say "Billing Information:" or be blank
        if "shipping information" in right_text.lower():
            in_shipping = True
            continue

        if in_shipping:
            # Right column = shipping address; left column is billing (usually "-", skip)
            if right_text and right_text != "-":
                if out["end_user"]:
                    out["end_user"] += "\n" + right_text
                else:
                    out["end_user"] = right_text

    return out


def _extract_items_generic(ws) -> List[Tuple]:
    first_data_row, desc_col, qty_col, unit_col = _find_generic_header(ws)
    items = []
    r = first_data_row
    while r <= ws.max_row:
        desc = _cell_to_text(ws.cell(r, desc_col).value)
        qty = ws.cell(r, qty_col).value
        unit = ws.cell(r, unit_col).value
        if not desc or desc.lower().startswith("total"):
            break
        try:
            qty_val = int(qty) if qty not in (None, "") else 0
        except Exception:
            qty_val = int(_parse_money(qty) or 0)
        unit_val = _parse_money(unit) or 0.0
        if qty_val > 0:
            items.append((desc, qty_val, unit_val, None))
        r += 1
    return items


def _extract_items_pdf(pdf_bytes: bytes) -> Tuple[List, Dict, List, str, str, str, float]:
    """Extract items, metadata, config_rows, quote_ref, date, expiry, consolidation_fee from PDF."""
    lines = _extract_pdf_lines(pdf_bytes)
    metadata = {"company name": "", "customer name": "", "customer number": "",
                 "end user": "", "reseller": "", "quote creator": "", "shipping info": ""}
    quote_ref = date_text = expiry_text = ""
    consolidation_fee = 0.0

    pending_keys: List[str] = []
    prev_label: Optional[str] = None
    in_items = False
    items: List[Tuple] = []
    config_rows: List[Tuple] = []
    # Flag: next non-empty line holds the values for quote number / date / expiry
    _next_line_is_quote_values = False

    for line in lines:
        low = line.lower().strip()
        if not low:
            continue

        # ---- Parse the value line that follows "Quote number: Quote date: Quote expiration:" ----
        if _next_line_is_quote_values:
            _next_line_is_quote_values = False
            # Pattern: <ref_number> <Month DD, YYYY> <Month DD, YYYY>
            # Find all date tokens in "Month DD, YYYY" form
            all_dates = _MONTH_DATE_PAT.findall(line)
            all_nums = re.findall(r"\b(\d{6,})\b", line)
            if not quote_ref and all_nums:
                quote_ref = all_nums[0]
            dates_fmt = []
            for month_s, day_s, year_s in all_dates:
                mn = _MONTH_NAMES.get(month_s.lower())
                if mn:
                    dates_fmt.append(f"{int(day_s):02d}/{mn:02d}/{year_s}")
            if not date_text and len(dates_fmt) >= 1:
                date_text = dates_fmt[0]
            if not expiry_text and len(dates_fmt) >= 2:
                expiry_text = dates_fmt[1]
            continue

        # Detect the combined label line (Dell PDF portal format)
        if "quote number" in low and "quote date" in low:
            _next_line_is_quote_values = True
            continue

        # Metadata extraction
        if not in_items:
            for key in ("company name", "customer name", "customer number", "end user", "reseller", "quote creator", "shipping info"):
                if low.rstrip(":") == key or low.startswith(key + ":"):
                    pending_keys = [key]
                    prev_label = key
                    rest = line[len(key):].lstrip(":").strip()
                    if rest:
                        metadata[key] = rest
                    break
            else:
                if pending_keys and prev_label:
                    if not any(low.rstrip(":") == k or low.startswith(k + ":") for k in metadata):
                        metadata[prev_label] = (metadata[prev_label] + " " + line.strip()).strip()
                    else:
                        pending_keys = []
                        prev_label = None

            # Fallback: inline date search (handles "dd/mm/yyyy" format)
            if not quote_ref:
                m = re.search(r"\b\d{6,}(?:\.[A-Za-z0-9]+)?[A-Za-z0-9\-]*\b", line)
                if m:
                    quote_ref = m.group(0)
            if not date_text:
                m = re.search(r"\d{2}/\d{2}/\d{4}", line)
                if m and any(t in low for t in ("quote date", "quoted on", "date")):
                    date_text = m.group(0)
            if not expiry_text:
                if any(t in low for t in ("quote expiration", "expiry", "expires", "expiration date")):
                    m = re.search(r"\d{2}/\d{2}/\d{4}", line)
                    if m:
                        expiry_text = m.group(0)
                    else:
                        d = _parse_month_date(line)
                        if d:
                            expiry_text = d

        if "quote summary" in low:
            in_items = True
            continue

        if in_items:
            # Stop extracting when we hit another major section
            if any(stop in low for stop in (
                "payment details", "product details", "ship to:", "subtotal:",
            )):
                in_items = False
                continue

            # Skip page footer lines (e.g. "Page 1", "Page 2")
            if re.match(r"^page\s+\d+$", low.strip()):
                continue

            # Match: description $unit qty $total  OR  description unit qty total
            m = re.search(
                r"^(.+?)\s+[$]?([\d,]+[.]?\d*)\s+(\d+)\s+[$]?([\d,]+[.]?\d*)\s*$",
                line,
            )
            if m:
                desc_s, unit_s, qty_s, total_s = m.groups()
                qty_val = int(qty_s)
                unit_val = _parse_money(unit_s) or 0.0
                total_val = _parse_money(total_s) or (qty_val * unit_val)
                items.append((desc_s.strip(), qty_val, unit_val, total_val))
            elif items and not _is_price_or_qty_line(line) and not re.match(r"item\s+unit", low):
                # Continuation line — append to last item description
                old_desc, qty, unit, total = items[-1]
                if old_desc.endswith(","):
                    joined = old_desc.rstrip(",").strip() + ", " + line.strip()
                else:
                    joined = old_desc + " " + line.strip()
                items[-1] = (joined, qty, unit, total)
            elif re.search(r"consolidation fee", low, re.I):
                m2 = re.search(r"[\d,]+\.?\d*", line)
                if m2:
                    consolidation_fee = _parse_money(m2.group(0)) or 0.0

    return items, metadata, config_rows, quote_ref, date_text, expiry_text, consolidation_fee


def _extract_config_from_pdf(pdf_bytes: bytes) -> List[Tuple]:
    """
    Parse the 'Product Details' section of a Dell portal PDF and return config rows.
    Each tuple: (item_number_str, "", category, description, "", "")
    Uses the x-position split (Category col < desc_x, Description col >= desc_x).
    """
    config_rows: List[Tuple] = []
    in_product_details = False
    in_config_section = False
    item_number = 0
    desc_x = 134.0  # will be detected from "Category Description" header

    try:
        import pdfplumber
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                words = page.extract_words(use_text_flow=True)
                if not words:
                    continue
                rows: Dict[int, List] = {}
                for w in words:
                    y = round(w.get("top", 0))
                    rows.setdefault(y, []).append(w)

                for y in sorted(rows):
                    row_words = sorted(rows[y], key=lambda w: w.get("x0", 0))
                    line = " ".join(w["text"] for w in row_words).strip()
                    low = line.lower().strip()

                    if not in_product_details:
                        if "product details" in low:
                            in_product_details = True
                        continue

                    # Item block header — each product's detail section starts with this
                    if "unit price" in low and "qty" in low and "item total" in low:
                        item_number += 1
                        in_config_section = False
                        continue

                    # Stop entirely at end-of-document sections
                    if any(stop in low for stop in (
                        "ship to:", "important notes", "governing terms",
                        "sincerely,", "thanks for shopping", "all orders are subject",
                    )):
                        in_config_section = False
                        in_product_details = False
                        break

                    # Skip page footers, catalog numbers, standalone "Description" header
                    if re.match(r"^page\s+\d+$", low):
                        continue
                    if low.startswith("catalog number"):
                        in_config_section = False
                        continue
                    if low == "description":
                        continue

                    # Skip item price lines inside product details
                    if re.search(r"[$][\d,]+[.]\d+\s+\d+\s+[$][\d,]+", line):
                        continue

                    # "Category Description" header — captures x boundary for this page
                    if "category" in low and "description" in low and len(line.split()) <= 4:
                        for w in row_words:
                            if "description" in w["text"].lower():
                                desc_x = w.get("x0", 134.0)
                                break
                        in_config_section = True
                        continue

                    if not in_config_section:
                        continue

                    # Split by x boundary into category vs description
                    cat_words = [w["text"] for w in row_words if w.get("x0", 0) < desc_x]
                    dsc_words = [w["text"] for w in row_words if w.get("x0", 0) >= desc_x]
                    cat_part = " ".join(cat_words).strip()
                    dsc_part = " ".join(dsc_words).strip()

                    if cat_part and dsc_part:
                        config_rows.append((str(item_number), "", cat_part, dsc_part, "", ""))
                    elif cat_part and not dsc_part and config_rows and config_rows[-1][0] == str(item_number):
                        # Category name wraps to next line — append to previous row's category
                        last = config_rows[-1]
                        config_rows[-1] = (last[0], last[1], last[2] + " " + cat_part, last[3], last[4], last[5])
    except Exception:
        pass

    return config_rows


def _extract_pdf_lines(pdf_bytes: bytes) -> List[str]:
    try:
        import pdfplumber
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            lines = []
            for page in pdf.pages:
                words = page.extract_words(use_text_flow=True)
                if not words:
                    continue
                rows: Dict[int, List] = {}
                for w in words:
                    y = round(w.get("top", 0))
                    rows.setdefault(y, []).append(w)
                for y in sorted(rows):
                    row_words = sorted(rows[y], key=lambda w: w.get("x0", 0))
                    lines.append(" ".join(w.get("text", "") for w in row_words).strip())
        if lines:
            return lines
    except Exception:
        pass
    try:
        from pypdf import PdfReader
    except ImportError:
        raise RuntimeError("pypdf is required to parse PDF quotes")
    reader = PdfReader(BytesIO(pdf_bytes))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    return [l.strip() for l in text.splitlines()]


# ==================== PDF: "PREMIER EQUOTE" TEMPLATE ====================
# Dell's Premier-portal eQuote confirmation PDF ("E-Quote Name:", "E-Quote
# Creator:", "Premier Page Name:" on page 1; items listed under a "Pricing
# Summary" table as "N. Description  Qty  $ListPrice  $UnitPrice  $Subtotal").
# This is structurally different from the "Quote Summary" BTO order-form PDF
# handled by _extract_items_pdf() above, so it gets its own self-contained
# parser. It only runs when its unique markers are detected; otherwise the
# existing _extract_items_pdf()/_extract_config_from_pdf() path is used
# unchanged.

_PREMIER_ITEM_LINE_PAT = re.compile(
    r"^\d+\.\s*(.+?)\s+(\d+)\s+[$]?([\d,]+\.\d+)\s+[$]?([\d,]+\.\d+)\s+[$]?([\d,]+\.\d+)\s*$"
)


def _try_extract_premier_pricing_summary_pdf(
    pdf_bytes: bytes,
) -> Optional[Tuple[List, Dict, List, str, str, str, float]]:
    """Parse the Dell Premier eQuote PDF template.

    Returns (items, metadata, config_rows, quote_ref, date_text, expiry_text,
    consolidation_fee), or None when this template isn't detected so the
    caller can fall back to the existing PDF parsing path.
    """
    lines = _extract_pdf_lines(pdf_bytes)
    if not any("e-quote name" in l.lower() or "e-quote creator" in l.lower() for l in lines):
        return None

    metadata = {"company name": "", "customer name": "", "customer number": "",
                "end user": "", "reseller": "", "quote creator": "", "shipping info": ""}
    quote_ref = date_text = expiry_text = ""
    consolidation_fee = 0.0
    items: List[Tuple] = []
    in_items = False

    for line in lines:
        stripped = line.strip()
        low = stripped.lower()
        if not low:
            continue

        if not quote_ref and low.startswith("quote no"):
            m = re.search(r"\d{6,}(?:\.[A-Za-z0-9]+)?", stripped)
            if m:
                quote_ref = m.group(0)
            continue
        if not date_text and low.startswith("quoted on"):
            m = re.search(r"\d{2}/\d{2}/\d{4}", stripped)
            if m:
                date_text = m.group(0)
            continue
        if low.startswith("expires by"):
            m = re.search(r"\d{2}/\d{2}/\d{4}", stripped)
            if m:
                expiry_text = m.group(0)
            continue
        if low.startswith("e-quote creator") and ":" in stripped:
            metadata["quote creator"] = stripped.split(":", 1)[1].strip()
            continue
        if low.startswith("premier page name") and ":" in stripped:
            val = stripped.split(":", 1)[1].strip()
            if val and val != "-":
                metadata["reseller"] = val
            continue
        if low.startswith("company name") and ":" in stripped:
            val = stripped.split(":", 1)[1].strip()
            if val and val != "-":
                metadata["company name"] = val
            continue
        if low.startswith("customer number") and ":" in stripped:
            val = stripped.split(":", 1)[1].strip()
            if val and val != "-":
                metadata["customer number"] = val
            continue
        if low.startswith("shipping:"):
            fee = _parse_money(stripped.split(":", 1)[1]) or 0.0
            if abs(fee) > 1e-9:
                consolidation_fee += fee
            continue

        if "pricing summary" in low:
            in_items = True
            continue

        if in_items:
            if low.startswith("subtotal:"):
                in_items = False
                continue
            m = _PREMIER_ITEM_LINE_PAT.match(stripped)
            if m:
                desc_s, qty_s, _list_s, unit_s, total_s = m.groups()
                qty_val = int(qty_s)
                unit_val = _parse_money(unit_s) or 0.0
                total_val = _parse_money(total_s) or (qty_val * unit_val)
                items.append((desc_s.strip(), qty_val, unit_val, total_val))
            elif items and not _is_price_or_qty_line(stripped):
                old_desc, qty, unit, total = items[-1]
                items[-1] = (f"{old_desc} {stripped}".strip(), qty, unit, total)

    if not items:
        return None

    config_rows = _extract_config_from_pdf_module_table(pdf_bytes)
    return items, metadata, config_rows, quote_ref, date_text, expiry_text, consolidation_fee


def _extract_config_from_pdf_module_table(pdf_bytes: bytes) -> List[Tuple]:
    """Parse the 'Module Description SKU Tax Type Qty' per-item component
    tables in the Premier eQuote PDF's Product Details section.

    Each tuple: (item_number_str, "", module, description, sku, qty)
    """
    sku_pat = re.compile(r"^\d{3,}-[A-Z0-9]+(?:,\d{3,}-[A-Z0-9]+)*$")
    config_rows: List[Tuple] = []
    in_product_details = False
    in_table = False
    item_number = 0
    # desc_x..tax_x is treated as one combined "description + SKU" zone since the
    # SKU header label sits a few points right of where SKU values actually start
    # (right-aligned data vs. left-aligned header) — the SKU token is picked out of
    # that zone by pattern instead of a second x boundary.
    desc_x = tax_x = qty_x = None

    try:
        import pdfplumber
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                words = page.extract_words(use_text_flow=True)
                if not words:
                    continue
                rows: Dict[int, List] = {}
                for w in words:
                    y = round(w.get("top", 0))
                    rows.setdefault(y, []).append(w)

                for y in sorted(rows):
                    row_words = sorted(rows[y], key=lambda w: w.get("x0", 0))
                    line = " ".join(w["text"] for w in row_words).strip()
                    low = line.lower().strip()

                    if not in_product_details:
                        if "product details" in low:
                            in_product_details = True
                        continue

                    # Repeating page header/footer (timestamp title bar, source-file footer)
                    if re.match(r"^page\s+\d+$", low) or "file:///" in low or "your dell quote" in low:
                        continue

                    # End of document
                    if any(stop in low for stop in (
                        "purchase order number", "customer signature", "connect with dell",
                    )):
                        in_table = False
                        in_product_details = False
                        break

                    # Per-item block header ("Qty Unit Price Subtotal") — marks a new item
                    if low.startswith("qty") and "unit price" in low and "subtotal" in low:
                        item_number += 1
                        in_table = False
                        continue

                    # "Module Description SKU Tax Type Qty" header — capture column x boundaries
                    if low.startswith("module") and "description" in low and "sku" in low:
                        for w in row_words:
                            wt = w["text"].lower()
                            if wt == "description":
                                desc_x = w.get("x0")
                            elif wt == "tax":
                                tax_x = w.get("x0")
                            elif wt == "qty":
                                qty_x = w.get("x0")
                        in_table = True
                        continue

                    if not in_table or item_number == 0 or None in (desc_x, tax_x, qty_x):
                        continue

                    module = " ".join(w["text"] for w in row_words if w.get("x0", 0) < desc_x).strip()
                    desc_sku_words = [w["text"] for w in row_words if desc_x <= w.get("x0", 0) < tax_x]
                    sku_words = [w for w in desc_sku_words if sku_pat.match(w)]
                    description = " ".join(w for w in desc_sku_words if not sku_pat.match(w)).strip()
                    sku = " ".join(sku_words).strip()
                    qty = " ".join(w["text"] for w in row_words if w.get("x0", 0) >= qty_x).strip()

                    if not any([module, description, sku, qty]):
                        continue

                    current_item = str(item_number)
                    # A real data row always carries a qty (from the trailing "SR N" tax+qty
                    # pair); a row with no qty is either a wrapped continuation of the row
                    # above (module/description/SKU text that overflowed onto the next line)
                    # or a bare section-header ("Components") — the latter always follows an
                    # item-number transition, so it never has a same-item previous row to
                    # merge into and is correctly kept standalone.
                    if not qty and config_rows and config_rows[-1][0] == current_item:
                        last = config_rows[-1]
                        merged_module = f"{last[2]} {module}".strip() if module else last[2]
                        merged_desc = f"{last[3]} {description}".strip() if description else last[3]
                        merged_sku = f"{last[4]}{sku}" if sku else last[4]
                        config_rows[-1] = (last[0], last[1], merged_module, merged_desc, merged_sku, last[5])
                        continue

                    config_rows.append((current_item, "", module, description, sku, qty))
    except Exception:
        pass

    return config_rows


# ==================== PDF: "CHECKOUT CONFIRMATION" TEMPLATE ====================
# Dell's "Your quote is ready for purchase." checkout-confirmation PDF
# (sent from the online checkout flow, with a "Place your order" button).
# Page 1 shows label/value pairs ("Quote No.:", "Company Name:", "End User:",
# ...) whose LABEL text is duplicated in the PDF's underlying content stream
# — e.g. what renders as "Quote No.: 123" is actually stored as
# "Quote Quote No.: No.: 123" — even though the rendered PDF looks completely
# normal. Items sit under a "Pricing Summary" table as
# "N. Description  Qty  $UnitPrice  $Subtotal" (two price columns), which is
# structurally different from both the "Quote Summary" order-form PDF
# handled by _extract_items_pdf() (single un-duplicated label lines, no
# price columns in the item line) and the Premier eQuote PDF handled by
# _try_extract_premier_pricing_summary_pdf() (three price columns, requires
# "E-Quote Name"/"E-Quote Creator" markers). This parser only runs when its
# own unique markers are detected; otherwise the existing PDF parsing paths
# are used unchanged. The "Product Details" component tables use the same
# "Module Description SKU Tax Type Qty" layout as the Premier template, so
# they're read with the existing _extract_config_from_pdf_module_table().

_CHECKOUT_ITEM_LINE_PAT = re.compile(
    r"^\d+\.\s*(.+?)\s+(\d+)\s+[$]?([\d,]+\.\d+)\s+[$]?([\d,]+\.\d+)\s*$"
)

# Each item's own Product Details heading is split across a few lines by the
# 2-column PDF layout: the description, then "N. Qty $Unit $Subtotal", then
# the item's real Dell SKU alone in parentheses on its own line. A simple
# accessory (mouse, keyboard, sleeve, ...) has no "Module Description SKU..."
# component table at all, so this is the ONLY place its SKU appears.
_CHECKOUT_ITEM_QTY_LINE_RE = re.compile(r"^(\d+)\.\s+\d+\s+[$]?[\d,]+\.\d+\s+[$]?[\d,]+\.\d+\s*$")
_CHECKOUT_SKU_ONLY_LINE_RE = re.compile(r"^\((\d{3,4}-[A-Za-z0-9]+)\)$")


def _extract_checkout_pricing_item_skus(pdf_bytes: bytes) -> Dict[str, str]:
    """Map item number -> Dell SKU read from the Product Details per-item
    heading block (see comment above)."""
    lines = _extract_pdf_lines(pdf_bytes)
    out: Dict[str, str] = {}
    for i, line in enumerate(lines):
        m = _CHECKOUT_ITEM_QTY_LINE_RE.match(line.strip())
        if not m:
            continue
        item_no = m.group(1)
        if item_no in out:
            continue
        for j in range(i + 1, min(i + 3, len(lines))):
            nxt = lines[j].strip()
            if not nxt:
                continue
            sm = _CHECKOUT_SKU_ONLY_LINE_RE.match(nxt)
            if sm:
                out[item_no] = sm.group(1)
            break
    return out

# Label -> quote_meta key. "company name" is folded into "reseller" below
# (this tool is dedicated to Southcomp Polaris quotes, and this template has
# no separate "Reseller:" label of its own); "sales representative" is
# folded into "quote creator" (Dell's internal owner of the quote).
_CHECKOUT_LABEL_MAP: Dict[str, str] = {
    "company name": "company name",
    "customer name": "customer name",
    "customer number": "customer number",
    "end user": "end user",
    "sales representative": "quote creator",
}


def _dedupe_adjacent_words(line: str) -> str:
    """Collapse immediately-repeated words: 'Quote Quote No.: No.: 123 123' -> 'Quote No.: 123'.

    Scoped to the checkout-confirmation PDF parser only — used nowhere else,
    so it cannot change behavior for any other template.
    """
    words = line.split(" ")
    out: List[str] = []
    for w in words:
        if out and out[-1] == w:
            continue
        out.append(w)
    return " ".join(out)


def _try_extract_checkout_confirmation_pdf(
    pdf_bytes: bytes,
) -> Optional[Tuple[List, Dict, List, str, str, str, float]]:
    """Parse the Dell 'Your quote is ready for purchase' checkout-confirmation PDF.

    Returns (items, metadata, config_rows, quote_ref, date_text, expiry_text,
    consolidation_fee), or None when this template isn't detected so the
    caller falls back to the existing PDF parsing paths.
    """
    raw_lines = _extract_pdf_lines(pdf_bytes)
    if not any(
        "quote is ready for purchase" in l.lower() or "place your order" in l.lower()
        for l in raw_lines
    ):
        return None

    lines = [_dedupe_adjacent_words(l) for l in raw_lines]

    metadata = {"company name": "", "customer name": "", "customer number": "",
                "end user": "", "reseller": "", "quote creator": "", "shipping info": ""}
    quote_ref = date_text = expiry_text = ""
    consolidation_fee = 0.0
    items: List[Tuple] = []
    in_items = False

    for raw_line, line in zip(raw_lines, lines):
        stripped = line.strip()
        low = stripped.lower()
        if not low:
            continue

        if not quote_ref and low.startswith("quote no"):
            m = re.search(r"\d{6,}(?:\.[A-Za-z0-9]+)?", stripped)
            if m:
                quote_ref = m.group(0)
            continue
        if not date_text and low.startswith("quoted on"):
            m = re.search(r"\d{2}/\d{2}/\d{4}", stripped)
            if m:
                date_text = m.group(0)
            continue
        if low.startswith("expires by"):
            m = re.search(r"\d{2}/\d{2}/\d{4}", stripped)
            if m:
                expiry_text = m.group(0)
            continue

        if not in_items:
            matched_label = False
            for label, meta_key in _CHECKOUT_LABEL_MAP.items():
                if low.startswith(label + ":"):
                    val = stripped.split(":", 1)[1].strip()
                    if val:
                        metadata[meta_key] = val
                    matched_label = True
                    break
            if matched_label:
                continue

        if low.startswith("consolidation fee:"):
            # Note the trailing colon: the Product Details tables also contain
            # module rows labelled "Consolidation Fees -" (plural, no colon,
            # followed by an unrelated SKU number) which must NOT match here.
            m = re.search(r"[\d,]+\.?\d*", stripped)
            if m:
                consolidation_fee += _parse_money(m.group(0)) or 0.0
            continue
        if low.startswith("shipping:"):
            fee = _parse_money(stripped.split(":", 1)[1]) or 0.0
            if abs(fee) > 1e-9:
                consolidation_fee += fee
            continue

        if "pricing summary" in low:
            in_items = True
            continue

        if in_items:
            if low.startswith("subtotal:"):
                in_items = False
                continue
            # Item lines are matched on the raw text first: de-duplicating
            # words also collapses "$9,800.20 $9,800.20" (unit price equals
            # subtotal whenever qty is 1) into a single price, which then no
            # longer fits the two-price pattern and the item is lost.
            m = _CHECKOUT_ITEM_LINE_PAT.match(raw_line.strip()) or _CHECKOUT_ITEM_LINE_PAT.match(stripped)
            if m:
                desc_s, qty_s, unit_s, total_s = m.groups()
                qty_val = int(qty_s)
                unit_val = _parse_money(unit_s) or 0.0
                total_val = _parse_money(total_s) or (qty_val * unit_val)
                items.append((desc_s.strip(), qty_val, unit_val, total_val))
            elif items and not _is_price_or_qty_line(stripped):
                old_desc, qty, unit, total = items[-1]
                items[-1] = (f"{old_desc} {stripped}".strip(), qty, unit, total)

    if not items:
        return None

    if metadata.get("company name") and not metadata.get("reseller"):
        metadata["reseller"] = metadata["company name"]

    config_rows = _extract_config_from_pdf_module_table(pdf_bytes)

    # A simple accessory (mouse, keyboard, dock, sleeve, ...) has no
    # component table at all, so it never gets a "Base" row above — but it
    # does have its own real Dell SKU in the Product Details heading. Add it
    # with a blank module name (not "Base") so it's available as an item
    # code but never mistaken for a config field when the description is
    # built.
    items_with_base = {row[0] for row in config_rows if (row[2] or "").strip().lower() == "base"}
    accessory_skus = _extract_checkout_pricing_item_skus(pdf_bytes)
    for idx, item in enumerate(items, start=1):
        item_no = str(idx)
        if item_no in items_with_base:
            continue
        sku = accessory_skus.get(item_no)
        if sku:
            config_rows.append((item_no, "", "", item[0], sku, ""))

    return items, metadata, config_rows, quote_ref, date_text, expiry_text, consolidation_fee


# ==================== CONFIGURATION SHEET ====================

def _find_config_sheet(wb) -> Optional[object]:
    """Exact copy of dell.py _find_configuration_sheet logic."""
    for name in wb.sheetnames:
        if re.sub(r"[^a-z0-9]", "", name.lower().strip()) in {
            "configuration", "config", "configsheet", "configurationsheet",
            "configurationdetails", "configdetails", "productdetails",
        }:
            return wb[name]
    return None


def _find_product_details_anchor(ws) -> Optional[int]:
    """Exact copy from dell.py."""
    max_c = min(ws.max_column, 40)
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if v and "product details" in str(v).lower():
                return r
    return None


def _find_config_table_header(ws, start_row: int = 1, search_rows: int = 30) -> Optional[Tuple[int, Dict[str, int]]]:
    """Exact copy from dell.py."""
    last_row = min(ws.max_row, start_row + search_rows)
    for r in range(start_row, last_row + 1):
        labels: Dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            name = _cell_to_text(ws.cell(r, c).value).lower()
            if not name:
                continue
            if "module" in name and "module" not in labels:
                labels["module"] = c
            if "description" in name and "description" not in labels:
                labels["description"] = c
            normalized_name = re.sub(r"\s+", " ", name.strip())
            if (
                normalized_name in ("sku", "part", "part #", "part#", "part number", "part no", "part no.", "dell part number")
                or ("sku" in normalized_name)
                or ("part" in normalized_name and "number" in normalized_name)
            ) and "sku" not in labels:
                labels["sku"] = c
            if name.strip() in ("qty", "quantity") and "qty" not in labels:
                labels["qty"] = c
        if all(k in labels for k in ("description", "sku")):
            labels.setdefault("module", labels["description"])
            return r, labels
    return None


def _extract_config_rows(ws) -> List[Tuple]:
    """Extract config rows from a dedicated Configuration sheet. Exact copy of dell.py _extract_config_rows_from_configuration_sheet."""
    header_info = _find_config_table_header(ws, 1, search_rows=50)
    if not header_info:
        return []
    header_row, colmap = header_info
    item_col: Optional[int] = None
    for c in range(1, ws.max_column + 1):
        header_text = _cell_to_text(ws.cell(header_row, c).value).lower()
        if header_text in ("item", "item#", "item #", "item no", "item number", "sr. no.", "sr no", "srno", "sr"):
            item_col = c
            break
    rows = []
    current_item = "1"
    has_real_module_col = bool(colmap.get("module")) and colmap.get("module") != colmap.get("description")
    for r in range(header_row + 1, ws.max_row + 1):
        row_text = _row_text(ws, r, 1, ws.max_column)
        if not row_text:
            continue
        if item_col:
            item_value = _cell_to_text(ws.cell(r, item_col).value).strip()
            if item_value:
                current_item = item_value.rstrip(".")
        module = _cell_to_text(ws.cell(r, colmap.get("module", 0)).value) if has_real_module_col else ""
        description = _cell_to_text(ws.cell(r, colmap.get("description", 0)).value)
        sku = _cell_to_text(ws.cell(r, colmap.get("sku", 0)).value)
        qty = _cell_to_text(ws.cell(r, colmap.get("qty", 0)).value)
        if not has_real_module_col and description and not any([sku, qty]):
            module = description
            description = ""
        if not any([module, description, sku, qty]):
            continue
        rows.append((current_item, "", module, description, sku, qty))
    return rows


def _extract_all_config_rows(ws) -> List[Tuple]:
    """Exact copy of dell.py _extract_all_config_rows — handles Product Details per-item tables."""
    anchor = _find_product_details_anchor(ws)
    if not anchor:
        return []

    rows: List[Tuple] = []
    r = anchor + 1
    max_col = min(ws.max_column, 40)

    def _clean_heading_text(text: str) -> str:
        return re.sub(r"\s+\d+(\.\d+)?\s+\$?[\d,\.]+\s+\$?[\d,\.]+$", "", text).strip()

    def _is_table_stop(text: str) -> bool:
        low = text.lower()
        return (
            _is_price_or_qty_line(text)
            or "estimated delivery" in low
            or "subtotal" in low
            or "total" in low
            or "ship to" in low
        )

    def _extract_item_heading(row_idx: int) -> Optional[Tuple[str, str]]:
        item_marker = _cell_to_text(ws.cell(row_idx, 1).value)
        if not re.match(r"^\d+\.$", item_marker):
            return None
        heading = _cell_to_text(ws.cell(row_idx, 2).value)
        if not heading:
            heading = _row_text(ws, row_idx, 1, max_col)
        return item_marker.rstrip("."), _clean_heading_text(heading)

    def _find_next_item_row(start_row: int) -> Optional[int]:
        for row_idx in range(start_row, ws.max_row + 1):
            if _extract_item_heading(row_idx):
                return row_idx
        return None

    item_counter = 0
    while r <= ws.max_row:
        item_info = _extract_item_heading(r)
        if not item_info:
            r += 1
            continue

        _source_item_number, current_heading = item_info
        item_counter += 1
        current_item = str(item_counter)
        next_item_row = _find_next_item_row(r + 1)
        search_end = (next_item_row - 1) if next_item_row is not None else ws.max_row

        header_info = None
        scan_row = r + 1
        while scan_row <= search_end:
            maybe_header = _find_config_table_header(ws, scan_row, search_rows=0)
            if maybe_header:
                header_info = maybe_header
                break
            scan_row += 1

        if not header_info:
            r = next_item_row if next_item_row is not None else ws.max_row + 1
            continue

        header_row, colmap = header_info
        data_row = header_row + 1
        has_real_module_col = bool(colmap.get("module")) and colmap.get("module") != colmap.get("description")

        while data_row <= search_end and not _row_text(ws, data_row, 1, max_col):
            data_row += 1

        blank_streak = 0
        while data_row <= search_end:
            row_text_all = _row_text(ws, data_row, 1, max_col)
            if not row_text_all:
                blank_streak += 1
                if blank_streak >= 2:
                    break
                data_row += 1
                continue
            blank_streak = 0
            mod = _cell_to_text(ws.cell(data_row, colmap.get("module", 0)).value) if has_real_module_col else ""
            desc = _cell_to_text(ws.cell(data_row, colmap.get("description", 0)).value)
            sku = _cell_to_text(ws.cell(data_row, colmap.get("sku", 0)).value)
            qty = _cell_to_text(ws.cell(data_row, colmap.get("qty", 0)).value)
            # A row with a real SKU is a config row even if its text mentions a
            # currency — e.g. power cords described as "... 10A (EUR)".
            has_real_sku = bool(sku.strip()) and sku.strip().lower() != "sku"
            if _is_table_stop(row_text_all) and not has_real_sku:
                data_row += 1
                continue
            if not has_real_module_col and desc and not any([sku, qty]):
                mod = desc
                desc = ""
            if not any([mod, desc, sku, qty]):
                break
            rows.append((current_item, current_heading, mod, desc, sku, qty))
            data_row += 1

        r = next_item_row if next_item_row is not None else ws.max_row + 1

    # Merge 2-line fragmented rows (common in Dell exports)
    cleaned: List[Tuple] = []
    i = 0
    while i < len(rows):
        item, head, mod, desc, sku, qty = rows[i]
        mod, desc = mod.strip(), desc.strip()
        if i + 1 < len(rows):
            ni, nh, nmod, ndesc, nsku, nqty = rows[i + 1]
            if ni == item and nh == head:
                if desc == "" and ndesc == "" and ":" not in mod and ":" not in nmod:
                    mod = f"{mod} {nmod}".strip()
                    i += 1
        cleaned.append((item, head, mod, desc, sku, qty))
        i += 1

    return cleaned


_HEADING_SKU_RE = re.compile(r"\((\d{3,4}-[A-Za-z0-9]{2,10})\)\s*$")


def _extract_product_detail_heading_skus(ws) -> Dict[str, str]:
    """{item_no: SKU} from the Product Details per-item headings, e.g.
    "3." | "12Gb HD-Mini SAS cable, 2m, Customer Kit [...] (470-ABDR)".

    _extract_all_config_rows() only reads headings that have a component
    table under them, so a plain accessory's SKU is otherwise lost. Items are
    numbered the same way (one per "N." heading, in order).
    """
    anchor = _find_product_details_anchor(ws)
    if not anchor:
        return {}
    out: Dict[str, str] = {}
    item_counter = 0
    for r in range(anchor + 1, ws.max_row + 1):
        if not re.match(r"^\d+\.$", _cell_to_text(ws.cell(r, 1).value)):
            continue
        item_counter += 1
        m = _HEADING_SKU_RE.search(_cell_to_text(ws.cell(r, 2).value))
        if m:
            out[str(item_counter)] = m.group(1)
    return out


def _extract_consolidation_fee(ws) -> float:
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            if not re.fullmatch(r"consolidation fees?\s*:?", cell.value.strip().lower()):
                continue
            for next_col in range(ws.max_column, cell.column, -1):
                nv = ws.cell(cell.row, next_col).value
                if nv in (None, ""):
                    continue
                parsed = _parse_money(nv)
                if parsed is not None:
                    return 0.0 if abs(parsed) < 1e-9 else parsed
    return 0.0


def _extract_shipping_fee(ws) -> float:
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            if not re.fullmatch(r"shipping(?:\s+(?:charge|charges|cost))?\s*:?", cell.value.strip().lower()):
                continue
            for next_col in range(cell.column + 1, ws.max_column + 1):
                nv = ws.cell(cell.row, next_col).value
                if nv in (None, ""):
                    continue
                fee = _parse_money(nv) or 0.0
                return 0.0 if abs(fee) < 1e-9 else fee
    return 0.0


def _extract_part_numbers(ws) -> Dict[str, str]:
    """Return {item_no: part_number}."""
    out: Dict[str, str] = {}
    sku_pat = re.compile(r"\b[A-Z]{2,4}[0-9]{2,}[A-Z0-9\-]*\b|"
                         r"\b[0-9]{3,}[A-Z]{2,}[0-9A-Z\-]*\b", re.IGNORECASE)
    current_item = "1"
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column, 5) + 1):
            v = _cell_to_text(ws.cell(r, c).value).strip()
            if re.match(r"^\d+\.$", v):
                current_item = v.rstrip(".")
                # Look for a SKU in the rest of the row
                for rc in range(c + 1, min(ws.max_column, 20) + 1):
                    rv = _cell_to_text(ws.cell(r, rc).value).strip()
                    m = sku_pat.match(rv)
                    if m:
                        out.setdefault(current_item, m.group(0))
                        break
    return out


# ==================== DOCX EXTRACTION (Dell "Prospective Customer Quote" Word form) ====================

_DOCX_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Labels from the optional "Quote details" key/value block -> quote_meta-ish keys.
_DOCX_DETAIL_LABELS = {
    "partner name": "reseller",
    "dell customer number": "customer number",
}


def _is_docx(input_bytes: bytes) -> bool:
    try:
        with zipfile.ZipFile(BytesIO(input_bytes)) as z:
            return "word/document.xml" in z.namelist()
    except Exception:
        return False


def _docx_cell_text(tc) -> str:
    lines = ["".join(t.text or "" for t in p.iter(_DOCX_W + "t")) for p in tc.findall(_DOCX_W + "p")]
    return "\n".join(lines).strip()


def _docx_tables(input_bytes: bytes) -> List[List[List[str]]]:
    with zipfile.ZipFile(BytesIO(input_bytes)) as z:
        xml_bytes = z.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    body = root.find(_DOCX_W + "body")
    tables: List[List[List[str]]] = []
    if body is None:
        return tables
    for tbl in body.findall(_DOCX_W + "tbl"):
        rows = [[_docx_cell_text(tc) for tc in tr.findall(_DOCX_W + "tc")] for tr in tbl.findall(_DOCX_W + "tr")]
        tables.append(rows)
    return tables


def _docx_norm_label(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower()).rstrip(":")


def _docx_find_item_columns(header_row: List[str]) -> Optional[Dict[str, int]]:
    cols: Dict[str, int] = {}
    for i, cell in enumerate(header_row):
        name = _docx_norm_label(cell)
        if "part" in name and "number" in name:
            cols.setdefault("sku", i)
        elif name == "description":
            cols.setdefault("description", i)
        elif name in ("qty", "quantity"):
            cols.setdefault("qty", i)
        elif "unit price" in name:
            cols.setdefault("unit", i)
        elif name == "total" or "total price" in name:
            cols.setdefault("total", i)
    if all(k in cols for k in ("description", "qty", "unit", "total")):
        return cols
    return None


def _docx_clean_description(text: str) -> str:
    """Strip the 'Unit Shipping Price : $' note some exports glue onto the description."""
    text = re.split(r"unit shipping price", text, flags=re.IGNORECASE)[0]
    text = re.sub(r"\s*\n\s*", " ", text).strip()
    return re.sub(r"\s{2,}", " ", text)


def _docx_consume_item_rows(
    rows: List[List[str]],
    cols: Dict[str, int],
    items: List[Tuple],
    part_numbers: Dict[str, str],
    price_text_blob: List[str],
) -> None:
    expected_cols = max(cols.values()) + 1
    for row in rows:
        # Merged summary/total rows collapse to fewer <w:tc> cells than the item rows.
        if len(row) < expected_cols:
            continue
        desc = _docx_clean_description(row[cols["description"]])
        qty_raw = row[cols["qty"]].strip()
        if not desc or not qty_raw or "total" in _docx_norm_label(row[0]):
            continue
        try:
            qty_val = int(_parse_money(qty_raw) or 0)
        except Exception:
            qty_val = 0
        if qty_val <= 0:
            continue
        unit_raw = row[cols["unit"]].strip()
        total_raw = row[cols["total"]].strip()
        unit_val = _parse_money(unit_raw) or 0.0
        total_val = _parse_money(total_raw) or (unit_val * qty_val)
        price_text_blob.append(unit_raw)
        price_text_blob.append(total_raw)
        items.append((desc, qty_val, unit_val, total_val))
        if "sku" in cols:
            sku = row[cols["sku"]].replace("\n", "").strip()
            if sku:
                part_numbers[str(len(items))] = sku


def _extract_southcomp_docx(
    input_bytes: bytes,
) -> Tuple[List[Tuple], Dict[str, str], str, str, str, Dict[str, str]]:
    """Parse a Dell 'Prospective Customer Quote' Word order form.

    Returns (items, quote_meta, quote_ref, date_text, source_currency, part_numbers).
    """
    tables = _docx_tables(input_bytes)

    quote_ref = ""
    date_text = ""
    bill_to = ""
    ship_to = ""
    detail_kv: Dict[str, str] = {}
    items: List[Tuple] = []
    part_numbers: Dict[str, str] = {}
    price_text_blob: List[str] = []
    pending_cols: Optional[Dict[str, int]] = None

    for rows in tables:
        if not rows or not rows[0]:
            continue
        first_row = rows[0]
        first_row_norm = [_docx_norm_label(c) for c in first_row]

        if not quote_ref or not date_text:
            for cell in first_row:
                if not quote_ref:
                    m = re.search(r"quote\s*#\s*:?\s*([A-Za-z0-9\-]+)", cell, re.IGNORECASE)
                    if m:
                        quote_ref = m.group(1)
                if not date_text:
                    m = re.search(r"\bdate\s*:\s*(.+)", cell, re.IGNORECASE)
                    if m:
                        date_text = m.group(1).strip()

        if (len(first_row_norm) >= 2 and first_row_norm[0].startswith("bill to")
                and first_row_norm[1].startswith("ship to")):
            if len(rows) >= 2:
                bill_to = rows[1][0] if len(rows[1]) > 0 else ""
                ship_to = rows[1][1] if len(rows[1]) > 1 else ""
            continue

        item_cols = _docx_find_item_columns(first_row)
        if item_cols is not None:
            if len(rows) > 1:
                _docx_consume_item_rows(rows[1:], item_cols, items, part_numbers, price_text_blob)
            else:
                pending_cols = item_cols
            continue

        if pending_cols is not None:
            _docx_consume_item_rows(rows, pending_cols, items, part_numbers, price_text_blob)
            pending_cols = None
            continue

        for row in rows:
            if len(row) != 2:
                continue
            label = _docx_norm_label(row[0])
            value = row[1].strip()
            if label in _DOCX_DETAIL_LABELS and value:
                detail_kv[_DOCX_DETAIL_LABELS[label]] = value

    source_currency = "EUR" if "€" in " ".join(price_text_blob) else "USD"

    bill_to_lines = bill_to.split("\n") if bill_to else []
    quote_meta: Dict[str, str] = {
        "customer name": bill_to_lines[0].strip() if bill_to_lines else "",
        "company name": (bill_to_lines[1].strip() if len(bill_to_lines) > 1 else "") or detail_kv.get("reseller", ""),
        "end user": ship_to,
        "reseller": detail_kv.get("reseller", ""),
    }

    return items, quote_meta, quote_ref, date_text, source_currency, part_numbers


# ==================== QUOTE GENERATION ====================

def _build_quote_workbook(
    items: List[Tuple],
    config_rows: List[Tuple],
    quote_ref: str,
    date_text: str,
    expiry_text: str,
    quote_meta: Dict[str, str],
    currency_code: str,
    exchange_rate: float,
    consolidation_fee: float,
    margin_percent: float,
    is_pdf: bool,
    part_numbers: Optional[Dict[str, str]] = None,
    include_config_sheet: bool = True,
) -> bytes:
    """Build the EUR-style quote workbook (Quote + optional Configuration sheet)."""
    currency_code = currency_code.upper()
    conversion_rate = exchange_rate if currency_code == "EUR" else 1.0
    margin_decimal = margin_percent / 100.0

    # Convert items to selected currency; keep originals for the USD helper columns
    original_usd_items = list(items)
    if conversion_rate != 1.0:
        items = [
            (desc, qty, (unit or 0.0) * conversion_rate,
             ((sub or (qty * (unit or 0.0))) * conversion_rate))
            for desc, qty, unit, sub in items
        ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Quote"
    ws.sheet_view.showGridLines = False

    include_part_number = bool(part_numbers)

    # Column layout (EUR-style: always has helper + USD columns)
    if include_part_number:
        desc_col, qty_col, unit_col, total_col = "C", "D", "E", "F"
        helper_unit_col, helper_fee_col = "G", "H"
        usd_unit_col, usd_total_col = "I", "J"
        helper_margin_col = "K"
    else:
        desc_col, qty_col, unit_col, total_col = "B", "C", "D", "E"
        helper_unit_col, helper_fee_col = "F", "G"
        usd_unit_col, usd_total_col = "H", "I"
        helper_margin_col = "J"

    # Styling constants
    section_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    helper_header_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    helper_body_fill = PatternFill(start_color="FCE5E5", end_color="FCE5E5", fill_type="solid")
    header_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
    helper_font = Font(bold=True, color="9C0006")
    header_font = Font(bold=True, color="000000")
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    def _style_section_title(addr: str) -> None:
        ws[addr].font = Font(bold=True, color="1F497D")
        ws[addr].alignment = Alignment(horizontal="left", vertical="center")
        ws[addr].fill = section_fill
        ws[addr].border = Border(
            left=Side(style="thin", color="9FBAD0"),
            right=Side(style="thin", color="9FBAD0"),
            top=Side(style="thin", color="9FBAD0"),
            bottom=Side(style="thin", color="9FBAD0"),
        )

    # --- Column widths ---
    description_width = min(max(44, int(max((len(_cell_to_text(it[0])) for it in items), default=0) * 0.55)), 68)
    widths: Dict[str, float] = {"A": 11}
    if include_part_number:
        widths.update({"B": 16, "C": min(max(42, description_width), 56), "D": 8, "E": 15, "F": 17, "G": 17, "H": 12})
    else:
        widths.update({"B": min(max(42, description_width), 56), "C": 8, "D": 15, "E": 18, "F": 17, "G": 12})
    widths[helper_fee_col] = 12
    widths[usd_unit_col] = 18
    widths[usd_total_col] = 18
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Row heights
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 26
    for rr in range(3, 11):
        ws.row_dimensions[rr].height = 20

    # --- Logo ---
    ws.merge_cells("A1:H2")
    _add_logo(ws, anchor="A1", width=780, height=52)

    # --- Address block (French / Southcomp) ---
    def _write_address(start_row: int, end_row: int, lines: List[str], merge: bool = True) -> None:
        if merge:
            rng = f"A{start_row}:D{end_row}"
            ws.merge_cells(rng)
            ws.unmerge_cells(rng)
        for offset, text in enumerate(lines):
            cell = ws.cell(row=start_row + offset, column=1, value=text)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if merge:
            ws.merge_cells(f"A{start_row}:D{end_row}")

    _write_address(5, 8, [
        "14, rue du Bas Marin",
        "94537 Orly cedex - France",
        "DL:     +33 1 49 79 42 24",
        "Fax:   +33 1 49 79 45 33",
    ], merge=False)
    address_end_row = 8
    for addr in ("A5", "A6", "A7", "A8"):
        ws[addr].font = Font(bold=True, size=11, color="1F497D")
        ws[addr].alignment = Alignment(horizontal="left", vertical="center")

    # --- Quote Summary ---
    has_expiry = bool(expiry_text)
    summary_title_row = 9
    ws.merge_cells(f"A{summary_title_row}:D{summary_title_row}")
    ws[f"A{summary_title_row}"] = "Quote Summary"
    _style_section_title(f"A{summary_title_row}")

    summary_rows = [(summary_title_row + 1, "Quote Ref", quote_ref),
                    (summary_title_row + 2, "Date", date_text)]
    if has_expiry:
        summary_rows.append((summary_title_row + 3, "Expires By", expiry_text))

    for row_idx, label, value in summary_rows:
        ws[f"A{row_idx}"] = label
        ws[f"A{row_idx}"].font = Font(bold=True, color="1F497D")
        ws[f"A{row_idx}"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=4)
        ws[f"B{row_idx}"] = value
        ws[f"B{row_idx}"].alignment = Alignment(horizontal="left", vertical="center")
        if label == "Expires By":
            ws[f"B{row_idx}"].font = Font(bold=True)

    customer_title_row = (summary_title_row + 4) if has_expiry else (summary_title_row + 3)

    # --- Quote metadata ---
    if is_pdf:
        meta_rows = [
            ("End Customer:", quote_meta.get("end user", "")),
            ("Reseller:", quote_meta.get("reseller", "")),
            ("Quote Creator:", quote_meta.get("quote creator", "")),
        ]
        if quote_meta.get("shipping info"):
            meta_rows.append(("Shipping Information:", quote_meta.get("shipping info", "")))
    else:
        meta_rows = [
            ("Company Name:", quote_meta.get("company name", "")),
            ("Customer Name:", quote_meta.get("customer name", "")),
            ("End User:", quote_meta.get("end user", "")),
            ("Reseller:", quote_meta.get("reseller", "")),
        ]

    ws.merge_cells(f"A{customer_title_row}:D{customer_title_row}")
    ws[f"A{customer_title_row}"] = "Customer Information"
    _style_section_title(f"A{customer_title_row}")

    for i, (label, value) in enumerate(meta_rows, start=1):
        row_idx = customer_title_row + i
        ws[f"A{row_idx}"] = label
        ws[f"A{row_idx}"].font = Font(bold=True, color="1F497D")
        ws[f"A{row_idx}"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=4)
        ws[f"B{row_idx}"] = value
        ws[f"B{row_idx}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        explicit_newlines = value.count("\n")
        text_len = len(str(value))
        estimated_lines = max(1, explicit_newlines + 1 + max(0, text_len // 32))
        ws.row_dimensions[row_idx].height = max(ws.row_dimensions[row_idx].height or 20, min(estimated_lines, 12) * 18)

    # --- Recalculate helper row positions ---
    last_meta_row = customer_title_row + len(meta_rows)
    helper_value_row = last_meta_row + 1
    helper_aux_row = helper_value_row + 1

    # F/G col of helper row: consolidation fee total (editable by user in Excel)
    ws[f"{helper_unit_col}{helper_value_row}"] = consolidation_fee
    ws[f"{helper_unit_col}{helper_value_row}"].font = helper_font
    ws[f"{helper_unit_col}{helper_value_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"{helper_unit_col}{helper_value_row}"].fill = helper_body_fill
    ws[f"{helper_unit_col}{helper_value_row}"].border = border_thin

    # Marge col of helper row: margin % as a static decimal — NO circular formula
    ws[f"{helper_margin_col}{helper_value_row}"] = margin_decimal
    ws[f"{helper_margin_col}{helper_value_row}"].number_format = "0.00%"
    ws[f"{helper_margin_col}{helper_value_row}"].font = helper_font
    ws[f"{helper_margin_col}{helper_value_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"{helper_margin_col}{helper_value_row}"].fill = helper_body_fill
    ws[f"{helper_margin_col}{helper_value_row}"].border = border_thin

    # --- Table header ---
    header_row = helper_aux_row + 1
    ws[f"A{header_row}"] = "N°"
    if include_part_number:
        ws[f"B{header_row}"] = "N° de pièce"
    ws[f"{desc_col}{header_row}"] = "Description"
    ws[f"{qty_col}{header_row}"] = "Qté"
    ws[f"{unit_col}{header_row}"] = "Prix unitaire"
    ws[f"{total_col}{header_row}"] = "Prix total"
    ws[f"{helper_unit_col}{header_row}"] = "Prix unitaire d'origine"
    ws[f"{helper_fee_col}{header_row}"] = "Fees"
    ws[f"{usd_unit_col}{header_row}"] = "Unit Price USD original"
    ws[f"{usd_total_col}{header_row}"] = "Total Price USD original"
    ws[f"{helper_margin_col}{header_row}"] = "Marge"

    header_cells = [f"A{header_row}", f"{desc_col}{header_row}", f"{qty_col}{header_row}",
                    f"{unit_col}{header_row}", f"{total_col}{header_row}",
                    f"{helper_unit_col}{header_row}", f"{helper_fee_col}{header_row}",
                    f"{usd_unit_col}{header_row}", f"{usd_total_col}{header_row}",
                    f"{helper_margin_col}{header_row}"]
    if include_part_number:
        header_cells.insert(1, f"B{header_row}")
    helper_header_cells = (f"{helper_unit_col}{header_row}", f"{helper_fee_col}{header_row}",
                           f"{usd_unit_col}{header_row}", f"{usd_total_col}{header_row}",
                           f"{helper_margin_col}{header_row}")
    for addr in header_cells:
        ws[addr].fill = helper_header_fill if addr in helper_header_cells else header_fill
        ws[addr].font = header_font
        ws[addr].alignment = Alignment(horizontal="center", vertical="center")
        ws[addr].border = border_thin
    ws.row_dimensions[header_row].height = 20

    # --- Data rows ---
    row_ptr = header_row + 1
    sr_no = 1
    currency_fmt = CURRENCY_FORMATS.get(currency_code, f'"{currency_code}" #,##0.00')
    usd_fmt = CURRENCY_FORMATS["USD"]
    yellow = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    total_cells = []

    for idx, (desc_text, qty_val, unit_val, subtotal_val) in enumerate(items):
        original_usd_unit = original_usd_items[idx][2] if idx < len(original_usd_items) else None

        ws[f"A{row_ptr}"] = sr_no
        if include_part_number and part_numbers:
            ws[f"B{row_ptr}"] = _sanitize_excel_text(part_numbers.get(str(sr_no), ""))
        ws[f"{desc_col}{row_ptr}"] = _sanitize_excel_text(desc_text)
        ws[f"{qty_col}{row_ptr}"] = qty_val or 0
        unit_val = unit_val or 0.0

        # "Prix unitaire d'origine" — static original cost price (in output currency, pre-margin)
        orig_helper = f"{helper_unit_col}{row_ptr}"
        ws[orig_helper] = unit_val
        ws[orig_helper].font = helper_font
        ws[orig_helper].fill = helper_body_fill
        ws[orig_helper].number_format = currency_fmt
        ws[orig_helper].border = border_thin
        ws[orig_helper].alignment = Alignment(horizontal="center", vertical="center")

        # "Fees" per unit — static 0, user can edit in Excel
        fee_helper = f"{helper_fee_col}{row_ptr}"
        ws[fee_helper] = 0
        ws[fee_helper].font = helper_font
        ws[fee_helper].fill = helper_body_fill
        ws[fee_helper].number_format = currency_fmt
        ws[fee_helper].border = border_thin
        ws[fee_helper].alignment = Alignment(horizontal="center", vertical="center")

        # Per-item margin cell: defaults to J17 but user can override individually
        ws[f"{helper_margin_col}{row_ptr}"] = f"={helper_margin_col}${helper_value_row}"
        ws[f"{helper_margin_col}{row_ptr}"].number_format = "0.00%"
        ws[f"{helper_margin_col}{row_ptr}"].font = helper_font
        ws[f"{helper_margin_col}{row_ptr}"].fill = helper_body_fill
        ws[f"{helper_margin_col}{row_ptr}"].border = border_thin
        ws[f"{helper_margin_col}{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")

        # "Prix unitaire" — selling price = (original + fees) / (1 - this row's margin),
        # so the margin is a share of the selling price (same convention as the Dell tool)
        ws[f"{unit_col}{row_ptr}"] = f"=({helper_unit_col}{row_ptr}+{helper_fee_col}{row_ptr})/(1-{helper_margin_col}{row_ptr})"
        ws[f"{unit_col}{row_ptr}"].number_format = currency_fmt
        ws[f"{unit_col}{row_ptr}"].border = border_thin
        ws[f"{unit_col}{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")

        # "Prix total"
        total_addr = f"{total_col}{row_ptr}"
        ws[total_addr] = f"={unit_col}{row_ptr}*{qty_col}{row_ptr}"
        ws[total_addr].number_format = currency_fmt
        ws[total_addr].border = border_thin
        ws[total_addr].alignment = Alignment(horizontal="center", vertical="center")
        total_cells.append(total_addr)

        # USD original columns (always USD values from original BOQ)
        usd_unit = original_usd_unit or (unit_val / conversion_rate if conversion_rate and conversion_rate != 1.0 else unit_val)
        ws[f"{usd_unit_col}{row_ptr}"] = usd_unit
        ws[f"{usd_unit_col}{row_ptr}"].number_format = usd_fmt
        ws[f"{usd_unit_col}{row_ptr}"].fill = helper_body_fill
        ws[f"{usd_unit_col}{row_ptr}"].border = border_thin
        ws[f"{usd_unit_col}{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"{usd_total_col}{row_ptr}"] = usd_unit * (qty_val or 0)
        ws[f"{usd_total_col}{row_ptr}"].number_format = usd_fmt
        ws[f"{usd_total_col}{row_ptr}"].fill = helper_body_fill
        ws[f"{usd_total_col}{row_ptr}"].border = border_thin
        ws[f"{usd_total_col}{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")

        for addr in [f"A{row_ptr}", f"{desc_col}{row_ptr}", f"{qty_col}{row_ptr}"]:
            ws[addr].border = border_thin
            ws[addr].alignment = Alignment(horizontal="center" if addr == f"A{row_ptr}" else "left", vertical="center", wrap_text=True)
        if include_part_number:
            ws[f"B{row_ptr}"].border = border_thin
            ws[f"B{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")

        ws[f"A{row_ptr}"].fill = yellow
        if include_part_number:
            ws[f"B{row_ptr}"].fill = yellow
        ws[f"{desc_col}{row_ptr}"].fill = yellow
        ws[f"{qty_col}{row_ptr}"].fill = yellow
        ws[f"{qty_col}{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")

        row_ptr += 1
        sr_no += 1

    # --- Total row ---
    total_label_col = "C" if include_part_number else "B"
    if include_part_number:
        ws.merge_cells(start_row=row_ptr, start_column=3, end_row=row_ptr, end_column=5)
    else:
        ws.merge_cells(start_row=row_ptr, start_column=2, end_row=row_ptr, end_column=4)
    ws[f"{total_label_col}{row_ptr}"] = "Prix total"
    ws[f"{total_label_col}{row_ptr}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"{total_label_col}{row_ptr}"].font = Font(bold=True, color="1F497D")
    ws[f"{total_col}{row_ptr}"] = f"=SUM({','.join(total_cells)})" if total_cells else 0
    ws[f"{total_col}{row_ptr}"].number_format = currency_fmt
    ws[f"{total_col}{row_ptr}"].font = Font(bold=True, color="1F497D")
    ws[f"{total_col}{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"{total_col}{row_ptr}"].border = border_thin
    ws[f"{helper_unit_col}{row_ptr}"].fill = helper_body_fill
    ws[f"{helper_margin_col}{row_ptr}"].fill = helper_body_fill
    ws[f"{helper_unit_col}{row_ptr}"].border = border_thin
    ws[f"{helper_margin_col}{row_ptr}"].border = border_thin

    # --- Configuration sheet (identical layout to dell.py AED output) ---
    if include_config_sheet:
        ws2 = wb.create_sheet("Configuration")
        ws2.sheet_view.showGridLines = False

        # Show SKU column only when at least one config row carries a real SKU value
        show_sku_col = any(
            len(row) >= 5 and str(row[4]).strip()
            for row in config_rows
        )

        ws2.column_dimensions["A"].width = 22   # Item #
        ws2.column_dimensions["B"].width = 70   # Module
        ws2.column_dimensions["C"].width = 100  # Description
        if show_sku_col:
            ws2.column_dimensions["D"].width = 20  # SKU
            ws2.column_dimensions["E"].width = 10  # Qty
        else:
            ws2.column_dimensions["D"].width = 10  # Qty

        last_col = 5 if show_sku_col else 4     # numeric index of last column
        last_col_letter = "E" if show_sku_col else "D"
        data_cols = ("A", "B", "C", "D", "E") if show_sku_col else ("A", "B", "C", "D")

        title_fill2 = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        section_fill2 = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        thin_gray = Border(
            left=Side(style="thin", color="DDDDDD"),
            right=Side(style="thin", color="DDDDDD"),
            top=Side(style="thin", color="DDDDDD"),
            bottom=Side(style="thin", color="DDDDDD"),
        )
        hdr_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        # Table header
        r2 = 1
        if show_sku_col:
            header_cols = (("A", "Item #"), ("B", "Module"), ("C", "Description"), ("D", "SKU"), ("E", "Qty"))
        else:
            header_cols = (("A", "Item #"), ("B", "Module"), ("C", "Description"), ("D", "Qty"))
        for col, label in header_cols:
            ws2[f"{col}{r2}"] = label
            ws2[f"{col}{r2}"].font = Font(bold=True)
            ws2[f"{col}{r2}"].fill = title_fill2
            ws2[f"{col}{r2}"].alignment = Alignment(horizontal="center", vertical="center")
            ws2[f"{col}{r2}"].border = hdr_border
        ws2.row_dimensions[r2].height = 20
        r2 += 1

        # Group config rows by item number
        config_by_item: Dict[str, List] = {}
        for row in config_rows:
            config_by_item.setdefault(row[0], []).append(row)

        # Item descriptions for headings (from the items list)
        original_descs = [_sanitize_excel_text(it[0]) for it in original_usd_items]

        total_items = max(len(original_descs), len(config_by_item)) if (original_descs or config_by_item) else 0
        for idx in range(1, total_items + 1):
            item_key = str(idx)
            rows_for_item = config_by_item.get(item_key, [])
            heading = original_descs[idx - 1] if idx - 1 < len(original_descs) else f"Item {idx}"

            # "Item N" row — merged across all columns
            ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=last_col)
            ws2[f"A{r2}"] = f"Item {idx}"
            ws2[f"A{r2}"].font = Font(bold=True, color="1F497D")
            ws2[f"A{r2}"].alignment = Alignment(horizontal="left", vertical="center")
            r2 += 1

            # Item description heading — merged across all columns
            ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=last_col)
            ws2[f"A{r2}"] = heading
            ws2[f"A{r2}"].font = Font(italic=True, color="1F497D")
            ws2[f"A{r2}"].alignment = Alignment(horizontal="left", vertical="center")
            r2 += 1

            if not rows_for_item:
                ws2.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=last_col)
                ws2[f"B{r2}"] = "(No configuration details found for this item)"
                ws2[f"B{r2}"].font = Font(italic=True, color="7F7F7F")
                ws2[f"B{r2}"].alignment = Alignment(horizontal="left", vertical="center")
                for col in data_cols:
                    ws2[f"{col}{r2}"].border = thin_gray
                r2 += 1
            else:
                for row_data in rows_for_item:
                    _, _, module, dsc, sku, qty = (row_data + ("", "", "", ""))[:6]
                    # Section header row: module name only, no description or SKU
                    if module and not dsc and not sku:
                        ws2[f"A{r2}"] = ""
                        ws2.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=last_col)
                        ws2[f"B{r2}"] = _sanitize_excel_text(module)
                        ws2[f"B{r2}"].font = Font(bold=True, color="1F1F1F")
                        ws2[f"B{r2}"].fill = section_fill2
                        ws2[f"B{r2}"].alignment = Alignment(horizontal="left", vertical="center")
                        for col in data_cols:
                            ws2[f"{col}{r2}"].border = thin_gray
                        r2 += 1
                        continue

                    ws2[f"A{r2}"] = ""
                    ws2[f"B{r2}"] = _sanitize_excel_text(module)
                    ws2[f"C{r2}"] = _sanitize_excel_text(dsc)
                    if show_sku_col:
                        ws2[f"D{r2}"] = _sanitize_excel_text(sku)
                        ws2[f"E{r2}"] = _sanitize_excel_text(qty)
                    else:
                        ws2[f"D{r2}"] = _sanitize_excel_text(qty)
                    for col in data_cols:
                        ws2[f"{col}{r2}"].alignment = Alignment(vertical="top", wrap_text=True)
                        ws2[f"{col}{r2}"].border = thin_gray
                    r2 += 1

            r2 += 1  # blank gap between items

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# ==================== PUBLIC API ====================

def describe_input_kind(input_bytes: bytes) -> str:
    """Return a short label for which extraction path an input will take (for usage tracking)."""
    if input_bytes.lstrip().startswith(b"%PDF"):
        return "pdf"
    if _is_docx(input_bytes):
        return "docx"
    try:
        wb = openpyxl.load_workbook(BytesIO(input_bytes), data_only=True)
        ws = wb.active
    except Exception:
        return "excel"
    if _is_qar_report(ws):
        return "qar"
    if _find_grouped_header(ws) is not None:
        return "boq_grouped"
    return "boq_generic"


def extract_quote_source_data(input_bytes: bytes, exchange_rate: float = 0.0) -> Dict[str, object]:
    """Extract items, config_rows and metadata from any supported Dell quote
    input (PDF, Excel BOQ, or Word order form).

    This is the exact dispatch logic generate_southcomp_quote() has always
    used, factored out unchanged so a second tool (item-code/description
    extraction) can reuse the same parsing without duplicating — or
    risking drifting from — it.

    exchange_rate (EUR/USD) is only used to normalize a EUR-priced Word form
    back to USD; callers that don't need prices can leave it at 0.
    """
    is_pdf = input_bytes.lstrip().startswith(b"%PDF")
    is_docx = not is_pdf and _is_docx(input_bytes)
    items: List[Tuple] = []
    config_rows: List[Tuple] = []
    quote_ref = date_text = expiry_text = ""
    quote_meta: Dict[str, str] = {}
    consolidation_fee = 0.0
    part_numbers: Dict[str, str] = {}
    item_quote_refs: Dict[str, str] = {}
    heading_skus: Dict[str, str] = {}

    if is_docx:
        items, quote_meta, quote_ref, date_text, source_currency, part_numbers = _extract_southcomp_docx(input_bytes)
        if source_currency == "EUR" and exchange_rate:
            # The Word form's own prices are already EUR; normalize to the USD baseline
            # the rest of the pipeline expects so downstream conversion isn't doubled up.
            items = [
                (desc, qty, (unit or 0.0) / exchange_rate, (total or 0.0) / exchange_rate)
                for desc, qty, unit, total in items
            ]
    elif is_pdf:
        premier_result = _try_extract_premier_pricing_summary_pdf(input_bytes)
        checkout_result = None if premier_result is not None else _try_extract_checkout_confirmation_pdf(input_bytes)
        if premier_result is not None:
            items, quote_meta, config_rows, quote_ref, date_text, expiry_text, consolidation_fee = premier_result
        elif checkout_result is not None:
            items, quote_meta, config_rows, quote_ref, date_text, expiry_text, consolidation_fee = checkout_result
        else:
            items, raw_meta, config_rows, quote_ref, date_text, expiry_text, consolidation_fee = _extract_items_pdf(input_bytes)
            config_rows = _extract_config_from_pdf(input_bytes)
            pos_meta = _extract_pdf_metadata_by_position(input_bytes)
            quote_meta = raw_meta
            if pos_meta.get("quote_creator"):
                quote_meta["quote creator"] = pos_meta["quote_creator"]
            if pos_meta.get("end_user"):
                quote_meta["end user"] = pos_meta["end_user"]
            if pos_meta.get("reseller"):
                quote_meta["reseller"] = pos_meta["reseller"]
            if pos_meta.get("quote_name") and not quote_meta.get("company name"):
                quote_meta["company name"] = pos_meta["quote_name"]
    else:
        src_wb = openpyxl.load_workbook(BytesIO(input_bytes), data_only=True)
        src_ws = src_wb.active
        config_ws = _find_config_sheet(src_wb)

        if _is_qar_report(src_ws):
            quote_ref, quote_meta = _extract_qar_metadata(src_ws)
            items, config_rows = _extract_items_qar(src_ws)
        # Try grouped template
        elif _find_grouped_header(src_ws) is not None:
            quote_ref, date_text = _extract_grouped_metadata(src_ws)
            items, grp_config_rows = _extract_items_grouped(src_ws, item_quote_refs)
            config_rows = _extract_config_rows(config_ws) if config_ws else grp_config_rows
        else:
            # Metadata
            quote_ref, date_text = _extract_metadata(src_ws)
            expiry_text = _extract_expiry(src_ws)
            quote_meta = _extract_quote_metadata(src_ws)

            # Items: pricing summary → compact → generic
            items_ps = _extract_items_pricing_summary(src_ws)
            if items_ps:
                items = items_ps
                config_rows = _extract_config_rows(config_ws) if config_ws else _extract_all_config_rows(src_ws)
            else:
                compact_items, compact_config = _extract_items_compact(src_ws)
                if compact_items:
                    items = compact_items
                    config_rows = _extract_config_rows(config_ws) if config_ws else compact_config
                else:
                    items = _extract_items_generic(src_ws)
                    config_rows = _extract_config_rows(config_ws) if config_ws else _extract_all_config_rows(src_ws)

            consolidation_fee = _extract_consolidation_fee(src_ws) + _extract_shipping_fee(src_ws)
            part_numbers = _extract_part_numbers(config_ws or src_ws)
            heading_skus = _extract_product_detail_heading_skus(src_ws)

        if not quote_meta:
            quote_meta = _extract_quote_metadata(src_ws)

    quote_meta = {k: _strip_trailing_asterisk(v) for k, v in quote_meta.items()}

    return {
        "items": items,
        "config_rows": config_rows,
        "quote_ref": quote_ref,
        "date_text": date_text,
        "expiry_text": expiry_text,
        "quote_meta": quote_meta,
        "consolidation_fee": consolidation_fee,
        "part_numbers": part_numbers,
        "item_quote_refs": item_quote_refs,
        "heading_skus": heading_skus,
        "is_pdf": is_pdf,
        "is_docx": is_docx,
    }


def generate_southcomp_quote(
    input_bytes: bytes,
    margin_percent: float,
    currency_code: str,
    exchange_rate: float,
) -> bytes:
    """
    Generate a Southcomp Polaris EUR-style quote workbook.
    currency_code: 'EUR' or 'USD'
    exchange_rate: EUR/USD rate (used when currency_code='EUR')
    Returns raw xlsx bytes.
    """
    currency_code = (currency_code or "EUR").upper()
    effective_rate = exchange_rate if currency_code == "EUR" else 1.0

    data = extract_quote_source_data(input_bytes, exchange_rate)

    # Apply margin to consolidation fee (margin as share of selling price)
    margin_factor = margin_percent / 100.0
    consolidation_fee = data["consolidation_fee"]
    adjusted_consolidation_fee = consolidation_fee / (1 - margin_factor) if margin_factor < 1 else consolidation_fee

    return _build_quote_workbook(
        items=data["items"],
        config_rows=data["config_rows"],
        quote_ref=data["quote_ref"],
        date_text=data["date_text"],
        expiry_text=data["expiry_text"],
        quote_meta=data["quote_meta"],
        currency_code=currency_code,
        exchange_rate=effective_rate,
        consolidation_fee=adjusted_consolidation_fee,
        margin_percent=margin_percent,
        is_pdf=data["is_pdf"],
        part_numbers=data["part_numbers"] or None,
        include_config_sheet=not data["is_docx"],
    )



# ==================== ITEM CREATION (Item code + compact spec Description) ====================
# Builds the Southcomp item-creation import sheet from any supported Dell
# quote input, reusing extract_quote_source_data() so it always sees exactly
# the same items/config_rows the quotation-generator tool does. The
# Description is assembled from a fixed set of Product Details /
# Configuration modules (Base, Display, Processor, Memory, Storage, Wireless,
# Operating System, Primary Battery), each compacted down to its essentials —
# a module that isn't present for a given item (e.g. a server has no Display
# or Wireless) is simply skipped rather than left as an error.

_TRADEMARK_RE = re.compile(r"[®™]|\(R\)|\(TM\)", re.IGNORECASE)

# Intel Core "Ultra 7 265U" / "Ultra 5 236V" style: captures tier + model code.
# The optional "(?:processor\s+)?" skips the filler word some PDFs insert
# between the tier and the model, e.g. "Ultra 7 processor 265HX".
_PROC_ULTRA_RE = re.compile(r"ultra\D{0,12}?(\d+)\s+(?:processor\s+)?([A-Za-z0-9]+)", re.IGNORECASE)
# Intel Core "i3-14100" / "i7-1370P" / "i5 14th Gen 14500" style — the
# generation words are skipped so the model ("14500"), not "14th", is kept.
_PROC_COREI_RE = re.compile(
    r"\bi([3579])[\s\-]+(?:\d{1,2}(?:st|nd|rd|th)\s+gen(?:eration)?\s+)?([A-Za-z0-9]+)",
    re.IGNORECASE,
)
# Intel "Core 5 320" / "Core 7 150U" style (no Ultra, no i-prefix).
_PROC_CORE_N_RE = re.compile(r"\bcore\s+([3579])\s+(\d{3,4}[A-Za-z]{0,2})\b", re.IGNORECASE)
# Intel Xeon "Silver 4410Y" / "Gold 6548Y+" / "w5-2455X" style — tier word is
# dropped, model kept (including a "+" suffix, which is a different CPU).
_PROC_XEON_RE = re.compile(
    r"xeon\s*(?:gold|silver|platinum|bronze)?\s*([A-Za-z0-9]+(?:-[A-Za-z0-9]+)?\+?)",
    re.IGNORECASE,
)
# AMD "EPYC 9555P" style.
_PROC_EPYC_RE = re.compile(r"epyc\s+([A-Za-z0-9]+)", re.IGNORECASE)
# Tried in order; the first family that matches names the processor.
_PROC_FAMILIES: List[Tuple["re.Pattern", str]] = [
    (_PROC_ULTRA_RE, "Cu{}-{}"),
    (_PROC_COREI_RE, "i{}-{}"),
    (_PROC_CORE_N_RE, "C{}-{}"),
    (_PROC_XEON_RE, "Xeon-{}"),
    (_PROC_EPYC_RE, "EPYC-{}"),
]
# "12 cores" or the "12C/24T" shorthand server BOQs use. \b before the digit
# keeps this from matching digits embedded in a compound token like "A725".
_PROC_CORES_RE = re.compile(r"\b(\d+)\s*(?:cores?\b|C/\d+T)", re.IGNORECASE)
# "up to 5.0 GHz" / "3.20GHz" (space optional).
_PROC_GHZ_RE = re.compile(r"\b(\d+(?:\.\d+)?)\s*GHz", re.IGNORECASE)
# Dell's server-BOQ shorthand, e.g. "...Silver 4410Y 2G, 12C/24T,..." -> "2G" = 2.0GHz.
_PROC_GSHORT_RE = re.compile(r"\b(\d+(?:\.\d+)?)G\b(?=[,\s])", re.IGNORECASE)
# "... with 32GB Memory" tail some Ultra chips fold their onboard RAM into.
_PROC_MEM_TAIL_RE = re.compile(r"with\s+(\d+)\s*GB\s+Memory", re.IGNORECASE)

_DISPLAY_INCH_RE = re.compile(r'(\d+(?:\.\d+)?)\s*"')
_GB_RE = re.compile(r"\b(\d+(?:\.\d+)?)\s*GB", re.IGNORECASE)
_STORAGE_SIZE_RE = re.compile(r"\b(\d+(?:\.\d+)?)\s*(GB|TB)\b", re.IGNORECASE)
_BATTERY_YEAR_RE = re.compile(r"\b(\d+)[\s-]*(?:year|yr)", re.IGNORECASE)
_WIFI_GEN_RE = re.compile(r"wi-?fi\s*(\d+(?:/\d+)?E?)\b", re.IGNORECASE)
_WIFI_MODEL_RE = re.compile(
    r"\b((?:AX|BE|AC)\d{3}[A-Z]?|RTL\d{4}[A-Z]*|MT\d{4}[A-Z]*|QCN[A-Z]*\d{3,4}[A-Z]*)\b",
    re.IGNORECASE,
)
# A Dell SKU-shaped bracketed code, e.g. "(210-BPPK)" — NOT "(AZERTY)" (no
# digit/dash). Used only against untrusted text (an item's own top-level
# pricing-line name, which can contain unrelated brackets like a language
# or layout code).
_SKU_BRACKET_RE = re.compile(r"\((\d{3,4}-[A-Za-z0-9]{2,10})\)")
_SKU_SHAPE_RE = re.compile(r"^\d{3,4}-[A-Za-z0-9]{2,10}")
# A looser bracketed model code, e.g. "(PB14250)" — allowed only against text
# already confirmed to be the "Base" module's own description, where a
# bracket is reliably a Dell model code even without the digit-dash shape.
_MODEL_BRACKET_RE = re.compile(r"\(([A-Za-z0-9\-]{4,15})\)")
# A Dell model code inside an item's own name, e.g. "Dell Pro Dock - WD25",
# "Dell Pro Slim QCS1250" — the only identifier Dell's online-store quote
# gives items that carry no part number at all.
_MODEL_CODE_RE = re.compile(r"\b([A-Z]{2,3}\d{2,5}[A-Z0-9]*)\b")
# Dell's order code in square brackets, e.g. "... Customer Kit - [VPXRKM]".
_ORDER_CODE_RE = re.compile(r"\[([A-Z0-9]{5,8})\]")
# "[config name]" / "- [order code]" tails Dell appends to item names.
_SQUARE_BRACKETS_RE = re.compile(r"\s*-?\s*\[[^\[\]]*\]")
_BASE_BUILD_SUFFIX_RE = re.compile(r"(?:\s+(?:XCTO|CTO|BTO|BTX|Base))+\s*$", re.IGNORECASE)
_BASE_MODEL_TAIL_RE = re.compile(r",\s*[A-Z]{1,3}\d{4,6}\s*$")
# A Dell client model code left unbracketed in a product name, e.g. "Dell Pro
# 24 All-in-One QC24251 35W" or "Dell Pro Slim QCS1250". Only 2 letters + 5
# digits or 3 letters + 4 digits: a PowerEdge "XE9680" / "HS5610" (2 + 4) is
# the server's own model name and must stay.
_BASE_MODEL_CODE_RE = re.compile(r"\s*\b(?:[A-Z]{2}\d{5}|[A-Z]{3}\d{4})\b")


def _strip_square_brackets(text: str) -> str:
    # Innermost first, so a nested "[A - [B]]" goes away in two passes.
    prev = None
    while prev != text:
        prev, text = text, _SQUARE_BRACKETS_RE.sub("", text)
    return text.strip()


def _abbreviate_base(desc: str) -> str:
    """"Dell Pro 14 Plus (PB14250) XCTO Base" -> "Dell Pro 14 Plus";
    "Precision 3280 CFF CTO BASE" -> "Precision 3280 CFF"."""
    text = _strip_square_brackets(_TRADEMARK_RE.sub("", desc))
    text = text.split("(", 1)[0].strip() or text.strip()
    text = _BASE_MODEL_TAIL_RE.sub("", text)
    text = _BASE_MODEL_CODE_RE.sub("", text)
    text = _BASE_BUILD_SUFFIX_RE.sub("", text).strip()
    return text or desc.strip()


def _abbreviate_processor(desc: str) -> str:
    text = _TRADEMARK_RE.sub("", desc)
    parts = []
    for pattern, fmt in _PROC_FAMILIES:
        m = pattern.search(text)
        if m:
            parts.append(fmt.format(*m.groups()))
            break
    mc = _PROC_CORES_RE.search(text)
    if mc:
        parts.append(f"{mc.group(1)}C")
    # "2.6 GHz to 5.0 GHz" -> the top speed, matching the "up to N GHz"
    # figure the other Dell processor descriptions give.
    speeds = _PROC_GHZ_RE.findall(text) or _PROC_GSHORT_RE.findall(text)
    if speeds:
        # One style whatever Dell wrote: "5.40" -> "5.4", "2G" -> "2.0".
        ghz = f"{max(float(s) for s in speeds):.2f}".rstrip("0")
        parts.append(f"{ghz}0GHz" if ghz.endswith(".") else f"{ghz}GHz")

    if parts:
        return " ".join(parts)
    # Unrecognized processor family — fall back to the same "before the (" trim as Base.
    return text.split("(", 1)[0].strip()


def _abbreviate_display(desc: str) -> str:
    m = _DISPLAY_INCH_RE.search(desc)
    return f'{m.group(1)}"' if m else ""


def _abbreviate_storage(desc: str) -> str:
    """"512 GB, TLC, SSD" -> "512GB SSD". Empty when there's no capacity:
    that's an empty bay ("No Hard Drive") or a mechanical part listed under
    the same module ("Thermal Pad, Screw and Rubber for SSD"), not a drive."""
    size_m = _STORAGE_SIZE_RE.search(desc)
    if not size_m:
        return ""
    low = desc.lower()
    if "ssd" in low:
        kind = "SSD"
    elif "nvme" in low:
        kind = "NVMe"
    elif "hdd" in low or "hard drive" in low:
        kind = "HDD"
    elif "m.2" in low or "pcie" in low:
        kind = "SSD"
    else:
        kind = ""
    return f"{size_m.group(1)}{size_m.group(2).upper()} {kind}".strip()


def _abbreviate_wireless(desc: str) -> str:
    """Normalized to "Wi-Fi <generation> <card>", e.g. "Wi-Fi 6E AX211",
    whichever order Dell wrote them in ("Intel BE201 Wi-Fi 7 2x2, ...")."""
    text = _TRADEMARK_RE.sub("", desc)
    gen_m = _WIFI_GEN_RE.search(text)
    model_m = _WIFI_MODEL_RE.search(text)
    if gen_m or model_m:
        parts = []
        if gen_m:
            # "Wi-Fi 6/6E" -> "6E" (the card supports both; 6E is the higher one)
            parts.append("Wi-Fi " + gen_m.group(1).upper().split("/")[-1])
        if model_m:
            parts.append(model_m.group(1).upper())
        return " ".join(parts)
    text = text.split(",", 1)[0].strip()
    text = re.sub(r"^(Intel|Dell)\s+", "", text, flags=re.IGNORECASE)
    return text.strip()


def _abbreviate_os(desc: str) -> str:
    return re.sub(r"\s{2,}", " ", _TRADEMARK_RE.sub("", desc)).split(",", 1)[0].strip()


def _abbreviate_battery(desc: str) -> str:
    m = _BATTERY_YEAR_RE.search(desc)
    return f"{m.group(1)}Y" if m else ""


_CPU_FAMILY_RE = re.compile(
    r"\b(?:xeon|epyc|ryzen|threadripper|core\s+(?:ultra|i[3579]\b|[3579]\b))",
    re.IGNORECASE,
)
_CPU_EXCLUDE_RE = re.compile(r"heatsink|thermal|label|branding|graphics", re.IGNORECASE)
# Grouped/compact BOQs list components with no module column, so those rows
# are recognized by their own text instead.
_UNLABELED_MEMORY_RE = re.compile(r"^\s*\d+\s*GB\b.*?(?:DIMM|DDR\d|CAMM)", re.IGNORECASE)
_UNLABELED_STORAGE_RE = re.compile(
    r"^\s*(?:\(\d+\)\s*)*\d+(?:\.\d+)?\s*(?:GB|TB)\b.*?(?:SSD|HDD|NVMe|Hard Drive)",
    re.IGNORECASE,
)
_UNLABELED_OS_RE = re.compile(
    r"^\s*(?:no operating system|(?:microsoft\s+)?windows\s+(?:server|\d+)|vmware|red hat|suse|ubuntu)",
    re.IGNORECASE,
)
_STORAGE_MODULE_EXCLUDE_RE = re.compile(
    r"controller|config|raid|cable|bracket|reader|boot|software|adapter|driver",
    re.IGNORECASE,
)
# Mechanical parts listed under the same module as the real component, e.g.
# a "Wireless" row for the WLAN card's screw, or an external antenna.
_ITEM_JUNK_RE = re.compile(r"\b(?:screw|antenna|bracket|filler|thermal|rubber|cable)\b", re.IGNORECASE)


def _is_cpu_text(desc: str) -> bool:
    text = _TRADEMARK_RE.sub("", desc or "")
    return bool(_CPU_FAMILY_RE.search(text)) and not _CPU_EXCLUDE_RE.search(text)


def _classify_config_row(module: str, desc: str, sku: str) -> str:
    """Which description field a config row feeds ("" = none).

    module is normalized (lowercase, single-spaced). "system" is a base row
    that isn't labelled "Base": some exports name the base module after the
    product itself ("PowerEdge R6715", "Dell Pro Max Slim FCS1250") or have
    no module column at all — Dell system bases are always 210- SKUs.
    """
    if not module:
        if sku.startswith("210-"):
            return "system"
        if _is_cpu_text(desc):
            return "processor"
        if _UNLABELED_MEMORY_RE.match(desc):
            return "memory"
        if _UNLABELED_STORAGE_RE.match(desc):
            return "storage"
        if _UNLABELED_OS_RE.match(desc):
            return "os"
        return ""
    if module == "base":
        return "base"
    if module == "display":
        return "display"
    if module == "processor":
        return "processor"
    if module == "additional processor":
        # "No Additional Processor" vs. a real second CPU
        return "processor" if _is_cpu_text(desc) else ""
    if module in ("memory", "memory capacity"):
        return "memory"
    if ("storage" in module or "hard drive" in module) and not _STORAGE_MODULE_EXCLUDE_RE.search(module):
        return "storage"
    if module == "wireless":
        return "wireless"
    if module == "operating system":
        return "os"
    if module == "primary battery":
        return "battery"
    return "system" if sku.startswith("210-") else ""


def _row_qty(row: Tuple) -> int:
    m = re.match(r"\s*(\d+)", str(row[5] or ""))
    return max(int(m.group(1)), 1) if m else 1


def _build_item_description(item_no: str, config_rows: List[Tuple], item_name: str = "") -> str:
    """Compact spec line for one item, e.g. 'Dell Pro 16 Plus 16" Cu7-265U 12C
    5.3GHz 16GB 512GB SSD Wi-Fi 6E AX211 Windows 11 Pro'.

    Servers keep their counts: "2x Xeon-6548Y+ ...", total memory over all
    DIMMs, and each drive group ("2x 3.84TB SSD + 6x 480GB SSD"). Returns ""
    when the item has no usable configuration, so the caller falls back to
    the item's own name.
    """
    base = system = display = cpu = cpu_mem_tail = wireless = os_name = battery = ""
    cpu_count = 0
    mem_total = 0.0
    storage: List[str] = []
    for row in config_rows:
        if row[0] != item_no:
            continue
        desc = (row[3] or "").strip()
        if not desc:
            continue
        module = re.sub(r"\s+", " ", row[2] or "").strip().lower()
        kind = _classify_config_row(module, desc, (row[4] or "").strip())
        if kind == "base":
            base = base or _abbreviate_base(desc)
        elif kind == "system":
            system = system or _abbreviate_base(desc)
        elif kind == "display":
            display = display or _abbreviate_display(desc)
        elif kind == "processor":
            abbr = _abbreviate_processor(desc)
            if abbr:
                cpu = cpu or abbr
                cpu_count += _row_qty(row)
            m = _PROC_MEM_TAIL_RE.search(desc)
            if m and not cpu_mem_tail:
                cpu_mem_tail = f"{m.group(1)}GB"
        elif kind == "memory":
            m = _GB_RE.search(desc)
            if m:
                mem_total += float(m.group(1)) * _row_qty(row)
        elif kind == "storage":
            abbr = _abbreviate_storage(desc)
            if abbr:
                qty = _row_qty(row)
                storage.append(f"{qty}x {abbr}" if qty > 1 else abbr)
        elif kind == "wireless":
            if not wireless and not _ITEM_JUNK_RE.search(desc):
                wireless = _abbreviate_wireless(desc)
        elif kind == "os":
            os_name = os_name or _abbreviate_os(desc)
        elif kind == "battery":
            battery = battery or _abbreviate_battery(desc)

    if cpu and cpu_count > 1:
        cpu = f"{cpu_count}x {cpu}"
    if mem_total:
        memory = f"{int(mem_total)}GB" if mem_total == int(mem_total) else f"{mem_total}GB"
    else:
        # Onboard memory folded into the processor name ("... with 32GB
        # Memory") stands in only when there's no Memory row of its own —
        # otherwise the same 32GB would be listed twice.
        memory = cpu_mem_tail
    specs = [display, cpu, memory, " + ".join(storage), wireless, os_name, battery]

    head = base or system
    if not head:
        # No base row at all (e.g. a desktop in Dell's online-store quote).
        # A real system still has several spec fields; anything less is an
        # accessory whose own name describes it better than one field would.
        if sum(1 for s in specs if s) < 2:
            return ""
        head = _abbreviate_base(item_name) if item_name else ""
    return " ".join(p for p in [head] + specs if p)


def _resolve_item_code(item_no: str, config_rows: List[Tuple], fallback_desc: str, heading_sku: str = "") -> str:
    """Resolve the Dell part number to show as this item's code.

    Tried in order, since not every template exposes a SKU the same way:
      1. The "Base" module's own SKU column (most templates).
      2. Any 210- (Dell system base) SKU among the item's rows — exports
         that name the base module after the product instead of "Base".
      3. A bracketed model code inside the "Base" module's own description,
         e.g. "(PB14250)" — still a trusted source (it IS the Base row),
         just not shaped like a full order SKU.
      4. A SKU-shaped bracketed code in the item's heading text — some
         Configuration-sheet Excel exports name the base module after the
         product itself (e.g. "PowerEdge R6715") instead of "Base", but still
         carry "(210-BPPK)"-style brackets in the per-item heading.
      5. heading_sku: the SKU from the item's Product Details heading even
         when no component table follows it (a plain accessory), or the
         Word order form's Part Number column.
      6. The first config row for this item that carries any real SKU at
         all (covers BOQs where even the first/base row uses the product
         name as its module label, with no separate "Base" row).
      7. A SKU-shaped bracketed code in the item's own top-level pricing
         description — deliberately strict (digit-dash shaped only) since
         this text is untrusted: it can carry unrelated brackets like
         "(AZERTY)", a keyboard layout, not a part number.
      8. A Dell model code in the item's own name ("Dell Pro Dock - WD25"),
         then a bracketed Dell order code ("[VPXRKM]") — last resorts for
         quotes that carry no part number at all for this item.
    """
    base_sku = ""
    base_desc = ""
    system_sku = ""
    heading_code = ""
    first_sku = ""
    for row in config_rows:
        if row[0] != item_no:
            continue
        sku = (row[4] or "").strip()
        if (row[2] or "").strip().lower() == "base" and not base_sku:
            base_desc = row[3] or ""
            base_sku = sku
        if sku.startswith("210-") and not system_sku:
            system_sku = sku
        if not first_sku and _SKU_SHAPE_RE.match(sku):
            first_sku = sku
        if not heading_code:
            heading = row[1] or ""
            if heading:
                m = _SKU_BRACKET_RE.search(heading)
                if m:
                    heading_code = m.group(1)

    if base_sku:
        return base_sku
    if system_sku:
        return system_sku
    m = _MODEL_BRACKET_RE.search(base_desc)
    if m:
        return m.group(1)
    if heading_code:
        return heading_code
    if heading_sku:
        return heading_sku
    if first_sku:
        return first_sku
    m = _SKU_BRACKET_RE.search(fallback_desc or "")
    if m:
        return m.group(1)
    models = _MODEL_CODE_RE.findall(_strip_square_brackets(fallback_desc or ""))
    if models:
        return models[-1]
    m = _ORDER_CODE_RE.search(fallback_desc or "")
    return m.group(1) if m else ""


def _normalize_quote_ref(ref: str) -> str:
    """First quote number in ref, with a BOQ-style "v4" version written the
    way Dell's PDFs (and the Southcomp sample) write it: "3400021144897.4"."""
    ref = (ref or "").split(",")[0].strip()
    return re.sub(r"^(\d{6,})v(\d+)$", r"\1.\2", ref, flags=re.IGNORECASE)


_SOUTHCOMP_IMPORT_HEADERS = [
    "Item", "Description", "ItemType", "LottedYN", "ShwRoom", "ProductLine", "SalesCat", "AccountCode", "Currency", "TaxClass", "Unit", "DeprecType", "StdProdLine", "StdProdCateg", "Userfield1", "Userfield2", "Userfield3", "UserFld4", "UserField 5", "COO", "HSCode", "VendorId", "ECCN", "ItemStatus", "StdProdLineType", "UPC", "ItemGroup", "SpecialLCId", "EcotaxeID", "SorecopID", "StCondId", "ProvCountry", "CTOYN", "ArabDescr", "ArabAddlDescr", "HighValueYN", "QtyOrderMin", "STKUseAsSerYN", "Weight", "ModelNo", "SplitSectorYN", "UserField6", "UserField7", "UserField8", "WHTVatId", "WHTIncId", "WhseItemYN", "RegNum", "RemoveDiscountFOCYN", "Userfield9", "Userfield10", "Userfield11", "Userfield12", "Userfield13", "Userfield14", "AcceptFOCYN", "Integration1", "Integration2", "Integration3", "MOHCode", "GTIN", "ExtWarr", "CTOItemId", "DemoYN", "AddlDescr", "StdItemId", "RptLoc",
]


def extract_item_creation_rows(input_bytes: bytes) -> List[Tuple[str, str]]:
    """Return [(item_code, description), ...] for one uploaded Dell quote.

    The item code is prefixed with the quote reference to match Southcomp item
    creation imports, e.g. "3400023336849.1-210-BPCK".
    """
    data = extract_quote_source_data(input_bytes)
    items = data["items"]
    config_rows = data["config_rows"]
    item_quote_refs = data.get("item_quote_refs") or {}
    heading_skus = data.get("heading_skus") or {}
    # The Word order form's "Part Number" column. Excel inputs fill
    # part_numbers from a looser, non-Dell pattern, so only SKU-shaped
    # values are trusted.
    part_numbers = {
        k: v.strip() for k, v in (data.get("part_numbers") or {}).items()
        if _SKU_SHAPE_RE.match((v or "").strip())
    }
    rows: List[Tuple[str, str]] = []
    for idx, item in enumerate(items, start=1):
        item_no = str(idx)
        raw_name = (item[0] if item else "").strip()
        item_name = _strip_square_brackets(_TRADEMARK_RE.sub("", raw_name))
        description = _build_item_description(item_no, config_rows, item_name) or item_name
        description = re.sub(r"\s+,", ",", re.sub(r"\s+", " ", description)).strip()
        item_code = _resolve_item_code(
            item_no, config_rows, raw_name, heading_skus.get(item_no) or part_numbers.get(item_no, "")
        )
        quote_ref = _normalize_quote_ref(item_quote_refs.get(item_no) or data.get("quote_ref") or "")
        if quote_ref and item_code and not item_code.startswith(quote_ref):
            item_code = f"{quote_ref}-{item_code}"
        if item_code or description:
            rows.append((item_code, description))
    return rows


def _build_item_import_row(item_code: str, description: str) -> list:
    row = [None] * len(_SOUTHCOMP_IMPORT_HEADERS)
    row[0] = item_code
    row[1] = description
    row[2] = 1
    row[3] = 0
    row[4] = "Entrepot"
    row[5] = "DELL"
    row[6] = "DL"
    row[7] = "SE"
    row[8] = "EUR"
    row[9] = 0
    row[10] = "PCS"
    row[12] = "BR_0042"
    row[14] = "DELL"
    row[15] = "DELL"
    row[21] = 4012142
    row[23] = 1
    row[66] = "CODE REPORTING"
    return row


def generate_item_creation_excel(rows: List[Tuple[str, str]]) -> bytes:
    """Build the Southcomp item-creation workbook: the sample template's
    header row, then one import row per item."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"  # same sheet name as the Southcomp sample template
    ws.sheet_view.showGridLines = False

    header_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
    header_font = Font(bold=True, color="000000")
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    ws.append(_SOUTHCOMP_IMPORT_HEADERS)
    for col, header in enumerate(_SOUTHCOMP_IMPORT_HEADERS, start=1):
        cell = ws.cell(1, col)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="left", vertical="center")

    for item_code, description in rows:
        ws.append(_build_item_import_row(item_code, description))

    for col_idx, _ in enumerate(_SOUTHCOMP_IMPORT_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 80
    ws.freeze_panes = "A2"

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row_idx, col_idx)
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def generate_item_creation_excel_from_inputs(
    inputs: List[Tuple[str, bytes]],
) -> Tuple[bytes, int, List[str]]:
    """inputs: [(source_name, file_bytes), ...].

    Returns (xlsx_bytes, row_count, notes). Rows from all uploaded quotes are
    combined into one workbook; an (item, description) pair already seen
    (e.g. the same laptop config quoted twice) is written once. notes holds
    one line per file the user should look at: unreadable, no Dell items
    found, or items left without a code.
    """
    all_rows: List[Tuple[str, str]] = []
    seen = set()
    variants: Dict[str, int] = {}
    notes: List[str] = []
    for name, data in inputs:
        try:
            file_rows = extract_item_creation_rows(data)
        except Exception as e:
            notes.append(f"{name}: could not be read ({e}).")
            continue
        if not file_rows:
            notes.append(f"{name}: no Dell items found — is it a Dell quote?")
            continue
        missing_codes = sum(1 for code, _ in file_rows if not code)
        if missing_codes:
            notes.append(f"{name}: {missing_codes} item(s) have no item code — fill them in before importing.")
        for code, desc in file_rows:
            if (code, desc) in seen:
                continue
            seen.add((code, desc))
            if code:
                # Same base SKU in one quote but a different configuration
                # (e.g. two PowerEdge R660 nodes with different drives): the
                # item code must stay unique, so number the later variants.
                variants[code] = variants.get(code, 0) + 1
                if variants[code] > 1:
                    code = f"{code}-{variants[code]}"
            all_rows.append((code, desc))
    return generate_item_creation_excel(all_rows), len(all_rows), notes


def build_item_creation_filename() -> str:
    return f"Southcomp_Polaris_Item_Creation_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"


def build_output_filename(currency_code: str = "EUR", source_name: str = "") -> str:
    stem = ""
    if source_name:
        raw_stem = os.path.splitext(os.path.basename(source_name))[0]
        stem = re.sub(r"[^\w\-]+", "_", raw_stem).strip("_")
    parts = ["Southcomp_Polaris"]
    if stem:
        parts.append(stem)
    parts.append(currency_code.upper())
    parts.append(datetime.now().strftime("%Y%m%d_%H%M"))
    return "_".join(parts) + ".xlsx"
