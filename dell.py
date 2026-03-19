# dell.py
from datetime import datetime
from io import BytesIO
from typing import Optional, Dict, List, Tuple
import base64
import logging
import os
import re
from logging.handlers import RotatingFileHandler

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string as colidx
from openpyxl.drawing.image import Image as XLImage


# ----------------- Logging -----------------
_LOG_FILE = os.path.join(os.path.dirname(__file__), "dell_quote.log")


def _get_logger():
    """Get a logger that writes to a rotating log file."""
    logger = logging.getLogger("dell_quote")
    if not logger.handlers:
        handler = RotatingFileHandler(_LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
    return logger


def _log_items(prefix: str, items: list, max_items: int = 20):
    """Log a summary of extracted items (truncates to avoid huge logs)."""
    logger = _get_logger()
    count = len(items) if items is not None else 0
    logger.debug("%s: %d items", prefix, count)
    if not items:
        return
    for i, item in enumerate(items[:max_items]):
        logger.debug("%s item %d: %s", prefix, i + 1, item)
    if count > max_items:
        logger.debug("%s: ... (skipping remaining %d items)", prefix, count - max_items)

try:
    from PIL import Image as PILImage
except Exception:
    PILImage = None


# ================= Helpers =================

def _parse_money(val):
    """Parse strings like '$ 902.00', '902,00', '36,080.00' to float."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r"[^\d,.\-]", "", s)  # remove currency & spaces
    if "," in s and "." in s:
        # treat comma as thousands
        s = s.replace(",", "")
    else:
        # if only comma present, treat as decimal
        if "," in s and "." not in s:
            s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _cell_to_text(v, fallback=""):
    if v is None:
        return fallback
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def _normalize_text(s: str) -> str:
    """Lowercase alnum-only text used for fuzzy matching."""
    return re.sub(r"[^a-z0-9]", "", s.lower()) if s else ""


def _make_item_key(s: str, key_len: int = 70) -> str:
    return _normalize_text(s)[:key_len]




def _row_text(ws, r, c1=1, c2=None) -> str:
    if c2 is None:
        c2 = ws.max_column
    parts = []
    for c in range(c1, c2 + 1):
        txt = _cell_to_text(ws.cell(r, c).value)
        if txt:
            parts.append(txt)
    return " ".join(parts).strip()


def _is_price_or_qty_line(text: str) -> bool:
    t = text.lower()
    if any(tok in t for tok in [
        "qty", "quantity", "unit price", "unitprice", "subtotal", "total",
        "price", "amount", "discount", "tax", "extended price", "net price",
        "grand total", "msrp", "usd", "aed", "eur", "sar"
    ]):
        return True
    if re.search(r"(\$|€|£|د\.إ|aed|usd|eur|sar)", t, flags=re.IGNORECASE):
        return True
    return False


def _add_logo(ws, logo_bytes: Optional[bytes], anchor="A1"):
    """
    Add logo from uploaded bytes or fallback to local 'image.png'.
    Anchor is inside the merged area A1:B4; merging prevents row-1 stretching issues.
    """
    if logo_bytes and PILImage is not None:
        try:
            pil_img = PILImage.open(BytesIO(logo_bytes))
            img = XLImage(pil_img)
            img.width = 180
            img.height = 60
            ws.add_image(img, anchor)
            return
        except Exception:
            pass
    try:
        img = XLImage("image.png")
        img.width = 180
        img.height = 60
        ws.add_image(img, anchor)
    except Exception:
        pass


def _extract_metadata_strict(ws):
    """Extract quote ref/date from strict positions in the worksheet."""
    logger = _get_logger()
    raw_ref = ws["E15"].value
    quote_ref = "" if raw_ref is None else (
        raw_ref.strftime("%d/%m/%Y") if isinstance(raw_ref, datetime) else str(raw_ref).strip()
    )

    raw_date = ws["E18"].value
    if isinstance(raw_date, datetime):
        quote_date = raw_date.strftime("%d/%m/%Y")
    else:
        quote_date = "" if raw_date is None else str(raw_date).strip()

    logger.debug("_extract_metadata_strict: quote_ref=%s, quote_date=%s", quote_ref, quote_date)
    return quote_ref, quote_date


def _find_header_row_strict_or_detect(ws):
    """
    Try to detect header in the source (within first 40 rows) by locating a row containing
    both 'Description' and 'Qty/Quantity'. If found, return the first data row (header+1)
    and the column indices for description/qty/unit price.

    Fallback (for INPUT): assume header is on row 7 and data starts on row 8 with
    columns C/D/E for Description/Qty/Unit Price.

    Returns:
      (first_data_row_index, desc_col, qty_col, unit_col)
    """
    for r in range(1, min(ws.max_row, 40) + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(row_vals):
            continue
        row_text = [(_cell_to_text(v).lower()) for v in row_vals]
        if any("description" in t for t in row_text) and any(("qty" in t) or ("quantity" in t) for t in row_text):
            desc_idx = qty_idx = unit_idx = None
            for i, v in enumerate(row_vals, start=1):
                name = _cell_to_text(v).lower()
                if desc_idx is None and "description" in name:
                    desc_idx = i
                if qty_idx is None and (("qty" in name) or ("quantity" in name)):
                    qty_idx = i
                if unit_idx is None and (("unit price" in name) or ("unitprice" in name) or (name == "price")):
                    unit_idx = i
            if desc_idx is None: desc_idx = 3
            if qty_idx  is None: qty_idx  = 4
            if unit_idx is None: unit_idx = 5
            return r + 1, desc_idx, qty_idx, unit_idx
    return 8, 3, 4, 5


# ---- Pricing Summary (Sheet 1 items) ----

def _locate_pricing_summary(ws):
    B = colidx('B')
    for r in range(30, min(ws.max_row, 120) + 1):
        v = ws.cell(r, B).value
        if v and "pricing" in str(v).lower() and "summary" in str(v).lower():
            header_row = r + 1
            start_row = header_row + 2
            return header_row, start_row
    return None


def _try_extract_items_from_pricing_summary(ws):
    logger = _get_logger()
    located = _locate_pricing_summary(ws)
    if not located:
        logger.debug("_try_extract_items_from_pricing_summary: pricing summary not found")
        return None

    header_row, start_row = located
    logger.debug("_try_extract_items_from_pricing_summary: header_row=%d start_row=%d", header_row, start_row)
    A, B, K, L, N = (colidx('A'), colidx('B'), colidx('K'), colidx('L'), colidx('N'))

    items = []
    r = start_row
    while r <= ws.max_row:
        sr = ws.cell(r, A).value
        if sr is None or str(sr).strip() == "":
            break
        if not re.match(r'^\d+', str(sr).strip()):
            break

        desc_text = _cell_to_text(ws.cell(r, B).value)
        if not desc_text:
            break

        qty_raw  = ws.cell(r, K).value
        unit_raw = ws.cell(r, L).value
        sub_raw  = ws.cell(r, N).value

        try:
            qty_val = int(_parse_money(qty_raw) or 0)
        except Exception:
            qty_val = 0
        unit_val = _parse_money(unit_raw) or 0.0
        subtotal_val = _parse_money(sub_raw)
        if subtotal_val is None:
            subtotal_val = qty_val * unit_val

        if (qty_val <= 0) and (unit_val == 0.0) and (subtotal_val is None or subtotal_val == 0.0):
            break

        items.append((desc_text, qty_val, unit_val, subtotal_val))
        r += 1

    if items:
        logger.debug("_try_extract_items_from_pricing_summary: extracted %d items", len(items))
        _log_items("Pricing summary", items)
    else:
        logger.debug("_try_extract_items_from_pricing_summary: no items extracted")

    return items if items else None


# ---- Product Details -> Configuration table (Sheet 2) ----

def _find_product_details_anchor(ws) -> Optional[int]:
    max_c = min(ws.max_column, 40)
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if v and "product details" in str(v).lower():
                return r
    return None


def _find_config_table_header(ws, start_row: int, search_rows: int = 30) -> Optional[Tuple[int, Dict[str, int]]]:
    """Find the row that contains the header like 'Module | Description | SKU | Tax Type | Qty'.
    Returns (header_row, columns_map). columns_map keys: module, description, sku, tax, qty (optional if missing).
    """
    last_row = min(ws.max_row, start_row + search_rows)
    for r in range(start_row, last_row + 1):
        labels = {}
        for c in range(1, ws.max_column + 1):
            name = _cell_to_text(ws.cell(r, c).value).lower()
            if not name:
                continue
            if 'module' in name and 'module' not in labels:
                labels['module'] = c
            if 'description' in name and 'description' not in labels:
                labels['description'] = c
            if name.strip() in ('sku', 'part', 'part #', 'part#') and 'sku' not in labels:
                labels['sku'] = c
            if 'tax' in name and 'type' in name and 'tax' not in labels:
                labels['tax'] = c
            if name.strip() in ('qty', 'quantity') and 'qty' not in labels:
                labels['qty'] = c
        # Require at least Module, Description, SKU
        if all(k in labels for k in ('module', 'description', 'sku')):
            return r, labels
    return None


def _collect_config_rows_for_product(ws, start_row: int, columns: Dict[str, int], next_product_start: Optional[int]) -> List[Tuple[str, str, str, str]]:
    """Collect configuration lines for a product starting at start_row (first line AFTER header).
    Stops before next_product_start if provided, otherwise when a strong stop condition occurs.
    Returns list of tuples (module, description, sku, tax_type). Qty/Price are intentionally ignored.
    """
    rows = []
    r = start_row
    while r <= ws.max_row and (next_product_start is None or r < next_product_start):
        # a blank row ends the table chunk
        txt_whole = _row_text(ws, r, 1, min(ws.max_column, 40))
        if not txt_whole:
            break
        # sometimes date or separator lines appear – skip those
        if _is_price_or_qty_line(txt_whole) or re.search(r"estimated delivery|ship|subtotal|total", txt_whole, re.I):
            r += 1
            continue
        m = _cell_to_text(ws.cell(r, columns.get('module', 0)).value)
        d = _cell_to_text(ws.cell(r, columns.get('description', 0)).value)
        s = _cell_to_text(ws.cell(r, columns.get('sku', 0)).value)
        t = _cell_to_text(ws.cell(r, columns.get('tax', 0)).value)
        # If the row is essentially empty, stop
        if not any([m, d, s, t]):
            break
        rows.append((m, d, s, t))
        r += 1
    return rows


def _extract_quote_metadata(ws):
    """Extract quote metadata (Company Name, Customer Name, etc.) from the input sheet.

    Dell quote layout puts labels in column B and values in column E, e.g.:
        B22: "Company Name:"   E22: "ACME"
    """
    keys = {
        "company name": "Company Name",
        "customer name": "Customer Name",
        "customer number": "Customer Number",
        "end user": "End User",
        "reseller": "Reseller",
    }

    out = {k: "" for k in keys}
    max_row = min(ws.max_row, 120)
    for r in range(1, max_row + 1):
        label = _cell_to_text(ws.cell(r, 2).value).strip().lower().rstrip(":")
        if label in keys:
            out[label] = _cell_to_text(ws.cell(r, 5).value)
    return out


def _extract_excel_consolidation_fee(ws) -> float:
    """Find 'Consolidation Fee:' in Excel and read the first non-empty cell to its right."""
    logger = _get_logger()

    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if not isinstance(value, str):
                continue
            if "consolidation fee" not in value.strip().lower():
                continue

            for next_col in range(cell.column + 1, ws.max_column + 1):
                next_value = ws.cell(cell.row, next_col).value
                if next_value in (None, ""):
                    continue

                parsed = _parse_money(next_value)
                consolidation_fee = parsed or 0.0
                logger.debug(
                    "Excel consolidation fee found at row=%s col=%s, value_col=%s, raw_value=%s, parsed=%s",
                    cell.row,
                    cell.column,
                    next_col,
                    next_value,
                    consolidation_fee,
                )
                return consolidation_fee

            logger.debug(
                "Excel consolidation fee label found at row=%s col=%s, but no value exists to the right",
                cell.row,
                cell.column,
            )
            return 0.0

    logger.debug("Excel consolidation fee not found; defaulting to 0.0")
    return 0.0


def _extract_pdf_lines(pdf_bytes: bytes) -> List[str]:
    """Extract a cleaned line list from PDF using pdfplumber (best-effort).

    Falls back to pypdf extraction if pdfplumber isn't available or fails.
    """
    lines: List[str] = []
    try:
        import pdfplumber

        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                words = page.extract_words(use_text_flow=True)
                if not words:
                    continue
                # group words by y0 (line position)
                rows = {}
                for w in words:
                    y = round(w.get("top", 0))
                    rows.setdefault(y, []).append(w)
                for y in sorted(rows.keys()):
                    row_words = sorted(rows[y], key=lambda w: w.get("x0", 0))
                    line = " ".join(w.get("text", "") for w in row_words)
                    lines.append(line.strip())
        if lines:
            return lines
    except Exception:
        pass

    # Fallback to pypdf extraction
    try:
        from pypdf import PdfReader
    except ImportError:
        raise RuntimeError("pypdf is required to parse PDF quotes")

    reader = PdfReader(BytesIO(pdf_bytes))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    return [l.strip() for l in text.splitlines()]


def _extract_pdf_quote_data(pdf_bytes: bytes):
    """
    FINAL VERSION — CLEAN, HYBRID (C3), PDF‑ONLY FIX
    Extracts:
      • FULL item list (13+ items across all pages)
      • FULL Product Details (C3 Hybrid)
      • Clean multi‑line merging
      • Correct item mapping
    """

    logger = _get_logger()
    lines = _extract_pdf_lines(pdf_bytes)
    logger.debug("_extract_pdf_quote_data: extracted %d lines", len(lines))

    # ---------------------- METADATA ----------------------
 # ---------------------- METADATA ----------------------
    metadata = {
        "company name": "",
        "customer name": "",
        "customer number": "",
        "end user": "",
        "reseller": "",
    }
    quote_ref_text = ""
    date_text = ""
    consolidation_fee = 0.0

    pending_keys = []
    prev_label = None
    pdf_label_aliases = {
        "authorized partner": "reseller",
    }

    MONTHS = r"(January|February|March|April|May|June|July|August|September|October|November|December)"
    DATE_RX = re.compile(rf"{MONTHS}\s+\d{{1,2}},\s+\d{{4}}")

    def extract_dates(s: str):
        return [m.group(0) for m in DATE_RX.finditer(s)]

    def _normalize_pdf_label(label: str) -> str:
        normalized = label.strip().lower().rstrip(":")
        mapped = pdf_label_aliases.get(normalized, normalized)
        if mapped != normalized:
            logger.debug("PDF metadata label alias: '%s' -> '%s'", normalized, mapped)
        return mapped

    def _extract_pdf_reseller(lines: List[str]) -> str:
        stop_markers = (
            "billing information",
            "shipping information",
            "quote summary",
            "product details",
        )

        for idx, line in enumerate(lines):
            lower_line = line.lower()
            if "authorized partner" not in lower_line:
                continue

            collected: List[str] = []

            same_line_match = re.search(r"authorized partner\s*:\s*(.+)$", line, re.IGNORECASE)
            if same_line_match and same_line_match.group(1).strip():
                collected.append(same_line_match.group(1).strip())

            next_idx = idx + 1
            while next_idx < len(lines):
                candidate = lines[next_idx].strip()
                candidate_lower = candidate.lower()
                if not candidate:
                    break
                if any(marker in candidate_lower for marker in stop_markers):
                    break
                collected.append(candidate)
                next_idx += 1

            if not collected:
                return ""

            reseller_text = " ".join(collected)
            reseller_text = re.sub(r"\s+", " ", reseller_text).strip()

            if "page name" in lower_line:
                uppercase_start = re.search(r"\b[A-Z]{2,}(?:\s+[A-Z]{2,})+", reseller_text)
                if uppercase_start:
                    reseller_text = reseller_text[uppercase_start.start():].strip()

            reseller_text = re.sub(r"\s*-\s*Authorized Partner\+?$", "", reseller_text, flags=re.IGNORECASE).strip()
            reseller_text = reseller_text.rstrip("+").strip()
            return reseller_text

        return ""

    for line in lines:
        stripped = line.strip()
        lower = stripped.lower()

        # ---------------- MULTILINE KEY HANDLING ----------------
        if pending_keys and stripped:
            line_val = stripped

            # Case: Quote number + Quote date (+ expiration) on the SAME LINE
            if {"quote number", "quote date"}.issubset(set(pending_keys)):
                m = re.match(r"^\s*(\d+)\b(.*)$", line_val)
                if m:
                    quote_ref_text = m.group(1)
                    tail = m.group(2).strip()
                else:
                    tail = line_val

                dates = extract_dates(tail)
                if len(dates) >= 1:
                    date_text = dates[0]

                pending_keys = []
                continue

            # Generic fallback: last key takes the rest
            parts = stripped.split()
            for idx, key in enumerate(pending_keys):
                if idx == len(pending_keys) - 1:
                    val = " ".join(parts[idx:])
                else:
                    val = parts[idx] if idx < len(parts) else ""

                if key == "quote number":
                    m = re.search(r"\b\d{6,}\b", val)
                    if m:
                            quote_ref_text = m.group(0)
                            logger.debug("PDF metadata multiline: quote number=%s", quote_ref_text)
                elif key == "quote date":
                    if DATE_RX.search(line_val):
                        val = DATE_RX.search(line_val).group(0)
                    date_text = val
                    logger.debug("PDF metadata multiline: quote date=%s", date_text)
                elif key in metadata:
                    metadata[key] = val
                    logger.debug("PDF metadata multiline: %s=%s", key, val)

            pending_keys = []
            continue

        # ---------------- LABEL WAITING FOR NEXT LINE ----------------
        if prev_label:
            if prev_label == "quote number":
                # extract ONLY the numeric Quote Number
                m = re.search(r"\b\d{6,}\b", stripped)
                if m:
                    quote_ref_text = m.group(0)
                    logger.debug("PDF metadata next-line: quote number=%s", quote_ref_text)
        
            elif prev_label == "quote date":
                if DATE_RX.search(stripped):
                    date_text = DATE_RX.search(stripped).group(0)
                else:
                    date_text = stripped
                logger.debug("PDF metadata next-line: quote date=%s", date_text)
        
            elif prev_label in metadata:
                metadata[prev_label] = stripped
                logger.debug("PDF metadata next-line: %s=%s", prev_label, stripped)
        
            prev_label = None
            continue

        # ---------------- KEY:VALUE ON SAME LINE ----------------
        if ":" in stripped:
            parts = [p.strip() for p in re.split(r"\s*:\s*", stripped) if p.strip()]
            normalized = [_normalize_pdf_label(p) for p in parts]

            if stripped.endswith(":"):
                for key in normalized:
                    if key in ("quote number", "quote date", "company name", "customer name", "customer number", "reseller"):
                        pending_keys.append(key)
                        logger.debug("PDF metadata pending label: %s", key)
                continue

            for i in range(0, len(parts), 2):
                key = normalized[i]
                val = parts[i+1] if i+1 < len(parts) else ""

                if key == "quote number":
                    quote_ref_text = val
                    logger.debug("PDF metadata same-line: quote number=%s", quote_ref_text)
                elif key == "quote date":
                    m = DATE_RX.search(val)
                    date_text = m.group(0) if m else val
                    logger.debug("PDF metadata same-line: quote date=%s", date_text)
                elif key in metadata:
                    metadata[key] = val
                    logger.debug("PDF metadata same-line: %s=%s", key, val)

            continue

        # ---------------- LABEL-ONLY LINE ----------------
        normalized_lower = _normalize_pdf_label(lower)
        if normalized_lower in ("quote number", "quote date", "company name", "customer name", "customer number", "reseller"):
            prev_label = normalized_lower
            logger.debug("PDF metadata label-only line detected: %s", normalized_lower)
            continue

    # FINAL FALLBACK: scan for date if still empty
    if not date_text:
        for l in lines:
            m = DATE_RX.search(l)
            if m:
                date_text = m.group(0)
                logger.debug("PDF metadata fallback date=%s", date_text)
                break

    reseller_from_layout = _extract_pdf_reseller(lines)
    if reseller_from_layout:
        metadata["reseller"] = reseller_from_layout
        logger.debug("PDF reseller layout extraction=%s", reseller_from_layout)

    # Extract consolidation fee when present; otherwise keep the zero fallback.
    for i, line in enumerate(lines):
        low = line.lower().strip()
        if "consolidation fee" not in low:
            continue

        same_line_match = re.search(r"consolidation fee[:\s]*([$€£]?\s*[\d,]+(?:\.\d+)?)", line, re.IGNORECASE)
        if same_line_match:
            consolidation_fee = _parse_money(same_line_match.group(1)) or 0.0
            logger.debug("PDF consolidation fee same-line=%s", consolidation_fee)
            break

        if i + 1 < len(lines):
            next_line_value = _parse_money(lines[i + 1])
            if next_line_value is not None:
                consolidation_fee = next_line_value
                logger.debug("PDF consolidation fee next-line=%s", consolidation_fee)
                break

    logger.debug(
        "PDF metadata summary: quote_ref=%s, date=%s, reseller=%s, company=%s, customer=%s, customer_number=%s, consolidation_fee=%s",
        quote_ref_text,
        date_text,
        metadata.get("reseller", ""),
        metadata.get("company name", ""),
        metadata.get("customer name", ""),
        metadata.get("customer number", ""),
        consolidation_fee,
    )

    # ---------------------- ITEMS (FULL 13+ extraction) ----------------------
    def _try_parse_item(line: str):
        # Example:
        #   "DELL USB-C Mobile Adapter - DA310 65.1 3 195.3"
        m = re.match(
            r"^(?P<desc>.*?)(?P<unit>\d[\d,\.]*?)\s+(?P<qty>\d+)\s+(?P<total>\d[\d,\.]*)$",
            line
        )
        if not m:
            return None
        desc = m.group("desc").strip()
        qty = int(m.group("qty"))
        unit = _parse_money(m.group("unit")) or 0.0
        total = _parse_money(m.group("total")) or 0.0
        return desc, qty, unit, total

    items = []
    in_items = False

    for line in lines:
        low = line.lower().strip()

        # Enter items region
        if "quote summary" in low:
            in_items = True
            continue
        if "unit price" in low and "qty" in low and "item total" in low:
            in_items = True
            continue

        # Stop items at Product Details
        if "product details" in low:
            break

        if not in_items:
            continue

        # Skip empty lines
        if not line.strip():
            continue

        parsed = _try_parse_item(line)
        if parsed:
            items.append(parsed)

    logger.debug("EXTRACTED PDF ITEMS = %d", len(items))

    # ---------------- CONFIGURATION (C3 HYBRID) -----------------
    config_rows = []

    try:
        pd_idx = next(i for i, l in enumerate(lines) if "product details" in l.lower())
    except StopIteration:
        pd_idx = None

    if pd_idx is not None:
        current_item = ""
        current_heading = ""
        awaiting_heading = False
        item_counter = 0

        # Helper: merge value line into previous module
        def attach_value_to_previous(text_line: str) -> bool:
            if not config_rows:
                return False
            (it, hd, mod, dsc, sku, tax) = config_rows[-1]
            if it == current_item and (dsc == "" or dsc is None):
                if ":" not in text_line and len(text_line.split()) >= 2:
                    config_rows[-1] = (it, hd, mod, text_line.strip(), sku, tax)
                    return True
            return False

        # Clean heading
        def _clean_heading_text(t: str):
            return re.sub(r"\s+\d[\d,\.]*\s+\d[\d,\.]*\s+\d[\d,\.]*$", "", t).strip()

        i = pd_idx + 1
        while i < len(lines):
            line = lines[i].strip()
            low = line.lower()

            if "unit price" in low and "qty" in low and "item total" in low:
                awaiting_heading = True
                i += 1
                continue

            if awaiting_heading:
                if line.lower() == "description":
                    i += 1
                    continue

                # Next line = heading
                parsed = _try_parse_item(line)
                if parsed:
                    desc, qty, unit, total = parsed
                    current_heading = _clean_heading_text(desc)
                else:
                    current_heading = _clean_heading_text(line)

                item_counter += 1
                current_item = str(item_counter)
                awaiting_heading = False
                i += 1
                continue

            # Stop when we captured all headings (items found)
            if item_counter >= len(items) and any(
                k in low for k in ["ship to", "subtotal", "total", "important notes"]
            ):
                break

            if not current_item:
                i += 1
                continue

            # Skip junk
            if low.startswith("page ") or low.startswith("category description"):
                i += 1
                continue

            # key:value pairs
            if ":" in line:
                parts = [p.strip() for p in re.split(r"\s*:\s*", line) if p.strip()]
                for j in range(0, len(parts), 2):
                    mod = parts[j]
                    dsc = parts[j+1] if j+1 < len(parts) else ""
                    if mod.lower() != "category description":
                        config_rows.append((current_item, current_heading, mod, dsc, "", ""))
                i += 1
                continue

            # Attach as value if previous row missing value
            if attach_value_to_previous(line):
                i += 1
                continue

            # Otherwise, treat as a module label
            config_rows.append((current_item, current_heading, line, "", "", ""))
            i += 1

    # ---------------- MERGE FRAGMENTED ROWS -----------------
    def collapse(rows):
        out = []
        i = 0
        while i < len(rows):
            item, head, mod, dsc, sku, tax = rows[i]
            mod = mod.strip()
            dsc = dsc.strip()

            # Merge rows like:
            #   "Select Power Cord Type"
            #   "UK/Irish Power Cord"
            if i + 1 < len(rows):
                ni, nh, nmod, ndsc, _, _ = rows[i+1]
                if ni == item and nh == head:
                    if dsc == "" and ":" not in nmod:
                        # second line is value
                        dsc = nmod
                        i += 1

            out.append((item, head, mod, dsc, sku, tax))
            i += 1
        return out

    config_rows = collapse(config_rows)
    logger.debug("PDF CONFIG ROWS = %d", len(config_rows))

    return items, metadata, config_rows, quote_ref_text, date_text, consolidation_fee

def _extract_all_config_rows(ws) -> List[Tuple[str, str, str, str, str, str]]:
    """
    FINAL CLEAN VERSION (Excel upload only)
    Extracts ALL configuration rows under Product Details from Excel files.

    Output rows have the structure:
      (item_number, item_heading, module, description, sku, tax_type)

    • Preserves original Excel order
    • Skips totals / shipping / price lines
    • Merges 2-line fragments
    • Cleans up module/value pairs
    • Avoids duplicated headings
    """

    anchor = _find_product_details_anchor(ws)
    if not anchor:
        return []

    rows: List[Tuple[str, str, str, str, str, str]] = []
    r = anchor + 1
    max_col = min(ws.max_column, 40)

    current_item = ""
    current_heading = ""
    item_counter = 0

    def _clean_heading_text(text: str) -> str:
        # remove trailing price fragments like "4 $200.00 $800.00"
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

    # Detect module header positions
    def _find_header(row_idx: int) -> Optional[Tuple[int, dict]]:
        return _find_config_table_header(ws, row_idx, search_rows=20)

    # ----------------------
    # Main scanning loop
    # ----------------------
    while r <= ws.max_row:
        txt = _row_text(ws, r, 1, max_col)

        # Detect new item heading (e.g. "1. Dell Latitude 5520")
        m = re.match(r"^\s*(\d+)\.", txt)
        if m:
            current_item = m.group(1)
            current_heading = _clean_heading_text(txt)
            r += 1
            continue

        header_info = _find_header(r)
        if not header_info:
            r += 1
            continue

        header_row, colmap = header_info
        data_row = header_row + 1

        # Skip empty rows after header
        while data_row <= ws.max_row and not _row_text(ws, data_row, 1, max_col):
            data_row += 1

        # If no item heading was detected, synthesize one
        if not current_item:
            item_counter += 1
            current_item = str(item_counter)
            current_heading = f"Item {current_item}"

        # Scan rows until table ends
        while data_row <= ws.max_row:
            row_text_all = _row_text(ws, data_row, 1, max_col)

            if not row_text_all:
                break
            if _is_table_stop(row_text_all):
                data_row += 1
                continue

            # Detect accidental new headings inside table
            m2 = re.match(r"^\s*(\d+)\.", row_text_all)
            if m2:
                current_item = m2.group(1)
                current_heading = _clean_heading_text(row_text_all)
                break

            # Extract cells
            mod = _cell_to_text(ws.cell(data_row, colmap.get("module", 0)).value)
            desc = _cell_to_text(ws.cell(data_row, colmap.get("description", 0)).value)
            sku = _cell_to_text(ws.cell(data_row, colmap.get("sku", 0)).value)
            tax = _cell_to_text(ws.cell(data_row, colmap.get("tax", 0)).value)

            if not any([mod, desc, sku, tax]):
                break

            rows.append((current_item, current_heading, mod, desc, sku, tax))
            data_row += 1

        r = data_row + 1

    # -----------------------
    # MERGE FRAGMENTED ROWS
    # -----------------------
    cleaned = []
    i = 0
    while i < len(rows):
        item, head, mod, desc, sku, tax = rows[i]
        mod, desc = mod.strip(), desc.strip()

        # CASE: 2-line module label (common in Dell exports)
        if i + 1 < len(rows):
            ni, nh, nmod, ndesc, nsku, ntax = rows[i + 1]
            if ni == item and nh == head:
                if desc == "" and ndesc == "" and ":" not in mod and ":" not in nmod:
                    # join "Smart" + "Dock SD25TB5"
                    mod = f"{mod} {nmod}".strip()
                    i += 1

        cleaned.append((item, head, mod, desc, sku, tax))
        i += 1

    return cleaned
# ================= Main =================

def generate_dell_quote(
    input_excel_bytes: bytes,
    logo_bytes: Optional[bytes] = None,
) -> bytes:
    """
    Generate a 2-sheet workbook from either:
      - Dell quote Excel file (as bytes) OR
      - Dell quote PDF (as bytes)

      Output:
      - 'Quote' formatted with strict template:
          - A1:B4 merged for logo
          - Address block in D1:F3 (merged per row)
          - 'Quote Ref' shown at C5 with value at D5 (read from INPUT E15 or PDF "Quote number")
          - 'Date' shown at C6 with value at D6 (read from INPUT E18 or PDF "Quote date")
          - Table header at row 8; data from row 9
      - 'Configuration' sheet that replicates the 'Product Details' configuration table(s)
        PER PRODUCT, preserving columns Module, Description, SKU, Tax Type (dropping Qty/Unit Price/Subtotal/date lines).
    """
    logger = _get_logger()
    logger.info("Generating Dell quote (bytes=%d)", len(input_excel_bytes) if input_excel_bytes is not None else 0)
    
    # --- Log full uploaded file in safe base64 format ---
    try:
        logger.debug(
            "Uploaded file FULL CONTENT (base64, length=%d): %s",
            len(input_excel_bytes),
            base64.b64encode(input_excel_bytes).decode()
        )
    except Exception as e:
        logger.error("Failed to base64-log uploaded file: %s", e)

    # Missing consolidation fee should be treated as zero for both Excel and PDF uploads.
    consolidation_fee = 0.0

    # ---- Load source ----
    is_pdf = input_excel_bytes.lstrip().startswith(b"%PDF")

    if is_pdf:
        items, quote_meta, config_rows, quote_ref_text, date_text, consolidation_fee = _extract_pdf_quote_data(input_excel_bytes)
        logger.info("Parsed PDF quote: %d items, quote_ref=%s, date=%s", len(items), quote_ref_text, date_text)
        _log_items("PDF items", items)
    else:
        src_wb = openpyxl.load_workbook(BytesIO(input_excel_bytes), data_only=True)
        src_ws = src_wb.active
        logger.info("Parsed Excel quote (sheets=%d, active=%s)", len(src_wb.sheetnames), src_ws.title)

        # ---- Extract metadata (STRICT E15/E18) ----
        quote_ref_text, date_text = _extract_metadata_strict(src_ws)
        logger.info("Extracted metadata from Excel: quote_ref=%s, date=%s", quote_ref_text, date_text)
        quote_meta = _extract_quote_metadata(src_ws)

        # ---- Extract items (Pricing Summary layout first; else generic) ----
        items_ps = _try_extract_items_from_pricing_summary(src_ws)
        if items_ps:
            items = items_ps
            logger.info("Found %d items via Pricing Summary extraction", len(items))
            _log_items("Pricing summary items", items)
        else:
            first_data_row, desc_col, qty_col, unit_col = _find_header_row_strict_or_detect(src_ws)
            logger.info("Using generic item extraction starting at row %d (desc_col=%d, qty_col=%d, unit_col=%d)", first_data_row, desc_col, qty_col, unit_col)
            items = []
            r = first_data_row
            while r <= src_ws.max_row:
                desc = src_ws.cell(r, desc_col).value
                qty  = src_ws.cell(r, qty_col).value
                unit = src_ws.cell(r, unit_col).value

                desc_text = _cell_to_text(desc)
                if not desc_text or desc_text.lower().startswith("total"):
                    break
                try:
                    qty_val = int(qty) if qty not in (None, "") else 0
                except Exception:
                    qty_val = int(_parse_money(qty) or 0)
                unit_val = _parse_money(unit) or 0.0
                if qty_val > 0:
                    items.append((desc_text, qty_val, unit_val, None))
                r += 1
            logger.info("Extracted %d items via generic table parsing", len(items))
            _log_items("Parsed table items", items)

        item_descs_order = [it[0] for it in items]

        # ---- Extract metadata (STRICT E15/E18) ----
        quote_ref_text, date_text = _extract_metadata_strict(src_ws)
        quote_meta = _extract_quote_metadata(src_ws)

        # ---- Extract items (Pricing Summary layout first; else generic) ----
        items_ps = _try_extract_items_from_pricing_summary(src_ws)
        if items_ps:
            items = items_ps
        else:
            first_data_row, desc_col, qty_col, unit_col = _find_header_row_strict_or_detect(src_ws)
            items = []
            r = first_data_row
            while r <= src_ws.max_row:
                desc = src_ws.cell(r, desc_col).value
                qty  = src_ws.cell(r, qty_col).value
                unit = src_ws.cell(r, unit_col).value

                desc_text = _cell_to_text(desc)
                if not desc_text or desc_text.lower().startswith("total"):
                    break
                try:
                    qty_val = int(qty) if qty not in (None, "") else 0
                except Exception:
                    qty_val = int(_parse_money(qty) or 0)
                unit_val = _parse_money(unit) or 0.0
                if qty_val > 0:
                    items.append((desc_text, qty_val, unit_val, None))
                r += 1

        item_descs_order = [it[0] for it in items]
        
        consolidation_fee = _extract_excel_consolidation_fee(src_ws)
        
            
    


    # ---- Build output workbook ----
    wb = Workbook()
    ws = wb.active
    ws.title = "Quote"
    ws.sheet_view.showGridLines = False
    
    
    # ---- Step 2: Write Consolidation Fee and Factor ----
    # Cell H2 → Consolidation Fee
    ws["H2"] = consolidation_fee
    ws["H2"].font = Font(bold=True, color="1F497D")
    ws["H2"].alignment = Alignment(horizontal="center", vertical="center")
    
    # Cell H3 stays empty.
    ws["H3"].value = ""
    ws["H3"].font = Font(bold=True, color="1F497D")
    ws["H3"].alignment = Alignment(horizontal="center", vertical="center")
    
    
    # Column widths (A..H; merge A+B for logo)
    widths = {"A": 12, "B": 42, "C": 12, "D": 16, "E": 18, "G": 18, "H": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Header rows height (1..5)
    for rr in range(1, 6):
        ws.row_dimensions[rr].height = 20

    # ===== HEADER: merge A1:B4 and add logo =====
    ws.merge_cells("A1:B4")
    _add_logo(ws, logo_bytes, anchor="A1")

    # ===== Address block in D..F (merge per row) =====
    ws.merge_cells("D1:F1")
    ws.merge_cells("D2:F2")
    ws.merge_cells("D3:F3")
    ws["D1"] = "P O Box 55609, Dubai, UAE"
    ws["D2"] = "Tel :  +9714 4500600    Fax : +9714 4500678"
    ws["D3"] = "Website :  www.mindware.ae"
    for cell in ("D1", "D2", "D3"):
        ws[cell].font = Font(bold=True, size=11, color="1F497D")
        ws[cell].alignment = Alignment(horizontal="left", vertical="center")

    # ===== META =====
    ws["C7"] = "Quote Ref"
    ws["C7"].font = Font(bold=True, color="1F497D")
    ws["D7"] = quote_ref_text

    ws["C8"] = "Date"
    ws["C8"].font = Font(bold=True, color="1F497D")
    ws["D8"] = date_text

    # ---- Quote metadata (Company Name / Customer Name / Customer Number / End User / Reseller)
    meta_rows = [
        ("Company Name:", quote_meta.get("company name", "")),
        ("Customer Name:", quote_meta.get("customer name", "")),
        ("Customer Number:", quote_meta.get("customer number", "")),
        ("End User:", quote_meta.get("end user", "")),
        ("Reseller:", quote_meta.get("reseller", "")),
    ]

    for idx, (label, value) in enumerate(meta_rows, start=4):
        ws[f"G{idx}"] = label
        ws[f"G{idx}"].font = Font(bold=True)
        ws[f"H{idx}"] = value
        ws[f"H{idx}"].alignment = Alignment(wrap_text=True, vertical="center")

    # ===== TABLE HEADER at row 8; data from row 9 =====
    header_row = 9
    ws["A9"] = "Sr. No."
    ws["B9"] = "Description"
    ws["C9"] = "Qty"
    ws["D9"] = "Unit Price"
    ws["E9"] = "Total Price"
    ws["F9"] = "Adj/Unit"
    ws["G9"] = "Final Unit Price"

    header_fill = PatternFill(start_color="9BBB59", end_color="9BBB59", fill_type="solid")
    header_font = Font(bold=True, color="000000")
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for addr in ("A9", "B9", "C9", "D9", "E9", "F9", "G9"):
        ws[addr].fill = header_fill
        ws[addr].font = header_font
        ws[addr].alignment = Alignment(horizontal="center", vertical="center")
        ws[addr].border = border_thin
    ws.row_dimensions[header_row].height = 20

    # ===== DATA ROWS (start at 9) =====
    row_ptr = header_row + 1
    sr_no = 1
    currency_fmt = '"$"#,##0.00'
    yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_cells = []

    for (desc_text, qty_val, unit_val, subtotal_val) in items:
        ws[f"A{row_ptr}"] = sr_no
        ws[f"B{row_ptr}"] = desc_text
        ws[f"C{row_ptr}"] = qty_val
    
        # Original Unit Price
        ws[f"D{row_ptr}"].value = unit_val
        ws[f"D{row_ptr}"].number_format = currency_fmt
    
        # ---- Adj/Unit (F) uses the consolidation fee stored in H2 directly
        ws[f"F{row_ptr}"].value = f"=IF(C{row_ptr}>0,$H$2/C{row_ptr},0)"
        ws[f"F{row_ptr}"].number_format = currency_fmt
    
        # ---- Final Unit Price (G) = Unit Price + Adj/Unit
        ws[f"G{row_ptr}"].value = f"=D{row_ptr}+F{row_ptr}"
        ws[f"G{row_ptr}"].number_format = currency_fmt
    
        # ---- Total Price (E) = Qty * Final Unit Price
        ws[f"E{row_ptr}"].value = f"=C{row_ptr}*G{row_ptr}"
        ws[f"E{row_ptr}"].number_format = currency_fmt
    
        # Styling
        for addr in (f"A{row_ptr}", f"B{row_ptr}", f"C{row_ptr}", f"D{row_ptr}", f"E{row_ptr}", f"F{row_ptr}", f"G{row_ptr}"):
            ws[addr].fill = yellow
            ws[addr].border = border_thin
            ws[addr].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"B{row_ptr}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
        total_cells.append(f"E{row_ptr}")
        sr_no += 1
        row_ptr += 1

    # ===== TOTAL ROW =====
    ws.merge_cells(start_row=row_ptr, start_column=2, end_row=row_ptr, end_column=4)
    ws[f"B{row_ptr}"] = "Total price"
    ws[f"B{row_ptr}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"B{row_ptr}"].font = Font(bold=True, color="1F497D")

    ws[f"E{row_ptr}"] = f"=SUM({','.join(total_cells)})" if total_cells else 0
    ws[f"E{row_ptr}"].number_format = currency_fmt
    ws[f"E{row_ptr}"].font = Font(bold=True, color="1F497D")
    ws[f"E{row_ptr}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"E{row_ptr}"].border = border_thin

    # Footer notes
    notes = [
        'Incoterms:',
        '',
        'Payment Terms:',
        '',
        'Quote validity:',
        '',
        'Estimated Delivery Time from the date of booking:',
        '',
        'These prices do not include installation of any kind',
        'Change in Qty or partial shipment is not acceptable',
        'For all B2B orders complete end customer details should be mentioned on the PO',
        'PO Should be addressed to Mindware FZ LLC and should be in USD',
        'Orders once placed with Dell cannot be cancelled',
        '',
        'And as an important note – All items are not proposed with any Professional Services to cater for installation.',
        '',
        'Please note that these prices are granted for this QTY and deal & Dell cannot guarantee same prices if QTY is reduced , it will also have to be one shot order.',
        '',
        'Kindly also ensure to review the proposal specifications from your end and ensure that they match the requirements exactly as per the End User.',
    ]
    footer_row = max(row_ptr + 2, 22)
    for line in notes:
        ws.merge_cells(start_row=footer_row, start_column=2, end_row=footer_row, end_column=6)
        ws.cell(footer_row, 2).value = line
        ws.cell(footer_row, 2).alignment = Alignment(wrap_text=True, vertical="top")
        footer_row += 1

    ws.freeze_panes = "A9"

    # ===== Sheet 2: Configuration =====
    if not is_pdf:
        config_rows = _extract_all_config_rows(src_ws)

    ws2 = wb.create_sheet("Configuration")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 22  # Item #
    ws2.column_dimensions["B"].width = 70  # Module
    ws2.column_dimensions["C"].width = 100  # Description
    ws2.column_dimensions["D"].width = 20  # SKU
    ws2.column_dimensions["E"].width = 14  # Tax Type

    r2 = 1
    title_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    def write_table_header(row_index: int):
        ws2[f"A{row_index}"] = "Item #"
        ws2[f"B{row_index}"] = "Module"
        ws2[f"C{row_index}"] = "Description"
        ws2[f"D{row_index}"] = "SKU"
        ws2[f"E{row_index}"] = "Tax Type"
        for addr in (f"A{row_index}", f"B{row_index}", f"C{row_index}", f"D{row_index}", f"E{row_index}"):
            ws2[addr].font = Font(bold=True)
            ws2[addr].fill = title_fill
            ws2[addr].alignment = Alignment(horizontal="center", vertical="center")
            ws2[addr].border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )
        ws2.row_dimensions[row_index].height = 20

    # Header row
    write_table_header(r2)
    r2 += 1

    if not config_rows:
        ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=5)
        ws2[f"A{r2}"] = "(No configuration details found)"
        ws2[f"A{r2}"].alignment = Alignment(horizontal="left", vertical="center")
        r2 += 1
    else:
        current_item = None
        current_heading = None
        for (item, heading, module, dsc, sku, tax) in config_rows:
            # Insert a product heading section whenever the item number changes
            if item and item != current_item:
                # Item row (e.g. "Item 1")
                ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=5)
                ws2[f"A{r2}"] = f"Item {item}"
                ws2[f"A{r2}"].font = Font(bold=True, color="1F497D")
                ws2[f"A{r2}"].alignment = Alignment(horizontal="left", vertical="center")
                r2 += 1

                # Heading row (e.g. "1. PowerEdge ...")
                if heading:
                    ws2.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=5)
                    ws2[f"A{r2}"] = heading
                    ws2[f"A{r2}"].font = Font(italic=True, color="1F497D")
                    ws2[f"A{r2}"].alignment = Alignment(horizontal="left", vertical="center")
                    r2 += 1

                current_item = item
                current_heading = heading

            ws2[f"A{r2}"] = ""
            ws2[f"B{r2}"] = module
            ws2[f"C{r2}"] = dsc
            ws2[f"D{r2}"] = sku
            ws2[f"E{r2}"] = tax
            for col in ("A", "B", "C", "D", "E"):
                ws2[f"{col}{r2}"].alignment = Alignment(vertical="top", wrap_text=True)
                ws2[f"{col}{r2}"].border = Border(
                    left=Side(style="thin", color="DDDDDD"),
                    right=Side(style="thin", color="DDDDDD"),
                    top=Side(style="thin", color="DDDDDD"),
                    bottom=Side(style="thin", color="DDDDDD"),
                )
            r2 += 1
            

    ws2.freeze_panes = "A2"
    

    # Save to bytes
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
