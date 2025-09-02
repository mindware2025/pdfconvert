import re
from typing import List, Optional, Dict, Any, Tuple

import pdfplumber
import pandas as pd

from utils.helpers import normalize_line


DELL_INVOICE_COLS = [
    "Item",
    "Description",
    "Quantity",
    "Unit Price",
    "Amount",
]


def extract_invoice_info(pdf_path) -> tuple[Optional[str], Optional[str]]:
    """Extract Invoice Number and Invoice Date from a Dell invoice PDF.

    Tries a few common label variants.
    """
    invoice_number: Optional[str] = None
    invoice_date: Optional[str] = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:2]:  # Typically on first page
            text = page.extract_text() or ""
            for raw in text.splitlines():
                line = normalize_line(raw)
                if invoice_number is None and ("invoice number" in line or "invoice no" in line):
                    m = re.search(r"invoice\s+(?:number|no)\s*[:#-]?\s*([A-Za-z0-9-]+)", line)
                    if m:
                        invoice_number = m.group(1)
                if invoice_date is None and ("invoice date" in line or "date:" in line):
                    m = re.search(r"invoice\s+date\s*[:#-]?\s*([0-9]{1,2}[\-/ ][A-Za-z0-9]{3,}[\-/ ][0-9]{2,4})", line)
                    if not m:
                        m = re.search(r"\bdate\b\s*[:#-]?\s*([0-9]{1,2}[\-/ ][A-Za-z0-9]{3,}[\-/ ][0-9]{2,4})", line)
                    if m:
                        invoice_date = m.group(1)
            if invoice_number and invoice_date:
                break
    return invoice_number, invoice_date


def _normalize_headers(headers: List[str]) -> List[str]:
    return [normalize_line(h).lower() for h in headers]


def _find_dell_items_table(table: List[List[Optional[str]]]) -> Optional[dict]:
    """Given a raw table (list of rows), detect a header row with item columns.

    Returns a mapping of column indices if detected, else None.
    """
    for ridx, row in enumerate(table):
        headers = [str(c or "").strip() for c in row]
        norm = _normalize_headers(headers)
        has_desc = any("description" in c for c in norm)
        has_qty = any(("qty" in c) or ("quantity" in c) for c in norm)
        has_unit_price = any("unit price" in c or ("price" == c) for c in norm)
        has_amount = any("amount" in c or "total" in c for c in norm)
        if has_desc and has_qty and has_unit_price and has_amount:
            # Build column index mapping
            def idx_of(pred) -> int:
                for i, c in enumerate(norm):
                    if pred(c):
                        return i
                return -1

            idx_item = 0  # Usually first column is item/SKU
            idx_desc = idx_of(lambda c: "description" in c)
            idx_qty = idx_of(lambda c: ("qty" in c) or ("quantity" in c))
            idx_unit = idx_of(lambda c: "unit price" in c or c == "price")
            idx_amt = idx_of(lambda c: "amount" in c or "total" in c)
            return {
                "header_row": ridx,
                "idx_item": idx_item,
                "idx_desc": idx_desc,
                "idx_qty": idx_qty,
                "idx_unit": idx_unit,
                "idx_amt": idx_amt,
            }
    return None


def extract_table_from_text(pdf_path) -> List[List[str]]:
    """Extract Dell invoice items as rows aligned to DELL_INVOICE_COLS.

    Uses pdfplumber's table extraction and a heuristic header detector.
    """
    rows: List[List[str]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            used_fallback = False
            try:
                raw_tables = page.extract_tables() or []
            except Exception:
                raw_tables = []
            for table in raw_tables:
                mapping = _find_dell_items_table(table)
                if not mapping:
                    continue
                start = mapping["header_row"] + 1
                for r in table[start:]:
                    if not any((c is not None and str(c).strip() != "") for c in r):
                        continue
                    def safe(idx: int) -> str:
                        if idx < 0 or idx >= len(r):
                            return ""
                        v = r[idx]
                        return str(v).strip() if v is not None else ""

                    item = safe(mapping["idx_item"]) if mapping["idx_item"] >= 0 else ""
                    desc = safe(mapping["idx_desc"]) if mapping["idx_desc"] >= 0 else ""
                    qty = safe(mapping["idx_qty"]) if mapping["idx_qty"] >= 0 else ""
                    unit = safe(mapping["idx_unit"]) if mapping["idx_unit"] >= 0 else ""
                    amt = safe(mapping["idx_amt"]) if mapping["idx_amt"] >= 0 else ""

                    # Skip subtotal/total rows
                    line_norm = normalize_line(" ".join([desc, qty, unit, amt]))
                    if any(k in line_norm for k in ["subtotal", "total", "vat", "tax"]):
                        continue

                    rows.append([item, desc, qty, unit, amt])

            # Fallback: parse from plain text between header and VAT Summary
            if not raw_tables or not rows:
                text = page.extract_text() or ""
                if not text:
                    continue
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                in_items = False
                for line in lines:
                    raw_line = line
                    ln_low = normalize_line(line).lower()
                    if not in_items and ("item no" in ln_low and "description" in ln_low and "quantity" in ln_low and "unit price" in ln_low):
                        in_items = True
                        continue
                    if in_items and (ln_low.startswith("vat summary") or ln_low.startswith("vat type")):
                        break
                    if in_items:
                        # Example row:
                        # 210-BMFF Dell Pro 24 Plus Monitor - P2425H 16 118.28 1,892.48 NL
                        m = re.match(r"^([A-Z0-9-]+)\s+(.+?)\s+(\d{1,6})\s+([0-9,]+(?:\.[0-9]{2})?)\s+([0-9,]+(?:\.[0-9]{2})?)\s+[A-Z]{2}$", raw_line)
                        if m:
                            item, desc, qty, unit, amt = m.groups()
                            rows.append([item, desc, qty, unit, amt])
    return rows


def extract_header_fields(pdf_path) -> Dict[str, Any]:
    """Extract top-level Dell invoice metadata used for pre-alert output.

    Returns keys: po_number, invoice_number, invoice_date, customer_no,
    dell_order_no, shipping_method, ed_order, consolidation_fee_usd.
    """
    out: Dict[str, Any] = {
        "po_number": "",
        "invoice_number": "",
        "invoice_date": "",
        "customer_no": "",
        "dell_order_no": "",
        "shipping_method": "",
        "ed_order": "",
        "consolidation_fee_usd": "",
    }

    def capture_after(label: str, line: str) -> Optional[str]:
        m = re.search(rf"{label}\s*[:#-]?\s*(.+)$", line)
        return m.group(1).strip() if m else None

    with pdfplumber.open(pdf_path) as pdf:
        # Concatenate first two pages' text for robust header extraction
        full_text_parts: List[str] = []
        raw_text_parts: List[str] = []
        for page in pdf.pages[:2]:
            t = page.extract_text()
            if t:
                # Keep original case for values; normalize spaces
                lines = [normalize_line(x) for x in t.splitlines()]
                full_text_parts.append("\n".join(lines))
                raw_text_parts.append("\n".join([x.strip() for x in t.splitlines()]))
        full_text = "\n".join(full_text_parts)
        raw_full_text = "\n".join(raw_text_parts)

        def get(pattern: str) -> Optional[str]:
            m = re.search(pattern, full_text, flags=re.IGNORECASE)
            return m.group(1).strip() if m else None

        out["po_number"] = get(r"your\s*ref\s*/\s*po\s*no\s*:\s*(?:PO)?\s*([A-Za-z0-9\-_/]+)") or out["po_number"]
        out["invoice_number"] = get(r"invoice\s*no\s*:\s*([A-Za-z0-9\-]+)") or out["invoice_number"]
        out["invoice_date"] = get(r"invoice\s*date\s*:\s*([0-9]{1,2}[\-/][0-9]{1,2}[\-/][0-9]{2,4}|[0-9]{1,2}\s+[A-Za-z]{3,}\s+[0-9]{4})") or out["invoice_date"]
        out["customer_no"] = get(r"customer\s*no\s*:\s*([A-Za-z0-9\-]+)") or out["customer_no"]
        out["dell_order_no"] = get(r"dell\s*order\s*no\s*:\s*([A-Za-z0-9\-]+)") or out["dell_order_no"]
        out["shipping_method"] = get(r"shipping\s*method\s*:?[\s\n]*([A-Za-z0-9 \-–/]+)") or out["shipping_method"]
        if not out["shipping_method"]:
            # Fallback: capture block from 'Solution Name' down to before 'Funded By'
            raw_lines = raw_full_text.splitlines()
            start_idx = next((i for i, l in enumerate(raw_lines) if re.search(r"solution\s*name\s*:", l, re.IGNORECASE)), None)
            end_idx = next((i for i, l in enumerate(raw_lines) if i > (start_idx or -1) and re.search(r"^\s*funded\s+by\b", l, re.IGNORECASE)), None)
            if start_idx is not None:
                block = raw_lines[start_idx:(end_idx if end_idx is not None else start_idx + 6)]
                # Remove the 'Solution Name:' label on the first line
                if block:
                    block[0] = re.sub(r"(?i)solution\s*name\s*:\s*", "", block[0]).strip()
                # Join non-empty lines as AWB text
                joined = " ".join([b.strip() for b in block if b.strip()])
                out["shipping_method"] = joined
        out["ed_order"] = (
            get(r"select\s+account\s+to\s+charge\s*:?[\s\n]*([A-Za-z0-9\-]+)")
            or get(r"\bed\s*order\b\s*:?[\s\n]*([A-Za-z0-9\-]+)")
            or out["ed_order"]
        )

        # Consolidation (currency-agnostic): pick last numeric on consolidation line
        raw_lines = raw_full_text.splitlines()
        for i, line in enumerate(raw_lines):
            ln = line.lower()
            if "consolidation" in ln:
                # Only take numbers AFTER the word if present on the same line
                try:
                    post = line[ln.index("consolidation") + len("consolidation"):]
                except Exception:
                    post = ""
                def nums_in(s: str) -> List[str]:
                    return re.findall(r"([0-9][0-9,]*\.[0-9]{2}|[0-9][0-9,]*)", s)
                nums_post = nums_in(post)
                candidate = None
                def is_non_zero(s: str) -> bool:
                    t = s.replace(",", "").replace(" ", "")
                    t = t.replace("\u00A0", "")
                    t = t.replace("0,00", "0.00")
                    try:
                        return abs(float(t)) > 0.000001
                    except Exception:
                        return False
                if nums_post:
                    # prefer the last number after the keyword
                    nz = [n for n in nums_post if is_non_zero(n)]
                    candidate = (nz[-1] if nz else nums_post[-1])
                else:
                    # Look ahead up to 2 lines for wrapped values; choose last non-zero number
                    lookahead = []
                    if i + 1 < len(raw_lines):
                        lookahead.append(raw_lines[i + 1])
                    if i + 2 < len(raw_lines):
                        lookahead.append(raw_lines[i + 2])
                    all_nums: List[str] = []
                    for la in lookahead:
                        all_nums += nums_in(la)
                    if all_nums:
                        nz = [n for n in all_nums if is_non_zero(n)]
                        candidate = (nz[-1] if nz else all_nums[-1])
                if candidate is not None:
                    out["consolidation_fee_usd"] = candidate
                break

        return out


PRE_ALERT_HEADERS = [
    "PO Txn Code",
    "PO Number",
    "Supplier Invoice No",
    "Supplier Ref Date",
    "Dell ED",
    "AWB",
    "Bill of leading date",
    "Shipping Agent",
    "From Port",
    "To Port",
    "ETS",
    "ETA",
    "Item Code",
    "Item Desc",
    "UOM",
    "Qty",
    "Unit Rate",
    "Item code as per Dell pdf",
    "Item desc as per Dell pdf",
    "Consolidation fees",
]


def _normalize_po(po: str) -> str:
    s = str(po or "").strip()
    s = re.sub(r"(?i)^po\s*", "", s)
    return s

def _normalize_item_code(raw: str) -> str:
    """Normalize supplier/item codes for matching.

    - Trim, uppercase
    - If the string contains a leading label like 'Item code:' keep the first code token
    - Extract the first token that looks like A-Z/0-9 with dashes/underscores
    """
    s = str(raw or "").strip().upper()
    # Common label removal
    s = re.sub(r"(?i)^item\s*code\s*[:\-]*\s*", "", s)
    # Take first code-like token
    m = re.search(r"([A-Z0-9][A-Z0-9\-_]*[A-Z0-9])", s)
    return m.group(1) if m else s


def read_master_mapping(path_or_stream) -> Tuple[
    Dict[Tuple[str, str], Tuple[str, str]],
    Dict[Tuple[str, str], int],
    Dict[Tuple[str, str], int],
]:
    """Read the master Excel (header at row 9) and build a lookup.

    Key: (Po Num normalized without 'PO', Supplier Item Code)
    Value: (Orion Item Code, Pi Item Desc)
    """
    df = pd.read_excel(path_or_stream, header=8, dtype=str)
    def col(name: str) -> str:
        for c in df.columns:
            if str(c).strip().lower() == name.lower():
                return c
        raise KeyError(f"Missing column '{name}' in master file")

    c_po = col("Po Num")
    c_supplier = col("Supplier Item Code")
    c_orion = col("Orion Item Code")
    c_pi_desc = col("Pi Item Desc")

    lookup: Dict[Tuple[str, str], Tuple[str, str]] = {}
    supplier_counts: Dict[Tuple[str, str], int] = {}
    orion_counts: Dict[Tuple[str, str], int] = {}
    for _, r in df.iterrows():
        po = _normalize_po(r.get(c_po, ""))
        supp = _normalize_item_code(r.get(c_supplier, ""))
        orion = str(r.get(c_orion, "") or "").strip()
        pi_desc = str(r.get(c_pi_desc, "") or "").strip()
        if po and supp:
            key = (po, supp)
            lookup[key] = (orion, pi_desc)
            supplier_counts[key] = supplier_counts.get(key, 0) + 1
        if po and orion:
            okey = (po, _normalize_item_code(orion))
            orion_counts[okey] = orion_counts.get(okey, 0) + 1
    return lookup, supplier_counts, orion_counts


def build_pre_alert_rows(
    pdf_path,
    tomorrow_date: str,
    master_lookup: Optional[Dict[Tuple[str, str], Tuple[str, str]]] = None,
    supplier_counts: Optional[Dict[Tuple[str, str], int]] = None,
    orion_counts: Optional[Dict[Tuple[str, str], int]] = None,
    diagnostics: Optional[List[Dict[str, Any]]] = None,
) -> List[List[Any]]:
    """Build rows for the PRE ALERT UPLOAD sheet from a single PDF."""
    headers = extract_header_fields(pdf_path)
    items = extract_table_from_text(pdf_path)
    rows: List[List[Any]] = []
    for item in items:
        item_no = item[0] if len(item) > 0 else ""
        item_no_norm = _normalize_item_code(item_no)
        desc = item[1] if len(item) > 1 else ""
        qty = item[2] if len(item) > 2 else ""
        unit_price = item[3] if len(item) > 3 else ""
        mapped_item_code = ""
        mapped_item_desc = ""

        if master_lookup:
            po_key = _normalize_po(headers.get("po_number", ""))
            key = (po_key, item_no_norm)
            if key in master_lookup:
                mapped_item_code, mapped_item_desc = master_lookup[key]
                if diagnostics is not None:
                    diagnostics.append({
                        "po": po_key,
                        "supplier_item_code": item_no_norm,
                        "matched": True,
                        "orion_item_code": mapped_item_code,
                        "pi_item_desc": mapped_item_desc,
                        "supplier_matches": (supplier_counts.get(key, 1) if supplier_counts else 1),
                        "orion_matches": (orion_counts.get((po_key, _normalize_item_code(mapped_item_code)), 1) if orion_counts else 1),
                    })
            else:
                # Fallback: allow prefix match on supplier item code within same PO
                # Example: master has '706-12539-ABC' and PDF has '706-12539'
                # Also allow flexible PO matching: master PO and PDF PO can be prefix-compatible
                def po_flex_match(master_po: str, pdf_po: str) -> bool:
                    return bool(master_po) and bool(pdf_po) and (master_po.startswith(pdf_po) or pdf_po.startswith(master_po))

                candidates: List[Tuple[Tuple[str, str], Tuple[str, str]]] = [
                    (k, v) for k, v in master_lookup.items()
                    if po_flex_match(k[0], po_key) and (k[1].startswith(item_no_norm) or item_no_norm.startswith(k[1]))
                ]
                if candidates:
                    # Prefer the longest supplier code match to reduce ambiguity
                    candidates.sort(key=lambda kv: len(kv[0][1]), reverse=True)
                    chosen_key, (mapped_item_code, mapped_item_desc) = candidates[0]
                    supplier_match_count = len(candidates)
                    if diagnostics is not None:
                        diagnostics.append({
                            "po": po_key,
                            "supplier_item_code": item_no_norm,
                            "matched": True,
                            "orion_item_code": mapped_item_code,
                            "pi_item_desc": mapped_item_desc,
                            "supplier_matches": supplier_match_count,
                            "orion_matches": (orion_counts.get((po_key, _normalize_item_code(mapped_item_code)), 1) if orion_counts else 1),
                        })
                else:
                    if diagnostics is not None:
                        diagnostics.append({
                            "po": po_key,
                            "supplier_item_code": item_no_norm,
                            "matched": False,
                            "orion_item_code": "",
                            "pi_item_desc": "",
                            "supplier_matches": 0,
                            "orion_matches": 0,
                        })
        row = [
            "PO",  # PO Txn Code
            headers.get("po_number", ""),
            headers.get("dell_order_no", ""),
            headers.get("invoice_date", ""),
            headers.get("customer_no", "") or headers.get("ed_order", "") or headers.get("dell_order_no", ""),
            headers.get("shipping_method", ""),  # AWB per spec
            "N/A",  # Bill of leading date
            "N/A",  # Shipping Agent
            "N/A",  # From Port
            "N/A",  # To Port
            tomorrow_date,  # ETS
            "",  # ETA
            mapped_item_code,  # Item Code (internal)
            mapped_item_desc,  # Item Desc (internal)
            "NOS",  # UOM
            qty,
            unit_price,
            item_no,
            desc,
            headers.get("consolidation_fee_usd", ""),
        ]
        rows.append(row)
    return rows


def debug_consolidation(pdf_path) -> Dict[str, Any]:
    """Return diagnostic info for Consolidation parsing: nearby lines and numbers."""
    info: Dict[str, Any] = {"matched_line": "", "next_line": "", "numbers": [], "value": ""}
    with pdfplumber.open(pdf_path) as pdf:
        raw_text_parts: List[str] = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                raw_text_parts.append("\n".join([x.strip() for x in t.splitlines()]))
        raw_full_text = "\n".join(raw_text_parts)
        raw_lines = raw_full_text.splitlines()
        for i, line in enumerate(raw_lines):
            if "consolidation" in line.lower():
                info["matched_line"] = line
                if i + 1 < len(raw_lines):
                    info["next_line"] = raw_lines[i + 1]
                nums = re.findall(r"([0-9][0-9,]*\.[0-9]{2}|[0-9][0-9,]*)", line)
                if not nums and i + 1 < len(raw_lines):
                    nums = re.findall(r"([0-9][0-9,]*\.[0-9]{2}|[0-9][0-9,]*)", raw_lines[i + 1])
                info["numbers"] = nums
                if nums:
                    info["value"] = nums[-1]
                break
    return info



