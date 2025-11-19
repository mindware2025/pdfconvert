import re
from typing import List, Optional, Dict, Any, Tuple

import pdfplumber
import pandas as pd

from utils.helpers import normalize_line
from datetime import datetime, timedelta


today_plus_10 = (datetime.today() + timedelta(days=10)).strftime("%m/%d/%Y")


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
    "Orion Unit Price",
    "Orion qty",
    "Orion Item code",
    "Matched By",
    "Chosen Orion Item Code",
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
    Dict[Tuple[str, str], List[Tuple[str, str, str, str]]],
    Dict[Tuple[str, str], List[Tuple[str, str, str, str]]],
    Dict[str, List[Tuple[str, str, str, str, str]]],
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
    # Optional columns
    try:
        c_unit_rate = col("Po Unit Rate")
    except Exception:
        c_unit_rate = None
    try:
        c_qty = col("Qty")
    except Exception:
        try:
            c_qty = col("Po Qty")
        except Exception:
            c_qty = None

    lookup: Dict[Tuple[str, str], Tuple[str, str]] = {}
    supplier_counts: Dict[Tuple[str, str], int] = {}
    orion_counts: Dict[Tuple[str, str], int] = {}
    supplier_index: Dict[Tuple[str, str], List[Tuple[str, str, str, str]]] = {}
    orion_index: Dict[Tuple[str, str], List[Tuple[str, str, str, str]]] = {}
    po_price_index: Dict[str, List[Tuple[str, str, str, str, str]]] = {}
    for _, r in df.iterrows():
        po = _normalize_po(r.get(c_po, ""))
        supp = _normalize_item_code(r.get(c_supplier, ""))
        orion = str(r.get(c_orion, "") or "").strip()
        pi_desc = str(r.get(c_pi_desc, "") or "").strip()
        unit_rate = str(r.get(c_unit_rate, "") or "").strip() if c_unit_rate else ""
        qty = str(r.get(c_qty, "") or "").strip() if c_qty else ""
        if po and supp:
            key = (po, supp)
            lookup[key] = (orion, pi_desc)
            supplier_counts[key] = supplier_counts.get(key, 0) + 1
            supplier_index.setdefault(key, []).append((orion, pi_desc, unit_rate, qty))
        if po and orion:
            okey = (po, _normalize_item_code(orion))
            orion_counts[okey] = orion_counts.get(okey, 0) + 1
            orion_index.setdefault(okey, []).append((orion, pi_desc, unit_rate, qty))
        if po:
            po_price_index.setdefault(po, []).append((orion, pi_desc, unit_rate, qty, supp))
    return lookup, supplier_counts, orion_counts, supplier_index, orion_index, po_price_index


def build_pre_alert_rows(
    pdf_path,
    tomorrow_date: str,
    master_lookup: Optional[Dict[Tuple[str, str], Tuple[str, str]]] = None,
    supplier_counts: Optional[Dict[Tuple[str, str], int]] = None,
    orion_counts: Optional[Dict[Tuple[str, str], int]] = None,
    supplier_index: Optional[Dict[Tuple[str, str], List[Tuple[str, str, str, str]]]] = None,
    orion_index: Optional[Dict[Tuple[str, str], List[Tuple[str, str, str, str]]]] = None,
    po_price_index: Optional[Dict[str, List[Tuple[str, str, str, str, str]]]] = None,
    diagnostics: Optional[List[Dict[str, Any]]] = None,
) -> List[List[Any]]:
    """Build rows for the PRE ALERT UPLOAD sheet from a single PDF.
    Enhanced heavy debug/logging version.
    """
    headers = extract_header_fields(pdf_path)
    items = extract_table_from_text(pdf_path)
    rows: List[List[Any]] = []

    for idx_item, item in enumerate(items):
        # Very verbose debug collector for this item
        debug_steps: List[str] = []
        try:
            debug_steps.append(f"Processing item index={idx_item} raw_item={item!r}")
            item_no = item[0] if len(item) > 0 else ""
            item_no_norm = _normalize_item_code(item_no)
            desc = item[1] if len(item) > 1 else ""
            qty = item[2] if len(item) > 2 else ""
            unit_price = item[3] if len(item) > 3 else ""
            debug_steps.append(f"Normalized item code: '{item_no_norm}' desc='{desc}' qty='{qty}' unit_price='{unit_price}'")

            mapped_item_code = ""
            mapped_item_desc = ""
            out_orion_unit_price = ""
            out_orion_qty = ""
            out_orion_item_code = ""
            matched_by = "none"
            chosen_orion_code_minimal = ""
            highlight = "none"
            status = ""

            if master_lookup:
                po_key = _normalize_po(headers.get("po_number", ""))
                key = (po_key, item_no_norm)
                debug_steps.append(f"PO key='{po_key}', lookup key={key!r}")

                def as_float(s: str) -> Optional[float]:
                    try:
                        return float(str(s).replace(",", "").strip())
                    except Exception:
                        return None

                pdf_unit_price_val = as_float(unit_price)
                pdf_qty_val = None
                try:
                    pdf_qty_val = float(str(qty).replace(",", "").strip())
                except Exception:
                    pdf_qty_val = None

                debug_steps.append(f"Parsed numeric: pdf_unit_price_val={pdf_unit_price_val} pdf_qty_val={pdf_qty_val}")

                # Build candidate lists from supplier_index (exact or flexible)
                exact_entries = (supplier_index.get(key, []) if supplier_index else [])
                debug_steps.append(f"Exact matches from supplier_index for key {key}: count={len(exact_entries)}")
                if exact_entries:
                    debug_steps.extend([f"  exact[{i}]={e!r}" for i, e in enumerate(exact_entries)])

                # Always also gather flex entries (candidates where ksupp startswith/pdf startswith ksupp)
                flex_entries: List[Tuple[str, str, str, str]] = []
                if supplier_index:
                    def po_flex_match(master_po: str, pdf_po: str) -> bool:
                        return bool(master_po) and bool(pdf_po) and (master_po.startswith(pdf_po) or pdf_po.startswith(master_po))
                    for (kpo, ksupp), entries in supplier_index.items():
                        if po_flex_match(kpo, po_key) and (ksupp.startswith(item_no_norm) or item_no_norm.startswith(ksupp)):
                            flex_entries.extend(entries)
                    debug_steps.append(f"Flexible matches found: count={len(flex_entries)}")
                    if flex_entries:
                        debug_steps.extend([f"  flex[{i}]={e!r}" for i, e in enumerate(flex_entries)])

                # Combine exact and flex candidates (dedupe) so we don't miss close variants like 210-BDUK-LCA
                if exact_entries:
                    seen = set(exact_entries)
                    supplier_candidates = list(exact_entries) + [e for e in flex_entries if e not in seen]
                else:
                    supplier_candidates = flex_entries

                total_supplier_matches = len(supplier_candidates)
                matching_mode = "exact" if exact_entries else ("flex" if flex_entries else "none")
                debug_steps.append(f"Using supplier_candidates count={total_supplier_matches} mode={matching_mode}")
                if total_supplier_matches == 1:
                    # Case A
                    mapped_item_code, mapped_item_desc, out_orion_unit_price, out_orion_qty = supplier_candidates[0]
                    out_orion_item_code = mapped_item_code
                    status = "A_single"
                    highlight = "none"
                    debug_steps.append("Case A: Single supplier match -> use mapped_item_code/mapped_item_desc")
                    matched_by = "supplier-exact" if matching_mode == "exact" else "supplier-flex"
                    chosen_orion_code_minimal = mapped_item_code
                elif total_supplier_matches > 1:
                    # Case B
                    debug_steps.append("Case B: multiple supplier candidates, computing price matches")
                    price_matched = []
                    if pdf_unit_price_val is not None:
                        for e in supplier_candidates:
                            # entries are (orion, pi_desc, unit_rate, qty)
                            e_price = as_float(e[2])  # unit_rate
                            e_qty = as_float(e[3])    # qty
                            debug_steps.append(f"  candidate e={e!r} parsed_price={e_price} parsed_qty={e_qty}")
                            if e_price is not None and e_price == pdf_unit_price_val:
                                price_matched.append(e)
                    debug_steps.append(f"price_matched count={len(price_matched)} list={[p for p in price_matched]}")

                    if len(price_matched) == 1:
                        mapped_item_code, mapped_item_desc, out_orion_unit_price, out_orion_qty = price_matched[0]
                        out_orion_item_code = mapped_item_code
                        mapped_item_code = ""
                        mapped_item_desc = ""
                        status = "B_price_single"
                        highlight = "yellow"
                        debug_steps.append("Price match success: exactly 1 price_matched -> output U/V/W")
                        matched_by = "supplier-" + matching_mode + "+price"
                        chosen_orion_code_minimal = out_orion_item_code
                    else:
                        # Deterministic qty tie-breaker: only accept an exact qty match.
                        debug_steps.append("Multiple or zero price matches -> try exact qty tie-breaker")
                        picked = None
                        # 1) look for first exact qty among price_matched
                        if pdf_qty_val is not None and price_matched:
                            for i_e, e in enumerate(price_matched):
                                try:
                                    e_qty = float(str(e[3]).replace(",", "").strip()) if e[3] not in (None, "") else None
                                except Exception:
                                    e_qty = None
                                debug_steps.append(f"  checking price_matched[{i_e}] qty={e_qty} against pdf_qty={pdf_qty_val}")
                                if e_qty is not None and e_qty == pdf_qty_val:
                                    picked = e
                                    debug_steps.append(f"  -> picked exact qty among price_matched at index {i_e}: {e!r}")
                                    break

                        # 2) if not found, look for first exact qty among all supplier_candidates where price is within small tolerance
                        if picked is None and pdf_qty_val is not None and supplier_candidates:
                            TOL = 0.01
                            debug_steps.append(f"  No exact qty in price_matched; searching all supplier_candidates with tolerance={TOL}")
                            for i_e, e in enumerate(supplier_candidates):
                                try:
                                    e_price = float(str(e[2]).replace(",", "").strip()) if e[2] not in (None, "") else None
                                    e_qty = float(str(e[3]).replace(",", "").strip()) if e[3] not in (None, "") else None
                                except Exception:
                                    e_price = None
                                    e_qty = None
                                debug_steps.append(f"    checking supplier_candidates[{i_e}] price={e_price} qty={e_qty}")
                                if e_qty is not None and e_qty == pdf_qty_val and e_price is not None and pdf_unit_price_val is not None and abs(e_price - pdf_unit_price_val) <= TOL:
                                    picked = e
                                    debug_steps.append(f"    -> picked exact qty with tolerant price at index {i_e}: {e!r}")
                                    break

                        if picked is not None:
                            mapped_item_code, mapped_item_desc, out_orion_unit_price, out_orion_qty = picked
                            out_orion_item_code = mapped_item_code
                            mapped_item_code = ""
                            mapped_item_desc = ""
                            status = "B_price_qty_first"
                            highlight = "none"
                            debug_steps.append("Qty tie-break: exact qty found -> output U/V/W (no highlight)")
                            matched_by = "supplier-" + matching_mode + "+price+qty_first"
                            chosen_orion_code_minimal = out_orion_item_code
                        else:
                            # STOP: no exact qty -> mark ambiguous, do NOT use closest-qty fallback
                            status = "B_multi_price_matches"
                            highlight = "yellow"
                            mapped_item_code = ""
                            mapped_item_desc = ""
                            debug_steps.append("No exact qty found -> Ambiguous price matches -> STOP and mark yellow (no UVW output)")
                else:
                    # Case C - no supplier match
                    highlight = "red"
                    status = "C_no_supplier_match"
                    debug_steps.append("Case C: No supplier match -> Highlight M/N red. Try Orion code + price.")
                    # Try by Orion item code + price
                    okey = (po_key, item_no_norm)
                    o_candidates = orion_index.get(okey, []) if orion_index else []
                    debug_steps.append(f"Orion candidates for key {okey}: count={len(o_candidates)}")
                    if o_candidates:
                        for i_e, e in enumerate(o_candidates):
                            debug_steps.append(f"  orion_candidate[{i_e}]={e!r} parsed_price={as_float(e[2])} parsed_qty={as_float(e[3])}")
                    price_matched = [e for e in o_candidates if pdf_unit_price_val is not None and as_float(e[2]) == pdf_unit_price_val]
                    debug_steps.append(f"Orion price_matched count={len(price_matched)}")
                    if len(price_matched) == 1:
                        e = price_matched[0]
                        out_orion_unit_price = e[2]
                        out_orion_qty = e[3]
                        out_orion_item_code = e[0]
                        mapped_item_code = ""
                        mapped_item_desc = ""
                        status = "C_orion_price_single"
                        debug_steps.append("Orion+price match success -> output UVW, keep M/N red")
                        matched_by = "orion+price"
                        chosen_orion_code_minimal = out_orion_item_code
                    else:
                        # New fallback: PO + price (ignore item codes)
                        po_candidates = po_price_index.get(po_key, []) if po_price_index else []
                        debug_steps.append(f"PO price candidates for PO {po_key}: count={len(po_candidates)}")
                        po_price_matched = [e for e in po_candidates if pdf_unit_price_val is not None and as_float(e[2]) == pdf_unit_price_val]
                        debug_steps.append(f"PO+price matched count={len(po_price_matched)}")
                        if len(po_price_matched) == 1:
                            e = po_price_matched[0]
                            out_orion_unit_price = e[2]
                            out_orion_qty = e[3]
                            out_orion_item_code = e[0]
                            mapped_item_code = ""
                            mapped_item_desc = ""
                            status = "C_po_price_single"
                            matched_by = "po+price"
                            chosen_orion_code_minimal = out_orion_item_code
                            debug_steps.append("PO+price match success -> output UVW, keep M/N red")
                        else:
                            status = "C_no_price_or_multi" if len(price_matched) != 1 else status
                            if len(po_price_matched) == 0:
                                debug_steps.append("PO+price match failure: 0 matches -> Keep red highlight; no output")
                            else:
                                debug_steps.append(f"PO+price ambiguous: {len(po_price_matched)} matches -> Keep red highlight; no output")

            # Always attach diagnostics entry with the very verbose message
            if diagnostics is not None:
                fill_MN = bool(mapped_item_code or mapped_item_desc)
                fill_UVW = bool(out_orion_item_code or out_orion_unit_price or out_orion_qty)
                diagnostics.append({
                    "item_index": idx_item,
                    "po": po_key if 'po_key' in locals() else "",
                    "supplier_item_code": item_no_norm,
                    "pdf_unit_price": unit_price,
                    "pdf_unit_price_num": (pdf_unit_price_val if 'pdf_unit_price_val' in locals() else ""),
                    "pdf_qty_num": (pdf_qty_val if 'pdf_qty_val' in locals() else ""),
                    "status": status,
                    "highlight": highlight,
                    "mapped_item_code": mapped_item_code,
                    "mapped_item_desc": mapped_item_desc,
                    "out_orion_unit_price": out_orion_unit_price,
                    "out_orion_qty": out_orion_qty,
                    "out_orion_item_code": out_orion_item_code,
                    "total_supplier_matches": total_supplier_matches if 'total_supplier_matches' in locals() else 0,
                    "matching_mode": matching_mode if 'matching_mode' in locals() else "none",
                    "supplier_candidate_rates": ", ".join([str(e[2] or "") for e in supplier_candidates]) if 'supplier_candidates' in locals() and supplier_candidates else "",
                    "price_match_count": (len(price_matched) if 'price_matched' in locals() else 0),
                    "orion_candidate_count": (len(o_candidates) if 'o_candidates' in locals() and o_candidates is not None else 0),
                    "fill_MN": fill_MN,
                    "fill_UVW": fill_UVW,
                    "message": " | ".join(debug_steps),
                })

            # Print debug to console (live) for immediate inspection
            try:
                print(f"DEBUG-DIAG: build_pre_alert_rows item_index={idx_item}", flush=True)
                for m in debug_steps:
                    print("DEBUG-DIAG:", m, flush=True)
                print("DEBUG-DIAG: ---- end debug item ----", flush=True)   
               
            except Exception:
                # avoid crashing on print errors
                pass

        except Exception as exc:
            # Ensure one item's exception does not break whole run; log it in diagnostics/console
            err_msg = f"EXCEPTION processing item idx={idx_item}: {exc}"
            try:
                print(err_msg)
            except Exception:
                pass
            if diagnostics is not None:
                diagnostics.append({"item_index": idx_item, "error": err_msg})

        # Build output row (keep same structure)
        row = [
            "PO",  # PO Txn Code
            headers.get("po_number", ""),
            headers.get("dell_order_no", ""),
            headers.get("invoice_date", ""),
            headers.get("customer_no", "") or headers.get("ed_order", "") or headers.get("dell_order_no", ""),
            ""  ,# AWB per spec
               today_plus_10,  # Bill of leading date
            "N/A",  # Shipping Agent
            "N/A",  # From Port
            "N/A",  # To Port
               today_plus_10,  # ETS
               today_plus_10,  # ETA (kept the same as ETS)
            mapped_item_code,  # Item Code (internal)
            mapped_item_desc,  # Item Desc (internal)
            "NOS",  # UOM
            qty,
            unit_price,
            item_no,
            desc,
            headers.get("consolidation_fee_usd", ""),
            out_orion_unit_price,
            out_orion_qty,
            out_orion_item_code,
            matched_by,
            chosen_orion_code_minimal,
        ]
        rows.append(row)
    return rows
# ...existing code...


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



