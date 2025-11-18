import re
import streamlit as st
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
            def as_float(s: str) -> Optional[float]:
                try:
                    return float(str(s).replace(",", "").strip())
                except Exception:
                    return None

            pdf_unit_price_val = as_float(unit_price)
            pdf_qty_val = as_float(qty)

            # Build candidate lists from supplier_index (exact or flexible)
            debug_steps: List[str] = []
            exact_entries = (supplier_index.get(key, []) if supplier_index else [])
            debug_steps.append(f"Step 1: Exact supplier key match (PO={po_key}, Supplier={item_no_norm}) -> {len(exact_entries)} record(s)")
            flex_entries: List[Tuple[str, str, str, str]] = []
            if not exact_entries and supplier_index:
                def po_flex_match(master_po: str, pdf_po: str) -> bool:
                    return bool(master_po) and bool(pdf_po) and (master_po.startswith(pdf_po) or pdf_po.startswith(master_po))
                for (kpo, ksupp), entries in supplier_index.items():
                    if po_flex_match(kpo, po_key) and (ksupp.startswith(item_no_norm) or item_no_norm.startswith(ksupp)):
                        flex_entries.extend(entries)
                debug_steps.append(f"Step 2: Flexible supplier/PO match -> {len(flex_entries)} record(s)")

            supplier_candidates = exact_entries if exact_entries else flex_entries
            total_supplier_matches = len(supplier_candidates)
            matching_mode = "exact" if exact_entries else ("flex" if flex_entries else "none")

            if total_supplier_matches == 1:
                # Case A
                mapped_item_code, mapped_item_desc, out_orion_unit_price, out_orion_qty = supplier_candidates[0]
                status = "A_single"
                highlight = "none"
                debug_steps.append("Case A: Single supplier match found -> Output M/N with Orion code & desc")
                matched_by = "supplier-exact" if matching_mode == "exact" else "supplier-flex"
                chosen_orion_code_minimal = mapped_item_code
            
            
            elif total_supplier_matches > 1:
                            # Case B - multiple supplier entries for same (po, supplier)
                            # 🔍 DEBUG (Shows in Streamlit UI)
                            st.write("---- DEBUG FOR CASE B ----")
                            st.write("PDF Unit Price:", unit_price)
                            st.write("PDF Qty:", qty)
            
                            for e in supplier_candidates:
                                st.write({
                                    "candidate_item_code": e[0],
                                    "candidate_unit_price": e[2],
                                    "candidate_qty": e[3],
                                    "same_price?": (as_float(e[2]) == pdf_unit_price_val) if pdf_unit_price_val is not None else False,
                                    "same_qty?": (as_float(e[3]) == pdf_qty_val) if pdf_qty_val is not None else False
                                })
            
                            # Matches by price+qty and by price-only
                            # Matches by price+qty and by price-only
                            price_qty_matched = [
                                e for e in supplier_candidates
                                if pdf_unit_price_val is not None and as_float(e[2]) == pdf_unit_price_val
                                and pdf_qty_val is not None and as_float(e[3]) == pdf_qty_val
                            ]
                            
                            price_matched = [
                                e for e in supplier_candidates
                                if pdf_unit_price_val is not None and as_float(e[2]) == pdf_unit_price_val
                            ]
                            
                            rates_list = ", ".join([e[2] or "" for e in supplier_candidates])
                            debug_steps.append(f"Case B: Multiple supplier matches ({total_supplier_matches}). PDF unit price={unit_price}. Candidate rates=[{rates_list}]")
                            debug_steps.append(f"Price-only candidates: {len(price_matched)}; Price+Qty candidates: {len(price_qty_matched)}")
                            
                            if len(price_qty_matched) == 1:
                                # Perfect match
                                e = price_qty_matched[0]
                                out_orion_unit_price = e[2]
                                out_orion_qty = e[3]
                                out_orion_item_code = e[0]
                                mapped_item_code = ""
                                mapped_item_desc = ""
                                status = "B_price_qty_single"
                                highlight = "yellow"
                                matched_by = "supplier-" + matching_mode + "+price+qty"
                                chosen_orion_code_minimal = out_orion_item_code
                                debug_steps.append("Price+Qty match success: exactly 1 -> Output U/V/W; keep M/N empty")
                            
                            elif len(price_qty_matched) == 0 and len(price_matched) == 1:
                                # Price matches but qty differs
                                e = price_matched[0]
                                out_orion_unit_price = e[2]
                                out_orion_qty = e[3]
                                out_orion_item_code = e[0]
                                mapped_item_code = ""
                                mapped_item_desc = ""
                                status = "B_price_qty_mismatch"
                                highlight = "yellow"
                                matched_by = "supplier-" + matching_mode + "+price_only_qty_diff"
                                chosen_orion_code_minimal = out_orion_item_code
                                debug_steps.append(f"Price matches but qty differs: PDF {pdf_qty_val}, candidate {e[3]}")
                            
                            elif len(price_matched) > 1:
                                # Multiple candidates with same price -> ambiguous
                                status = "B_multi_price_matches"
                                highlight = "yellow"
                                mapped_item_code = ""
                                mapped_item_desc = ""
                                debug_steps.append(f"Ambiguous price-only matches: {len(price_matched)} candidates -> highlight M/N yellow; no output U/V/W")
                            
                            else:
                                # No matches at all
                                status = "B_no_price_qty_or_price_match"
                                highlight = "yellow"
                                mapped_item_code = ""
                                mapped_item_desc = ""
                                debug_steps.append("No price or price+qty matches -> highlight M/N yellow; no output U/V/W")
                
                
                
                
            else:
                # Case C - no supplier match
                highlight = "red"
                status = "C_no_supplier_match"
                debug_steps.append("Case C: No supplier match -> Highlight M/N red. Try Orion code + price.")
                # Try by Orion item code + price
                okey = (po_key, item_no_norm)
                o_candidates = orion_index.get(okey, []) if orion_index else []
                price_matched = [e for e in o_candidates if pdf_unit_price_val is not None and as_float(e[2]) == pdf_unit_price_val]
               
               
               
               
                debug_steps.append(f"Orion match attempt: Candidates by Orion code = {len(o_candidates)}; price matches = {len(price_matched)}")
                if len(price_matched) == 1:
                    e = price_matched[0]
                    out_orion_unit_price = e[2]
                    out_orion_qty = e[3]
                    out_orion_item_code = e[0]
                    mapped_item_code = ""
                    mapped_item_desc = ""
                    status = "C_orion_price_single"
                    debug_steps.append("Orion+price match success: exactly 1 -> Output U/V/W; keep M/N empty; keep red highlight")
                    matched_by = "orion+price"
                    chosen_orion_code_minimal = out_orion_item_code
                else:
                    # New fallback: PO + price (ignore item codes)
                    po_candidates = po_price_index.get(po_key, []) if po_price_index else []
                    po_price_matched = [e for e in po_candidates if pdf_unit_price_val is not None and as_float(e[2]) == pdf_unit_price_val]
                    debug_steps.append(f"PO+price attempt: Candidates in PO = {len(po_candidates)}; price matches = {len(po_price_matched)}")
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
                        debug_steps.append("PO+price match success: exactly 1 -> Output U/V/W; keep M/N empty; keep red highlight")
                    else:
                        status = "C_no_price_or_multi" if len(price_matched) != 1 else status
                        if len(po_price_matched) == 0:
                            debug_steps.append("PO+price match failure: 0 matches -> Keep red highlight; no output")
                        else:
                            debug_steps.append(f"PO+price ambiguous: {len(po_price_matched)} matches -> Keep red highlight; no output")

            if diagnostics is not None:
                fill_MN = bool(mapped_item_code or mapped_item_desc)
                fill_UVW = bool(out_orion_item_code or out_orion_unit_price or out_orion_qty)
                diagnostics.append({
                    "po": po_key,
                    "supplier_item_code": item_no_norm,
                    "pdf_unit_price": unit_price,
                    "pdf_unit_price_num": (pdf_unit_price_val if pdf_unit_price_val is not None else ""),
                    "status": status,
                    "highlight": highlight,
                    "mapped_item_code": mapped_item_code,
                    "mapped_item_desc": mapped_item_desc,
                    "out_orion_unit_price": out_orion_unit_price,
                    "out_orion_qty": out_orion_qty,
                    "out_orion_item_code": out_orion_item_code,
                    "total_supplier_matches": total_supplier_matches,
                    "matching_mode": matching_mode,
                    "supplier_candidate_rates": ", ".join([e[2] or "" for e in supplier_candidates]) if supplier_candidates else "",
                    "price_match_count": (len(price_matched) if 'price_matched' in locals() else 0),
                    "orion_candidate_count": (len(o_candidates) if 'o_candidates' in locals() and o_candidates is not None else 0),
                    "fill_MN": fill_MN,
                    "fill_UVW": fill_UVW,
                    "message": " | ".join(debug_steps),
                })
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



