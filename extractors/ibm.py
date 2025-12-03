# ibm.py
import os
import re
import logging
from datetime import datetime
from io import BytesIO
import pandas as pd
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from terms_template import get_terms_section

# ----------------------------------------------------------------------
# Debug info collector for live environment
# ----------------------------------------------------------------------
debug_info = []

def add_debug(message):
    """Add debug info that can be displayed in Streamlit"""
    debug_info.append(message)
    if len(debug_info) > 300:  # Keep more debug messages
        debug_info.pop(0)

def get_debug_info():
    """Get collected debug info"""
    return debug_info.copy()

def clear_debug():
    """Clear debug info"""
    debug_info.clear()

# ----------------------------------------------------------------------
# Simplified logging for live environment
# ----------------------------------------------------------------------
logging.basicConfig(level=logging.ERROR)  # Only log errors in live

# ----------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------
USD_TO_AED = 3.6725

# ----------------------------------------------------------------------
# Number parsers
# ----------------------------------------------------------------------
def parse_euro_number(value: str):
    """
    Parse EU-formatted numbers like:
    - '733,00'         -> 733.00
    - '114.030,00'     -> 114030.00
    - '60,770'         -> 60.770
    """
    try:
        if value is None:
            return None
        s = str(value).strip().replace(" ", "")
        if "." in s and "," in s:
            if s.rfind(",") > s.rfind("."):
                # thousands '.', decimal ','
                s = s.replace(".", "").replace(",", ".")
            else:
                # thousands ',', decimal '.'
                s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
        return float(s)
    except Exception:
        return None

# ----------------------------------------------------------------------
# Debug function for data integrity
# ----------------------------------------------------------------------
def debug_extracted_data(extracted_data):
    """Debug function to check data integrity"""
    add_debug(f"[DATA DEBUG] Total extracted rows: {len(extracted_data)}")
    for i, row in enumerate(extracted_data):
        add_debug(f"[DATA ROW {i+1}] Length: {len(row)}, SKU: '{row[0]}', Desc: '{row[1][:30]}...'")
        if len(row) != 7:
            add_debug(f"[DATA ERROR] Row {i+1} has {len(row)} columns instead of 7!")
        if not row[0]:  # Check if SKU is empty
            add_debug(f"[DATA ERROR] Row {i+1} has empty SKU!")
        if i == 10:  # Row 11 specifically
            add_debug(f"[ROW 11 SPECIFIC] Full content: {row}")
    return extracted_data

# ----------------------------------------------------------------------
# Description correction
# ----------------------------------------------------------------------
def correct_descriptions(extracted_data, master_data=None):
    """
    Each row: [sku, desc, qty, start_date, end_date, bid_unit_svp_aed, bid_ext_svp_aed]
    Description policy:
    - If master_data uploaded: Use ONLY master CSV descriptions (blank if SKU not found)
    - If master_data NOT uploaded: Set ALL descriptions to blank
    - Never use PDF descriptions
    """
    corrected = []
    
    add_debug(f"[DESC CORRECTION] Starting with {len(extracted_data)} rows")
    
    # First debug the data before correction
    debug_extracted_data(extracted_data)
    
    if master_data is not None:
        try:
            master_map = dict(zip(master_data['SKU'], master_data['SKU DESCRIPTION']))
            add_debug(f"[MASTER DATA] Using master CSV with {len(master_map)} SKU mappings")
            
            for i, row in enumerate(extracted_data):
                try:
                    sku = row[0]
                    if sku in master_map:
                        row[1] = master_map[sku]
                        add_debug(f"[DESC FROM MASTER] Row {i+1} - SKU '{sku}': Found in master CSV")
                    else:
                        row[1] = ""  # Blank if SKU not found in master
                        add_debug(f"[DESC BLANK] Row {i+1} - SKU '{sku}' not found in master CSV")
                        
                except Exception as e:
                    add_debug(f"[DESC ERROR] Row {i+1} - Error: {e}")
                    row[1] = ""
                    
                corrected.append(row)
                
        except Exception as e:
            add_debug(f"[MASTER DATA ERROR] Could not process master data: {e}")
            for row in extracted_data:
                row[1] = ""
            corrected = extracted_data
    else:
        add_debug(f"[NO MASTER DATA] Setting all descriptions to blank")
        for i, row in enumerate(extracted_data):
            row[1] = ""
        corrected = extracted_data
    
    add_debug(f"[DESC CORRECTION COMPLETE] Processed {len(corrected)} rows")
    
    # Debug after correction
    add_debug(f"[AFTER CORRECTION] Row 11: {corrected[10] if len(corrected) > 10 else 'N/A'}")
    
    return corrected

# ----------------------------------------------------------------------
# Enhanced regexes
# ----------------------------------------------------------------------
date_re = re.compile(r'\b\d{2}[‐‑–-][A-Za-z]{3}[‐‑–-]\d{4}\b')  # hyphen variants
sku_line_re = re.compile(r'^[A-Z0-9\-\._/]{5,20}$')
token_sku_re = re.compile(r'\b[A-Z0-9\-\._/]{5,20}\b')
int_re = re.compile(r'\b\d+\b')
money_with_sep_re = re.compile(r'\d[\d.,]*[.,]\d+')

header_blacklist_re = re.compile(
    r'\b('
    r'Bid Number|Bid Request|Opportunity Number|Distributor|Distributor Name|Supplier|'
    r'Page \d+ of|IBM Ireland Product Distribution Limited|IBM Terms|Customer Information'
    r')\b', re.I
)

def looks_like_valid_sku(tok: str) -> bool:
    """Enhanced SKU validation for IBM part numbers"""
    if not tok:
        return False
    
    # CRITICAL: Explicitly reject serial numbers that start with IE and are long
    if tok.startswith('IE') and len(tok) > 8:
        return False
    
    # Skip pure numbers
    if tok.isdigit():
        return False
    
    # Basic pattern check
    if not sku_line_re.match(tok):
        return False
    
    # Must contain at least one letter and one digit
    if not re.search(r'[A-Z]', tok):
        return False
    if not re.search(r'\d', tok):
        return False
    
    # IBM SKUs are typically 7-8 characters, allow some flexibility
    if not (6 <= len(tok) <= 9):
        return False
    
    # Additional IBM-specific patterns (most start with D0, Y0, etc.)
    if re.match(r'^[A-Z]\d[A-Z0-9]{5,7}$', tok):
        return True
    
    return False

# ----------------------------------------------------------------------
# Qty inference (ANY qty, no small-number assumptions)
# ----------------------------------------------------------------------
def _pick_qty_from_candidates(candidates, unit, ext, abs_tol=0.02):
    """
    Choose qty q in candidates that minimizes |ext - unit*q|.
    Accept if min absolute error <= abs_tol (≈ 2 decimal rounding).
    If no candidates fit, try round(ext/unit).
    """
    if unit is None or ext is None or unit <= 0:
        return None
    best_q, best_err = None, float("inf")
    for q in candidates:
        err = abs(ext - unit * q)
        if err < best_err:
            best_q, best_err = q, err
    if best_q is not None and best_err <= abs_tol:
        return best_q
    # Fallback by division
    q_est = int(round(ext / unit))
    if q_est > 0 and abs(ext - unit * q_est) <= abs_tol:
        return q_est
    return None

def infer_qty_and_prorate(after_end: str, abs_tol=0.02):
    """
    Infer Qty using Entitled Ext ≈ Qty * Entitled Unit (to cent rounding).
    Split using STRICT money regex (must contain ',' or '.') so plain ints remain as ints.
    Returns: (qty:int|None, prorate:int|None, money_tokens:list[str])
    """
    # 1) First strict money token
    m_first = money_with_sep_re.search(after_end)
    if not m_first:
        return None, None, []
    # 2) Pre-money integers zone and money zone
    ints_zone = after_end[:m_first.start()]
    money_zone = after_end[m_first.start():]
    ints = [int(x) for x in int_re.findall(ints_zone)]
    tokens = money_with_sep_re.findall(money_zone)
    qty = None
    # Try Entitled pair first (tokens[0]=Entitled Unit, tokens[1]=Entitled Ext)
    if len(tokens) >= 2:
        m0 = parse_euro_number(tokens[0])
        m1 = parse_euro_number(tokens[1])
        qty = _pick_qty_from_candidates(ints, m0, m1, abs_tol=abs_tol)
    # Then try Bid pair (tokens[3], tokens[4]) if needed
    if qty is None and len(tokens) >= 5:
        bu = parse_euro_number(tokens[3])
        be = parse_euro_number(tokens[4])
        qty = _pick_qty_from_candidates(ints, bu, be, abs_tol=abs_tol)
    # Last resort: if we have Entitled pair, try division only
    if qty is None and len(tokens) >= 2:
        m0 = parse_euro_number(tokens[0])
        m1 = parse_euro_number(tokens[1])
        if m0 and m1:
            q_est = int(round(m1 / m0))
            if q_est > 0 and abs(m1 - m0 * q_est) <= abs_tol:
                qty = q_est
    # Optional: deduce prorate as the first different int in pre-money zone
    prorate = None
    if ints and qty is not None:
        for n in ints:
            if n != qty:
                prorate = n
                break
    return qty, prorate, tokens

# ----------------------------------------------------------------------
# Helpers to handle wrapped rows (extend after_end)
# ----------------------------------------------------------------------
def _extend_after_end_with_following_lines(lines, start_idx, window, max_extra_lines=8):
    """
    If the chunk cut off the amounts, extend the 'after_end' text with a few following
    lines to capture remaining tokens (Disc%, Bid Unit, Bid Ext). Stop early if we see
    strong signs of a new section/item; otherwise just append.
    """
    ext_parts = []
    for j in range(start_idx + window, min(start_idx + window + max_extra_lines, len(lines))):
        ln = lines[j].strip()
        # Heuristics to cautiously stop if a new section is likely
        if header_blacklist_re.search(ln):
            break
        ext_parts.append(" " + ln)
    return "".join(ext_parts)

# ----------------------------------------------------------------------
# Core PDF extraction
# ----------------------------------------------------------------------
def extract_ibm_data_from_pdf(file_like) -> tuple[list, dict]:
    """
    Extracts line items and header info from an IBM Quotation PDF.
    Args:
        file_like: PDF file stream
    Returns:
      - extracted_data: list of rows
          [sku, desc, qty, start_date, end_date, bid_unit_svp_aed, bid_ext_svp_aed]
      - header_info: dict of customer/bid metadata
    """
    clear_debug()  # Clear previous debug info
    
    # Open PDF
    doc = fitz.open(stream=file_like.read(), filetype="pdf")
    # Collect lines
    lines = []
    for page in doc:
        page_text = page.get_text("text") or page.get_text()
        for l in page_text.splitlines():
            if l and l.strip():
                lines.append(l.rstrip())
    
    add_debug(f"[PDF INFO] Total lines extracted: {len(lines)}")
    
    # Header fields
    header_info = {
        "Customer Name": "",
        "Bid Number": "",
        "PA Agreement Number": "",
        "PA Site Number": "",
        "Select Territory": "",
        "Government Entity (GOE)": "",
        "IBM Opportunity Number": "",
        "Reseller Name": "",
        "City": "",
        "Country": ""
    }
    
    # Parse header info (simple look-ahead by 1 line)
    for i, line in enumerate(lines):
        if "Customer Name:" in line:
            header_info["Customer Name"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "Reseller Name:" in line:
            header_info["Reseller Name"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "Bid Number:" in line:
            header_info["Bid Number"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "PA Agreement Number:" in line:
            header_info["PA Agreement Number"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "PA Site Number:" in line:
            header_info["PA Site Number"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "Select Territory:" in line:
            header_info["Select Territory"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "Government Entity" in line:
            header_info["Government Entity (GOE)"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "IBM Opportunity Number:" in line:
            header_info["IBM Opportunity Number"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "City:" in line:
            header_info["City"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if "Country:" in line:
            header_info["Country"] = lines[i + 1].strip() if i + 1 < len(lines) else ""
    
    extracted_data = []
    i = 0
    max_window = 12  # Try wider chunks first to capture wrapped rows
    processed_positions = set()  # Track processed line positions to avoid duplicates
    
    add_debug(f"[EXTRACTION START] Beginning extraction from {len(lines)} lines")
    
    while i < len(lines):
        matched = False
        
        # Prefer larger chunks first (helps capture qty + amounts in one chunk)
        for window in range(max_window, 0, -1):
            if i + window > len(lines):
                continue
            chunk_lines = lines[i:i + window]
            chunk = " | ".join(chunk_lines)
            
            # Skip header-ish chunks
            if header_blacklist_re.search(chunk):
                continue
                
            # Must have at least two date tokens
            dates = date_re.findall(chunk)
            if len(dates) < 2:
                continue
            start_date, end_date = dates[0], dates[1]
            
            # ENHANCED SKU identification - look for the ACTUAL SKU, not serial numbers
            sku = None
            desc_start_index = None
            
            add_debug(f"[CHUNK ANALYSIS] Lines {i}-{i+window}: {[line.strip() for line in chunk_lines]}")
            
            # Strategy: Find all valid SKUs, then pick the one that's NOT a serial number
            valid_skus_found = []
            for line_idx, line in enumerate(chunk_lines):
                line = line.strip()
                
                # Skip obvious serial/row numbers  
                if line.isdigit():
                    continue
                
                # Find SKU patterns in this line
                sku_candidates = token_sku_re.findall(line)
                for candidate in sku_candidates:
                    if looks_like_valid_sku(candidate):
                        # Additional check: reject obvious serial numbers
                        if candidate.startswith('IE') and len(candidate) > 8:
                            add_debug(f"[SKU REJECT] Serial number: '{candidate}' in line {line_idx}")
                            continue
                        
                        valid_skus_found.append((candidate, line_idx, line))
                        add_debug(f"[SKU CANDIDATE] Found '{candidate}' in line {line_idx}: '{line}'")

            # Pick the BEST SKU (prefer shorter, IBM-style part numbers)
            if valid_skus_found:
                # Sort by: 1) Not starting with 'IE', 2) Length (shorter preferred), 3) Position
                def sku_priority(sku_info):
                    candidate, line_idx, line = sku_info
                    starts_with_ie = candidate.startswith('IE')
                    length = len(candidate)
                    return (starts_with_ie, length, line_idx)
                
                best_sku_info = sorted(valid_skus_found, key=sku_priority)[0]
                sku, sku_line_idx, _ = best_sku_info
                desc_start_index = sku_line_idx + 1  # Description starts after SKU line
                add_debug(f"[SKU SELECTED] Best SKU: '{sku}' from {len(valid_skus_found)} candidates")
            else:
                add_debug(f"[SKU NOT FOUND] No valid SKUs in chunk lines {i}-{i+window}")
                continue

            # Additional validation: Avoid processing same position multiple times 
            position_key = f"{i}_{window}"
            if position_key in processed_positions:
                add_debug(f"[POSITION SKIP] Already processed position {i}-{i+window}")
                continue
            processed_positions.add(position_key)
            
            # Keep only chunks where a money token is "near" the start date (reduces false positives)
            pos_date = chunk.find(start_date)
            near_money = False
            if pos_date >= 0:
                window_text = chunk[max(0, pos_date - 60): pos_date + len(start_date) + 60]
                if money_with_sep_re.search(window_text):
                    near_money = True
            if not near_money:
                continue
            
            # Enhanced description extraction with cleaning
            if desc_start_index is not None and desc_start_index < len(chunk_lines):
                desc_parts = []
                for ln in chunk_lines[desc_start_index:]:
                    # Stop at date patterns
                    if date_re.search(ln):
                        break
                    # Clean the line
                    clean_line = ln.strip()
                    # Remove leading/trailing pipes and extra characters
                    clean_line = re.sub(r'^\|?\s*', '', clean_line)
                    clean_line = re.sub(r'\s*\|?\s*$', '', clean_line)
                    if clean_line and not clean_line.isdigit():  # Skip digit-only lines
                        desc_parts.append(clean_line)
                
                desc = " ".join(desc_parts).strip()
                # Remove any remaining pipe characters and clean up
                desc = re.sub(r'\s*\|\s*', ' ', desc)
                desc = re.sub(r'\s+', ' ', desc).strip()
                add_debug(f"[DESC CLEANED] SKU '{sku}': '{desc[:50]}...'")
            else:
                # Fallback description extraction
                pos_sku = chunk.find(sku)
                pos_date0 = chunk.find(start_date)
                desc = chunk[pos_sku + len(sku):pos_date0].strip() if pos_sku >= 0 and pos_date0 > pos_sku else ""
                desc = re.sub(r'\s*\|\s*', ' ', desc)
                desc = re.sub(r'\s+', ' ', desc).strip()
                add_debug(f"[DESC FALLBACK] SKU '{sku}': '{desc[:50]}...'")
            
            # ---- Robust Qty inference (ANY value) ----
            chunk_flat = " ".join(chunk_lines)
            all_date_matches = list(date_re.finditer(chunk_flat))
            if len(all_date_matches) < 2:
                continue
            end_date_match = all_date_matches[1]
            after_end = chunk_flat[end_date_match.end():].strip()
            
            # 1) First pass qty + tokens
            qty, prorate, money_tokens = infer_qty_and_prorate(after_end, abs_tol=0.02)
            add_debug(f"[QTY] sku={sku} qty={qty}, money_tokens={len(money_tokens)}")
            
            # 2) If we didn't get all money tokens, extend with a few following lines
            if len(money_tokens) < 5:
                extra = _extend_after_end_with_following_lines(lines, i, window, max_extra_lines=8)
                if extra:
                    qty2, prorate2, money_tokens2 = infer_qty_and_prorate(after_end + " " + extra, abs_tol=0.02)
                    # keep first valid qty but prefer longer token list
                    if qty is None and qty2 is not None:
                        qty, prorate = qty2, prorate2
                    if len(money_tokens2) > len(money_tokens):
                        money_tokens = money_tokens2
                        add_debug(f"[EXTENDED] Extended tokens for sku={sku}: {len(money_tokens)} tokens")
            
            
            
                        # Replace the fallback quantity detection section around line 395-405:
            
            # To this (more permissive):
            if qty is None:
                # Strategy 1: Try simple first-line detection
                first_line = chunk_lines[0].strip()
                if first_line.isdigit() and 1 <= int(first_line) <= 100000:
                    qty = int(first_line)
                    add_debug(f"[FALLBACK QTY] sku={sku} using first line qty={qty}")
                else:
                    # Strategy 2: Look for decimal quantities (like 1.780)
                    for line in chunk_lines[:5]:  # Check first 5 lines
                        line = line.strip()
                        # Check for decimal numbers that could be quantities
                        if re.match(r'^\d+\.\d{3}$', line):  # Pattern like 1.780
                            decimal_qty = float(line)
                            if 1 <= decimal_qty <= 100000:
                                qty = int(decimal_qty)  # Convert 1.780 to 1780
                                add_debug(f"[DECIMAL QTY] sku={sku} converted {line} to {qty}")
                                break
                        # Check for comma-separated thousands (like 1,780)
                        elif re.match(r'^\d{1,3}(,\d{3})+$', line):  # Pattern like 1,780
                            comma_qty = int(line.replace(',', ''))
                            if 1 <= comma_qty <= 100000:
                                qty = comma_qty
                                add_debug(f"[COMMA QTY] sku={sku} converted {line} to {qty}")
                                break
            
            if qty is None or not (1 <= qty <= 100000):
                add_debug(f"[QTY INVALID] sku={sku} invalid qty={qty}")
                continue
            
            # ---- Bid Unit/Ext SVP (via money token positions) ----
            bid_unit_svp = None
            bid_ext_svp = None
            try:
                if len(money_tokens) >= 5:
                    # tokens: [EntUnit, EntExt, Disc%, BidUnit, BidExt, ...]
                    bid_unit_svp = parse_euro_number(money_tokens[3])
                    bid_ext_svp  = parse_euro_number(money_tokens[4])
                    add_debug(f"[SVP DIRECT] SKU '{sku}' - BidUnit={bid_unit_svp}, BidExt={bid_ext_svp}")
                else:
                    add_debug(f"[SVP TOKENS <5] sku={sku} tokens={money_tokens}")
                    # Fallback via Entitled + Disc%
                    if len(money_tokens) >= 3:
                        ent_unit = parse_euro_number(money_tokens[0])
                        ent_ext  = parse_euro_number(money_tokens[1])
                        disc_pct = parse_euro_number(money_tokens[2])  # e.g., 84,196 -> 84.196%
                        if ent_unit is not None and ent_ext is not None and disc_pct is not None:
                            factor = max(0.0, 1.0 - (disc_pct / 100.0))
                            bid_unit_svp = round(ent_unit * factor, 2)
                            bid_ext_svp  = round(ent_ext  * factor, 2)
                            add_debug(f"[SVP FALLBACK] sku={sku} -> bid=({bid_unit_svp},{bid_ext_svp})")
            except Exception as e:
                add_debug(f"[SVP ERROR] sku={sku} err={e}")
            
            # Convert to AED
            bid_unit_svp_aed = round(bid_unit_svp * USD_TO_AED, 2) if bid_unit_svp is not None else None
            bid_ext_svp_aed  = round(bid_ext_svp  * USD_TO_AED, 2) if bid_ext_svp  is not None else None
            
            # Final description cleanup
            desc = re.sub(r'\s{2,}', ' ', desc).strip()
            
            extracted_data.append([sku, desc, qty, start_date, end_date, bid_unit_svp_aed, bid_ext_svp_aed])
            i += window
            matched = True
            add_debug(f"[ROW EXTRACTED] Row {len(extracted_data)}: SKU='{sku}', Qty={qty}")
            break  # break window loop
        if not matched:
            i += 1
    
    add_debug(f"[EXTRACTION COMPLETE] Total rows extracted: {len(extracted_data)}")
    return extracted_data, header_info

# ----------------------------------------------------------------------
# Extract last page text (for "IBM Terms" sheet)
# ----------------------------------------------------------------------
def extract_last_page_text(file_like) -> str:
    doc = fitz.open(stream=file_like.read(), filetype="pdf")
    last_page = doc[-1]
    full_text = last_page.get_text("text") or last_page.get_text()
    
    # Filter to extract IBM terms content
    lines = full_text.splitlines()
    
    # First, collect the "Useful/Important web resources" section if it appears before IBM Terms
    useful_resources_section = []
    ibm_terms_section = []
    
    # Split into sections
    capture_useful = False
    capture_ibm_terms = False
    
    for line in lines:
        line = line.strip()
        
        # Start capturing useful resources section
        if "Useful/Important web resources:" in line:
            capture_useful = True
            useful_resources_section.append(line)
            continue
            
        # Start capturing IBM terms section
        if "IBM Terms and Conditions" in line:
            capture_ibm_terms = True
            capture_useful = False  # Stop capturing useful resources
            continue  # Skip the header itself
            
        # Skip company header info at the top
        if not capture_useful and not capture_ibm_terms:
            if any(skip_pattern in line for skip_pattern in ["IBM Ireland Product", "VAT: Reg", "Building", "Mulhuddart", "Dublin"]):
                continue
                
        # Capture useful resources content
        if capture_useful and line:
            useful_resources_section.append(line)
            
        # Capture IBM terms content  
        if capture_ibm_terms and line:
            # Stop at page numbers
            if line.lower().startswith("page ") and line.count(" ") <= 3:
                break
            ibm_terms_section.append(line)
    
    # Now reconstruct the IBM terms section into proper paragraphs
    reconstructed_ibm_terms = []
    current_paragraph = []
    
    for line in ibm_terms_section:
        line = line.strip()
        if not line:
            continue
            
        # Check if this line starts a new section/paragraph
        if (line.startswith("IBM International") or 
            line.startswith("The quote or order") or
            line.startswith("Unless specifically") or
            line.startswith("The terms of the IBM") or
            line.startswith("If you have any trouble")):
            
            # Save previous paragraph if exists
            if current_paragraph:
                reconstructed_ibm_terms.append(" ".join(current_paragraph))
                current_paragraph = []
            
            # Start new paragraph
            current_paragraph = [line]
        else:
            # Continue current paragraph
            if current_paragraph:
                current_paragraph.append(line)
            else:
                current_paragraph = [line]
    
    # Add the last paragraph
    if current_paragraph:
        reconstructed_ibm_terms.append(" ".join(current_paragraph))
    
    # Combine sections: IBM Terms first, then Useful Resources at the end
    all_content = reconstructed_ibm_terms
    if useful_resources_section:
        all_content.append("")  # Add spacing
        all_content.extend(useful_resources_section)
    
    result = "\n\n".join(all_content)  # Use double newlines for paragraph separation
    return result

# ----------------------------------------------------------------------
# Excel creation with enhanced debugging
# ----------------------------------------------------------------------
def create_styled_excel(
    data: list,
    header_info: dict,
    logo_path: str,
    output: BytesIO,
    compliance_text: str,
    ibm_terms_text: str
):
    """
    data rows: [SKU, Product Description, Quantity, Start Date, End Date, Unit Price AED, Total Price AED]
    """
    add_debug(f"[EXCEL START] Creating Excel with {len(data)} rows")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws.sheet_view.showGridLines = False
    
    # --- Header / Branding ---
    ws.merge_cells("B2:C3")
    if logo_path and os.path.exists(logo_path):
        img = Image(logo_path)
        img.width = 1.87 * 96  # 1.87 inches * 96 dpi
        img.height = 0.56 * 96
        ws.add_image(img, "B2")
    ws.merge_cells("D4:G4")
    ws["D4"] = "Quotation"
    ws["D4"].font = Font(size=20, color="1F497D")
    ws["D4"].alignment = Alignment(horizontal="center", vertical="center")
    
    # Balanced column widths - show full content but PDF-friendly
    ws.column_dimensions[get_column_letter(2)].width  = 8   # B (Sl) 
    ws.column_dimensions[get_column_letter(3)].width  = 15  # C (SKU)
    ws.column_dimensions[get_column_letter(4)].width  = 50  # D (Description) - balanced width
    ws.column_dimensions[get_column_letter(5)].width  = 10  # E (Qty)
    ws.column_dimensions[get_column_letter(6)].width  = 14  # F (Start Date)
    ws.column_dimensions[get_column_letter(7)].width  = 14  # G (End Date) 
    ws.column_dimensions[get_column_letter(8)].width  = 16  # H (Unit Price AED)
    ws.column_dimensions[get_column_letter(9)].width  = 16  # I (Total Price AED)
    
    # Left block
    left_labels = ["Date:", "From:", "Email:", "Contact:", "", "Company:", "Attn:", "Email:"]
    left_values = [
        datetime.today().strftime('%d/%m/%Y'),
        "Sneha Lokhandwala",
        "s.lokhandwala@mindware.net",
        "+971 55 456 6650",
        "",
        header_info.get('Reseller Name', 'empty'),
        "empty",
        "empty"
    ]
    row_positions = [6, 7, 8, 9, 10, 11, 12, 13]
    for row, label, value in zip(row_positions, left_labels, left_values):
        if label:
            ws[f"C{row}"] = label
            ws[f"C{row}"].font = Font(bold=True, color="1F497D")
        if value:
            ws[f"D{row}"] = value
            ws[f"D{row}"].font = Font(color="1F497D")
    
    # IBM Opp no.
    ws["C15"] = "IBM Opportunity Number: "
    ws["C15"].font = Font(bold=True, underline="single", color="000000")
    ws["D15"] = header_info.get('IBM Opportunity Number', '')
    ws["D15"].font = Font(bold=True, italic=True, underline="single", color="000000")
    
    # Right block
    right_labels = [
        "End User:", "Bid Number:", "Agreement Number:", "PA Site Number:", "",
        "Select Territory:", "Government Entity (GOE):", "Payment Terms:"
    ]
    right_values = [
        header_info.get('Customer Name', ''),
        header_info.get('Bid Number', ''),
        header_info.get('PA Agreement Number', ''),
        header_info.get('PA Site Number', ''),
        "",
        header_info.get('Select Territory', ''),
        header_info.get('Government Entity (GOE)', ''),
        "empty"
    ]
    for row, label, value in zip(row_positions, right_labels, right_values):
        ws.merge_cells(f"H{row}:I{row}")  # Only merge to column I to stay within table bounds
        ws[f"H{row}"] = f"{label} {value}"
        ws[f"H{row}"].font = Font(bold=True, color="1F497D")
        ws[f"H{row}"].alignment = Alignment(horizontal="left", vertical="center")
    
    # --- Table header (8 columns + serial) ---
    headers = ["Sl", "SKU", "Product Description", "Quantity", "Start Date", "End Date",
               "Unit Price in AED", "Total Price in AED"]
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    for col, header in enumerate(headers, start=2):
        ws.merge_cells(start_row=17, start_column=col, end_row=18, end_column=col)
        cell = ws.cell(row=17, column=col, value=header)
        cell.font = Font(bold=True, size=13, color="1F497D")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill
    
    # --- Data rows with enhanced debugging ---
    row_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    start_row = 19
    
    # Build a name->column index map from the header row for robust formatting
    headers_map = {ws.cell(row=17, column=c).value: c for c in range(2, 2 + len(headers))}
    col_unit = headers_map.get("Unit Price in AED")
    col_total = headers_map.get("Total Price in AED")
    
    for idx, row in enumerate(data, start=1):
        excel_row = start_row + idx - 1
        
        # Debug: Show what we're processing
        add_debug(f"[EXCEL WRITE] Processing row {idx}: SKU={row[0]}, Desc={row[1][:30]}...")
        
        # Serial number in column B (2)
        cell_sl = ws.cell(row=excel_row, column=2, value=idx)
        cell_sl.font = Font(size=11, color="1F497D")
        cell_sl.alignment = Alignment(horizontal="center", vertical="center")
        
        # Data values in columns C-I (3-9)
        for j, value in enumerate(row):
            excel_col = j + 3  # C=3, D=4, E=5, F=6, G=7, H=8, I=9
            cell = ws.cell(row=excel_row, column=excel_col, value=value)
            cell.font = Font(size=11, color="1F497D")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Debug specific problematic row
            if idx == 11:
                add_debug(f"[EXCEL ROW 11] Col {excel_col} ({get_column_letter(excel_col)}): Writing '{value}' (type: {type(value)})")
        
        # Row fill
        for col in range(2, 2 + len(headers)):
            ws.cell(row=excel_row, column=col).fill = row_fill
        
        # Description wrap & left align (column D = 4)
        ws.cell(row=excel_row, column=4).alignment = Alignment(wrap_text=True, horizontal="left", vertical="center")
        
        # Currency formats
        if col_unit:
            ws.cell(row=excel_row, column=col_unit).number_format = '"AED"#,##0.00'
        if col_total:
            ws.cell(row=excel_row, column=col_total).number_format = '"AED"#,##0.00'
    
    # --- Summary rows ---
    summary_row = start_row + len(data) + 2
    ws.merge_cells(f"C{summary_row}:G{summary_row}")
    ws[f"C{summary_row}"] = "TOTAL Bid Discounted Price"
    ws[f"C{summary_row}"].font = Font(bold=True, color="1F497D")
    ws[f"C{summary_row}"].alignment = Alignment(horizontal="right")
    
    # Sum Total Price (index 6 in data list)
    total_bid_aed = sum((row[6] or 0) for row in data if len(row) >= 7 and row[6] is not None)
    ws[f"H{summary_row}"] = total_bid_aed
    ws[f"H{summary_row}"].number_format = '"AED"#,##0.00'
    ws[f"H{summary_row}"].font = Font(bold=True, color="1F497D")
    ws[f"H{summary_row}"].fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    
    # --- Dynamic Terms block (main sheet) ---
    total_price_sum = total_bid_aed
    terms = get_terms_section(header_info, total_price_sum)
    
    def estimate_line_count(text, max_chars_per_line=80):
        lines = text.split('\n')
        total_lines = 0
        for line in lines:
            if not line:
                total_lines += 1
            else:
                wrapped = len(line) // max_chars_per_line + (1 if (len(line) % max_chars_per_line) else 0)
                total_lines += max(1, wrapped)
        return total_lines
    
    # Calculate the actual end row of the table content dynamically
    table_end_row = start_row + len(data) + 4  # data rows + summary + spacing
    terms_start_row = max(29, table_end_row + 2)  # Ensure terms start after table
    
    # Adjust terms positioning dynamically
    adjusted_terms = []
    row_offset = terms_start_row - 29  # Calculate offset from original row 29
    
    for cell_addr, text, *style in terms:
        try:
            # Extract row number and adjust it
            if len(cell_addr) >= 2 and cell_addr[1:].isdigit():
                col_letter = cell_addr[0]
                original_row = int(cell_addr[1:])
                new_row = original_row + row_offset
                new_cell_addr = f"{col_letter}{new_row}"
                adjusted_terms.append((new_cell_addr, text, *style))
            else:
                # Keep original if parsing fails
                adjusted_terms.append((cell_addr, text, *style))
        except Exception as e:
            adjusted_terms.append((cell_addr, text, *style))
    
    # Render the terms blocks
    for cell_addr, text, *style in adjusted_terms:
        try:
            if len(cell_addr) >= 2 and cell_addr[1:].isdigit():
                row_num = int(cell_addr[1:])
                col_letter = cell_addr[0]
                merge_rows = style[0].get("merge_rows") if style else None
                end_row = row_num + (merge_rows - 1 if merge_rows else 0)
                
                ws.merge_cells(f"{col_letter}{row_num}:H{end_row}")
                ws[cell_addr] = text
                ws[cell_addr].alignment = Alignment(wrap_text=True, vertical="top")
                # Height by estimated wrap - balanced for content visibility
                line_count = estimate_line_count(text, max_chars_per_line=80)
                total_height = max(18, line_count * 16)
                if merge_rows:
                   per_row = total_height / merge_rows
                   for r in range(row_num, end_row + 1):
                        ws.row_dimensions[r].height = per_row
                else:
                    ws.row_dimensions[row_num].height = total_height
                if style and "bold" in style[0]:
                    ws[cell_addr].font = Font(**style[0])
        except Exception as e:
            pass

    # Divider line across current header row
    border_row = 4
    bottom_border = Border(bottom=Side(style="thin", color="000000"))
    table_last_col = 9  # Only go to column I
    for col in range(1, table_last_col + 1):
        ws.cell(row=border_row, column=col).border = bottom_border
    
    # Calculate IBM Terms start position
    try:
        last_terms_row = max([int(addr[1:]) + (style[0].get("merge_rows", 1) - 1) 
                             for addr, text, *style in adjusted_terms 
                             if style and len(addr) >= 2 and addr[1:].isdigit()], 
                             default=terms_start_row + 10)
    except Exception:
        last_terms_row = terms_start_row + 10
        
    current_row = last_terms_row + 3
    
    # IBM Terms header
    ibm_header_cell = ws[f"C{current_row}"]
    ibm_header_cell.value = "IBM Terms and Conditions"
    ibm_header_cell.font = Font(bold=True, size=12, color="1F497D")
    current_row += 2
    
    # Add IBM Terms content
    lines = ibm_terms_text.splitlines()
    
    for i, line in enumerate(lines):
        if line.strip():
            line_text = line.strip()
            
            ws.merge_cells(f"C{current_row}:H{current_row}")
            cell = ws[f"C{current_row}"]
            
            # Check if line contains a URL
            url_pattern = r'https?://[^\s]+'
            import re
            urls = re.findall(url_pattern, line_text)
            
            if urls:
                for url in urls:
                    cell.hyperlink = url
                    cell.value = line_text
                    cell.font = Font(size=10, color="0563C1", underline="single")
            else:
                cell.value = line_text
                cell.font = Font(size=10, color="000000")
            
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            current_row += 1
            
            if "Useful/Important web resources" in line_text:
                current_row += 1
    
    # Page setup
    first_col = 2
    last_col = 9
    last_row = ws.max_row
    
    ws.print_area = f"B1:I{last_row}"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3
    
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.draft = False
    ws.page_setup.blackAndWhite = False
    ws.page_setup.scale = 85
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    
    add_debug(f"[EXCEL COMPLETE] Saved Excel with {len(data)} data rows")
    wb.save(output)

# Function to get debug info for Streamlit display
def get_extraction_debug():
    """Get debug info for display in Streamlit"""
    return get_debug_info()