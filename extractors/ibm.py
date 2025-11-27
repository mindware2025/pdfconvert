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
# Logging
# ----------------------------------------------------------------------
logging.basicConfig(
    filename="pdf_extraction_debug.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

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
# SKU map (for description correction)
# ----------------------------------------------------------------------
def _load_sku_map(csv_path=None, master_data=None):
    """
    Load SKU map from file or master_data DataFrame.
    Args:
        csv_path: Path to local CSV file (optional)
        master_data: pandas DataFrame with SKU mappings (optional)
    """
    sku_map = {}
    
    # First try to load from local file (existing logic)
    if csv_path is None:
        # Try multiple possible locations
        possible_paths = [
            "Quotation IBM PriceList csv.csv",
            os.path.join(os.path.dirname(__file__), "Quotation IBM PriceList csv.csv"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "Quotation IBM PriceList csv.csv"),
            "/mount/src/pdfconvert/Quotation IBM PriceList csv.csv",
            "./Quotation IBM PriceList csv.csv"
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                csv_path = path
                logging.info(f"[SKU MAP] Found CSV at: {path}")
                break
        
        if csv_path is None:
            logging.warning(f"[SKU MAP] CSV not found in any of these locations: {possible_paths}")
    
    # Load from local file if found
    if csv_path and os.path.exists(csv_path):
        try:
            sku_map_df = pd.read_csv(csv_path, usecols=["SKU", "SKU DESCRIPTION"])
            sku_map_df.dropna(subset=["SKU", "SKU DESCRIPTION"], inplace=True)
            sku_map = dict(zip(sku_map_df["SKU"], sku_map_df["SKU DESCRIPTION"]))
            logging.info(f"[SKU MAP] Loaded {len(sku_map)} mappings from local file")
        except Exception as e:
            logging.warning(f"[SKU MAP] Could not load local CSV: {e}")
    
    # Add/override with master_data if provided
    if master_data is not None:
        try:
            master_map = dict(zip(master_data["SKU"], master_data["SKU DESCRIPTION"]))
            sku_map.update(master_map)
            logging.info(f"[SKU MAP] Added {len(master_map)} mappings from master data")
        except Exception as e:
            logging.warning(f"[SKU MAP] Could not process master data: {e}")
    
    return sku_map

# Load initial SKU map from local file
SKU_MAP = _load_sku_map()

def correct_descriptions(extracted_data, master_data=None):
    """
    Each row: [sku, desc, qty, start_date, end_date, bid_unit_svp_aed, bid_ext_svp_aed]
    Replace desc if SKU_MAP has a match, or use master_data if provided.
    Args:
        extracted_data: List of extracted rows
        master_data: Optional pandas DataFrame with columns like ['SKU', 'SKU DESCRIPTION']
    """
    corrected = []
    
    # Create a combined lookup - prioritize master_data if available
    lookup_map = SKU_MAP.copy()  # Start with existing SKU_MAP
    
    if master_data is not None:
        try:
            # Convert master data to dictionary, assuming columns 'SKU' and 'SKU DESCRIPTION'
            master_map = dict(zip(master_data['SKU'], master_data['SKU DESCRIPTION']))
            lookup_map.update(master_map)  # Master data takes precedence
            logging.info(f"[MASTER DATA] Added {len(master_map)} SKU mappings from uploaded CSV")
        except Exception as e:
            logging.warning(f"[MASTER DATA ERROR] Could not process master data: {e}")
    
    for row in extracted_data:
        try:
            sku = row[0]
            if sku in lookup_map:
                original_desc = row[1]
                row[1] = lookup_map[sku]
                logging.info(f"[DESC CORRECTION] SKU {sku}: '{original_desc}' -> '{row[1]}'")
        except Exception as e:
            logging.warning(f"[DESC CORRECTION ERROR] Error processing row: {e}")
        corrected.append(row)
    
    return corrected

# ----------------------------------------------------------------------
# Robust regexes
# ----------------------------------------------------------------------
date_re = re.compile(r'\b\d{2}[‐‑–-][A-Za-z]{3}[‐‑–-]\d{4}\b')  # hyphen variants
sku_line_re = re.compile(r'^[A-Z0-9\-\._/]{5,20}$')
token_sku_re = re.compile(r'\b[A-Z0-9\-\._/]{5,20}\b')
int_re = re.compile(r'\b\d+\b')
# IMPORTANT: Money tokens must contain at least one separator (comma or dot)
# Prevents plain integers (e.g., '280') from being treated as money.
money_with_sep_re = re.compile(r'\d[\d.,]*[.,]\d+')
# Examples matched: 543,00 | 171.045,00 | 84,196 | 287,53 | 90.571,95
# NOT matched: 280

header_blacklist_re = re.compile(
    r'\b('
    r'Bid Number|Bid Request|Opportunity Number|Distributor|Distributor Name|Supplier|'
    r'Page \d+ of|IBM Ireland Product Distribution Limited|IBM Terms|Customer Information'
    r')\b', re.I
)

def looks_like_valid_sku(tok: str) -> bool:
    if not tok:
        return False
    if not sku_line_re.match(tok):
        return False
    # must contain at least one letter and one digit
    if not re.search(r'[A-Z]', tok):
        return False
    if not re.search(r'\d', tok):
        return False
    return True

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
def extract_ibm_data_from_pdf(file_like, master_data=None) -> tuple[list, dict]:
    """
    Extracts line items and header info from an IBM Quotation PDF.
    Args:
        file_like: PDF file stream
        master_data: Optional pandas DataFrame with master CSV data
    Returns:
      - extracted_data: list of rows
          [sku, desc, qty, start_date, end_date, bid_unit_svp_aed, bid_ext_svp_aed]
      - header_info: dict of customer/bid metadata
    """
    # Open PDF
    doc = fitz.open(stream=file_like.read(), filetype="pdf")
    # Collect lines
    lines = []
    for page in doc:
        page_text = page.get_text("text") or page.get_text()
        for l in page_text.splitlines():
            if l and l.strip():
                lines.append(l.rstrip())
    logging.info("---- RAW PDF LINES (first 80) ----")
    for idx, line in enumerate(lines[:80]):
        logging.info(f"{idx}: {repr(line)}")
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
            # Identify SKU and where description starts
            sku = None
            desc_start_index = None
            # Pattern: optional leading serial number
            if chunk_lines[0].strip().isdigit() and len(chunk_lines) > 1 and looks_like_valid_sku(chunk_lines[1].strip()):
                sku = chunk_lines[1].strip()
                desc_start_index = 2
            elif looks_like_valid_sku(chunk_lines[0].strip()):
                sku = chunk_lines[0].strip()
                desc_start_index = 1
            else:
                # Try to find a token that looks like a SKU
                t = token_sku_re.search(chunk)
                if t and looks_like_valid_sku(t.group(0)):
                    sku = t.group(0)
            if not sku:
                continue
            # Keep only chunks where a money token is "near" the start date (reduces false positives)
            pos_date = chunk.find(start_date)
            near_money = False
            if pos_date >= 0:
                window_text = chunk[max(0, pos_date - 60): pos_date + len(start_date) + 60]
                if money_with_sep_re.search(window_text):
                    near_money = True
            if not near_money:
                continue
            # Build description (text between SKU and first date)
            if desc_start_index is not None:
                desc_parts = []
                for ln in chunk_lines[desc_start_index:]:
                    if date_re.search(ln):
                        break
                    desc_parts.append(ln)
                desc = " ".join(desc_parts).strip()
            else:
                pos_sku = chunk.find(sku)
                pos_date0 = chunk.find(start_date)
                desc = chunk[pos_sku + len(sku):pos_date0].strip() if pos_sku >= 0 and pos_date0 > pos_sku else ""
            # ---- Robust Qty inference (ANY value) ----
            chunk_flat = " ".join(chunk_lines)
            all_date_matches = list(date_re.finditer(chunk_flat))
            if len(all_date_matches) < 2:
                continue
            end_date_match = all_date_matches[1]
            after_end = chunk_flat[end_date_match.end():].strip()
            # Optional debug: show strict tokens
            m_first_dbg = money_with_sep_re.search(after_end)
            if m_first_dbg:
                ints_zone_dbg = after_end[:m_first_dbg.start()]
                money_zone_dbg = after_end[m_first_dbg.start():]
                dbg_ints = [int(x) for x in int_re.findall(ints_zone_dbg)]
                dbg_amount = money_with_sep_re.findall(money_zone_dbg)
                logging.info(f"[TOKENS STRICT] INTs(pre)={dbg_ints}  MONEY(after)={dbg_amount[:6]}")
            # 1) First pass qty + tokens
            qty, prorate, money_tokens = infer_qty_and_prorate(after_end, abs_tol=0.02)
            logging.info(f"[QTY] sku={sku} start={start_date} end={end_date} -> qty={qty}, prorate={prorate}")
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
                        logging.info(f"[EXTENDED] Added tokens for sku={sku}: {money_tokens[:6]}")
            if qty is None or not (1 <= qty <= 100000):
                logging.info(f"[QTY ANOMALY] sku={sku} chunk={chunk_flat}")
                continue
            # ---- Bid Unit/Ext SVP (via money token positions) ----
            bid_unit_svp = None
            bid_ext_svp = None
            try:
                if len(money_tokens) >= 5:
                    # tokens: [EntUnit, EntExt, Disc%, BidUnit, BidExt, ...]
                    bid_unit_svp = parse_euro_number(money_tokens[3])
                    bid_ext_svp  = parse_euro_number(money_tokens[4])
                else:
                    logging.warning(f"[SVP TOKENS <5] sku={sku} tokens={money_tokens}")
                    # Fallback via Entitled + Disc%
                    if len(money_tokens) >= 3:
                        ent_unit = parse_euro_number(money_tokens[0])
                        ent_ext  = parse_euro_number(money_tokens[1])
                        disc_pct = parse_euro_number(money_tokens[2])  # e.g., 84,196 -> 84.196%
                        if ent_unit is not None and ent_ext is not None and disc_pct is not None:
                            factor = max(0.0, 1.0 - (disc_pct / 100.0))
                            bid_unit_svp = round(ent_unit * factor, 2)
                            bid_ext_svp  = round(ent_ext  * factor, 2)
                            logging.info(f"[SVP FALLBACK] sku={sku} ent=({ent_unit},{ent_ext}) disc={disc_pct}% -> bid=({bid_unit_svp},{bid_ext_svp})")
            except Exception as e:
                logging.warning(f"[SVP PARSE ERROR] sku={sku} err={e}")
            # Convert to AED
            bid_unit_svp_aed = round(bid_unit_svp * USD_TO_AED, 2) if bid_unit_svp is not None else None
            bid_ext_svp_aed  = round(bid_ext_svp  * USD_TO_AED, 2) if bid_ext_svp  is not None else None
            # Clean description
            desc = re.sub(r'\s{2,}', ' ', desc).strip()
            extracted_data.append([sku, desc, qty, start_date, end_date, bid_unit_svp_aed, bid_ext_svp_aed])
            i += window
            matched = True
            logging.info(f"Extracted row: {extracted_data[-1]}")
            break  # break window loop
        if not matched:
            i += 1
    return extracted_data, header_info

# ----------------------------------------------------------------------
# Extract last page text (for "IBM Terms" sheet)
# ----------------------------------------------------------------------
def extract_last_page_text(file_like) -> str:
    doc = fitz.open(stream=file_like.read(), filetype="pdf")
    last_page = doc[-1]
    full_text = last_page.get_text("text") or last_page.get_text()
    
    # Log the raw text for debugging
    logging.info("---- RAW LAST PAGE TEXT ----")
    for idx, line in enumerate(full_text.splitlines()):
        logging.info(f"{idx}: {repr(line)}")
    
    # Filter to extract IBM terms content
    lines = full_text.splitlines()
    filtered_lines = []
    
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
    logging.info("---- FILTERED IBM TERMS TEXT ----")
    logging.info(result)
    return result

# ----------------------------------------------------------------------
# Excel creation
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
    # --- Data rows ---
    row_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    start_row = 19
    # Build a name->column index map from the header row for robust formatting
    headers_map = {ws.cell(row=17, column=c).value: c for c in range(2, 2 + len(headers))}
    col_unit = headers_map.get("Unit Price in AED")
    col_total = headers_map.get("Total Price in AED")
    for idx, row in enumerate(data, start=1):
        excel_row = start_row + idx - 1
        # Serial
        cell_sl = ws.cell(row=excel_row, column=2, value=idx)
        cell_sl.font = Font(size=11, color="1F497D")
        cell_sl.alignment = Alignment(horizontal="center", vertical="center")
        # Data columns C..I (8 values)
        # Expected row = [SKU, Desc, Qty, Start, End, Unit AED, Total AED]
        for j, value in enumerate(row, start=3):
            cell = ws.cell(row=excel_row, column=j, value=value)
            cell.font = Font(size=11, color="1F497D")
            cell.alignment = Alignment(horizontal="center", vertical="center")
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
    ws.merge_cells("B29:C29")
    terms = get_terms_section(header_info, total_price_sum)
    def estimate_line_count(text, max_chars_per_line=80):  # Balanced for readability
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
    
    # Log the calculated positions for debugging
    logging.info(f"[TERMS POSITIONING] start_row={start_row}, data_rows={len(data)}, table_end={table_end_row}, terms_start={terms_start_row}")
    
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
                logging.info(f"[TERMS ADJUST] {cell_addr} -> {new_cell_addr} (offset={row_offset})")
                adjusted_terms.append((new_cell_addr, text, *style))
            else:
                # Keep original if parsing fails
                adjusted_terms.append((cell_addr, text, *style))
        except Exception as e:
            logging.error(f"[TERMS ADJUST ERROR] Failed to adjust {cell_addr}: {e}")
            adjusted_terms.append((cell_addr, text, *style))
    # Render the terms blocks
    for cell_addr, text, *style in adjusted_terms:
        try:
            if len(cell_addr) >= 2 and cell_addr[1:].isdigit():
                row_num = int(cell_addr[1:])
                col_letter = cell_addr[0]
                merge_rows = style[0].get("merge_rows") if style else None
                end_row = row_num + (merge_rows - 1 if merge_rows else 0)
                
                # Debug: Log what we're rendering
                logging.info(f"[TERMS RENDER] {cell_addr}, rows {row_num}-{end_row}, merge_rows={merge_rows}")
                
                ws.merge_cells(f"{col_letter}{row_num}:H{end_row}")
                ws[cell_addr] = text
                ws[cell_addr].alignment = Alignment(wrap_text=True, vertical="top")
                # Height by estimated wrap - balanced for content visibility
                line_count = estimate_line_count(text, max_chars_per_line=80)
                total_height = max(18, line_count * 16)  # Restored reasonable height
                if merge_rows:
                   per_row = total_height / merge_rows
                   for r in range(row_num, end_row + 1):
                        ws.row_dimensions[r].height = per_row  # Allow natural height
                else:
                    ws.row_dimensions[row_num].height = total_height  # Allow natural height
                if style and "bold" in style[0]:
                    ws[cell_addr].font = Font(**style[0])
        except Exception as e:
            logging.error(f"[TERMS RENDER ERROR] Failed to render {cell_addr}: {e}")

    # Divider line across current header row if needed - only for our table columns
    border_row = 4
    bottom_border = Border(bottom=Side(style="thin", color="000000"))
    table_last_col = 9  # Only go to column I (Total Price column)
    for col in range(1, table_last_col + 1):
        ws.cell(row=border_row, column=col).border = bottom_border
    
       # Calculate IBM Terms start position more safely
    try:
        last_terms_row = max([int(addr[1:]) + (style[0].get("merge_rows", 1) - 1) 
                             for addr, text, *style in adjusted_terms 
                             if style and len(addr) >= 2 and addr[1:].isdigit()], 
                             default=terms_start_row + 10)  # Fallback with safe spacing
    except Exception:
        last_terms_row = terms_start_row + 10  # Safe fallback
        
    current_row = last_terms_row + 3
    logging.info(f"[IBM TERMS START] Starting IBM Terms at row {current_row}")
    
    
    # IBM Terms header - blue like in the screenshot
    ibm_header_cell = ws[f"C{current_row}"]
    ibm_header_cell.value = "IBM Terms and Conditions"
    ibm_header_cell.font = Font(bold=True, size=12, color="1F497D")  # Blue header like screenshot
    current_row += 2
    
    # Add IBM Terms content with balanced formatting
    lines = ibm_terms_text.splitlines()
    
    for i, line in enumerate(lines):
        if line.strip():  # Only add non-empty lines
            line_text = line.strip()
            
            # Merge cells C to H for IBM terms (start from column C)
            ws.merge_cells(f"C{current_row}:H{current_row}")
            cell = ws[f"C{current_row}"]
            
            # Check if line contains a URL
            url_pattern = r'https?://[^\s]+'
            import re
            urls = re.findall(url_pattern, line_text)
            
            if urls:
                # If line contains URLs, make them clickable hyperlinks
                for url in urls:
                    cell.hyperlink = url
                    cell.value = line_text
                    cell.font = Font(size=10, color="0563C1", underline="single")  # Blue hyperlink color
            else:
                # Regular text in black
                cell.value = line_text
                cell.font = Font(size=10, color="000000")  # Standard readable size
            
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            current_row += 1
            
            # Only add spacing before "Useful/Important web resources" section
            if "Useful/Important web resources" in line_text:
                current_row += 1  # One extra space before this section only
    # Explicitly controlled page setup - prevent column 41 issues
    first_col = 2  # Start from B (exclude column A for better margins)
    last_col = 9   # I (Total Price column) - EXPLICITLY SET, not ws.max_column
    last_row = ws.max_row
    
    # Set print area with explicit boundaries to avoid column 41+ issues
    ws.print_area = f"B1:I{last_row}"  # Explicit range instead of calculated
    
    # Page orientation and scaling
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # Allow multiple pages vertically if needed
    
    # Balanced margins - not too tight, not too loose
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3
    
    # Additional PDF optimization settings
    ws.page_setup.paperSize = ws.PAPERSIZE_A4  # Ensure A4 paper size
    ws.page_setup.draft = False  # High quality output
    ws.page_setup.blackAndWhite = False  # Keep colors for professional look
    
    # Use intelligent scaling - start with 85% for better fit
    ws.page_setup.scale = 85  # Better balance between readability and fit
    
    # Set print options for better PDF output
    ws.sheet_properties.pageSetUpPr.fitToPage = False  # Use scale instead
    wb.save(output)