# extractors/ibm_template2.py
import fitz
import re
from datetime import datetime

# Use the same debug system as ibm.py
debug_info = []

def add_debug(message):
    """Add debug info"""
    debug_info.append(message)
    if len(debug_info) > 300:
        debug_info.pop(0)

def get_extraction_debug():
    """Get collected debug info"""
    return debug_info.copy()

def clear_debug():
    """Clear debug info"""
    debug_info.clear()

# Constants
USD_TO_AED = 3.6725

def parse_number(value: str):
    """Parse numbers with various formats"""
    try:
        if value is None:
            return None
        s = str(value).strip().replace(" ", "").replace(",", ".")
        return float(s)
    except Exception:
        return None

def extract_ibm_template2_from_pdf(file_like) -> tuple[list, dict]:
    """
    Extract data from IBM Template 2 (Software as a Service / Subscription format)
    Returns: (extracted_data, header_info)
    """
    clear_debug()
    
    doc = fitz.open(stream=file_like.read(), filetype="pdf")
    
    # Collect all text
    lines = []
    for page in doc:
        page_text = page.get_text("text") or page.get_text()
        for line in page_text.splitlines():
            if line and line.strip():
                lines.append(line.strip())
    
    add_debug(f"[TEMPLATE 2] Total lines extracted: {len(lines)}")
    
    # Header info extraction
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
    
    # Parse header fields
    for i, line in enumerate(lines):
        if "Customer Name:" in line:
            header_info["Customer Name"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "City:" in line:
            header_info["City"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "Country:" in line:
            header_info["Country"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "Bid Number:" in line:
            header_info["Bid Number"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "IBM Agreement Number:" in line or "PA Agreement Number:" in line:
            header_info["PA Agreement Number"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "IBM Site Number:" in line or "PA Site Number:" in line:
            header_info["PA Site Number"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "Select Territory:" in line:
            header_info["Select Territory"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "Government Entity" in line:
            header_info["Government Entity (GOE)"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "Reseller Name:" in line:
            header_info["Reseller Name"] = lines[i + 1] if i + 1 < len(lines) else ""
        elif "IBM Opportunity Number:" in line:
            # Extract the opportunity number from the same or next line
            opp_match = re.search(r'[A-Z0-9]{10,}', line)
            if opp_match:
                header_info["IBM Opportunity Number"] = opp_match.group()
            elif i + 1 < len(lines):
                opp_match = re.search(r'[A-Z0-9]{10,}', lines[i + 1])
                if opp_match:
                    header_info["IBM Opportunity Number"] = opp_match.group()
    
    add_debug(f"[TEMPLATE 2] Header info extracted: {header_info}")
    
    # Extract line items (Subscription Parts)
    extracted_data = []
    
    # Pattern for subscription part numbers (like D1009ZX, D100AZX)
    subscription_part_re = re.compile(r'\b[A-Z]\d{3,4}[A-Z]{2,3}\b')
    date_pattern = re.compile(r'\b\d{2}-[A-Za-z]{3}-\d{4}\b')
    
    # Look for "Software as a Service" sections
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Check if we're in a service section
        if 'Subscription Part#:' in line or 'Subscription Part:' in line:
            try:
                # Extract SKU from current or next line
                sku = None
                sku_match = subscription_part_re.search(line)
                if sku_match:
                    sku = sku_match.group()
                elif i + 1 < len(lines):
                    sku_match = subscription_part_re.search(lines[i + 1])
                    if sku_match:
                        sku = sku_match.group()
                
                if not sku:
                    i += 1
                    continue
                
                add_debug(f"[TEMPLATE 2] Found SKU: {sku}")
                
                # Extract service description (look backwards for service name)
                desc = ""
                for j in range(max(0, i - 10), i):
                    if 'IBM' in lines[j] and ('Service' in lines[j] or 'Integration' in lines[j] or 'Multicloud' in lines[j]):
                        desc = lines[j].strip()
                        break
                
                # Extract start date (Projected Service Start Date)
                start_date = ""
                for j in range(max(0, i - 5), min(i + 5, len(lines))):
                    if 'Start Date:' in lines[j]:
                        date_match = date_pattern.search(lines[j])
                        if date_match:
                            start_date = date_match.group()
                        elif j + 1 < len(lines):
                            date_match = date_pattern.search(lines[j + 1])
                            if date_match:
                                start_date = date_match.group()
                        break
                
                # Extract subscription length to calculate end date
                end_date = ""
                subscription_length = 12  # Default
                for j in range(i, min(i + 20, len(lines))):
                    if 'Subscription Length:' in lines[j]:
                        length_match = re.search(r'(\d+)\s*Months?', lines[j], re.I)
                        if length_match:
                            subscription_length = int(length_match.group(1))
                        break
                
                # Calculate end date if we have start date
                if start_date:
                    try:
                        from datetime import datetime
                        from dateutil.relativedelta import relativedelta
                        start_dt = datetime.strptime(start_date, '%d-%b-%Y')
                        end_dt = start_dt + relativedelta(months=subscription_length)
                        end_date = end_dt.strftime('%d-%b-%Y')
                    except:
                        end_date = ""
                
                # Extract quantity (look in table rows)
                qty = 1
                for j in range(i, min(i + 30, len(lines))):
                    # Look for quantity in table structure
                    if re.match(r'^\d{1,4}$', lines[j].strip()):
                        potential_qty = int(lines[j].strip())
                        if 1 <= potential_qty <= 10000:
                            qty = potential_qty
                            add_debug(f"[TEMPLATE 2] Found quantity: {qty}")
                            break
                
                # Extract pricing from table rows
                bid_unit_price = None
                bid_total_price = None
                
                # Look for "Bid Unit Price" and "Partner Bid Total Commit Value" columns
                for j in range(i, min(i + 35, len(lines))):
                    line_text = lines[j]
                    
                    # Look for lines with multiple decimal numbers (table data rows)
                    price_matches = re.findall(r'\b\d+[\.,]\d+\b', line_text)
                    
                    if len(price_matches) >= 2:
                        try:
                            # Usually: [...other values...] [unit_price] [total_price] USD
                            unit_str = price_matches[-2].replace(',', '.')
                            total_str = price_matches[-1].replace(',', '.')
                            
                            unit_val = float(unit_str)
                            total_val = float(total_str)
                            
                            # Validate that total ≈ unit * qty
                            if abs(total_val - (unit_val * qty)) < 1.0:
                                bid_unit_price = unit_val
                                bid_total_price = total_val
                                add_debug(f"[TEMPLATE 2] Found prices: Unit={bid_unit_price}, Total={bid_total_price}")
                                break
                        except:
                            continue
                
                # Convert USD to AED
                bid_unit_aed = round(bid_unit_price * USD_TO_AED, 2) if bid_unit_price else None
                bid_total_aed = round(bid_total_price * USD_TO_AED, 2) if bid_total_price else None
                
                # Add to extracted data
                # Format: [sku, desc, qty, start_date, end_date, bid_unit_aed, bid_total_aed]
                extracted_data.append([
                    sku,
                    desc,
                    qty,
                    start_date,
                    end_date,
                    bid_unit_aed,
                    bid_total_aed
                ])
                
                add_debug(f"[TEMPLATE 2] Extracted row: SKU={sku}, Qty={qty}, Unit={bid_unit_aed}, Total={bid_total_aed}")
                
            except Exception as e:
                add_debug(f"[TEMPLATE 2 ERROR] Error extracting line item: {e}")
        
        i += 1
    
    add_debug(f"[TEMPLATE 2 COMPLETE] Total rows extracted: {len(extracted_data)}")
    return extracted_data, header_info