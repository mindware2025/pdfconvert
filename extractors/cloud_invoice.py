import io
import re
import logging
from openpyxl import Workbook
import pandas as pd
from datetime import datetime
from dateutil import parser as _parser

# Configure logging to file
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('cloud_parsing_debug.log', mode='w'),  # Write to file
        logging.StreamHandler()  # Also keep console output
    ]
)
logger = logging.getLogger(__name__)

# === GUID Pattern Definitions ===
# GUID with whitespace and hyphen tolerance
GUID_WS = r'[0-9a-f]{8}[\s-]*[0-9a-f]{4}[\s-]*[0-9a-f]{4}[\s-]*[0-9a-f]{4}[\s-]*[0-9a-f]{12}'

# Updated manual token pattern for actual data format
# Looks for "Manual" followed by optional text, then space and GUID
manual_token_re = re.compile(rf'(?i)manual\w*\s+({GUID_WS})', re.IGNORECASE)

# === Header Definition ===
CLOUD_INVOICE_HEADER = [
    "Invoice No.", "Customer Code", "Customer Name", "Invoice Date", "Document Location",
    "Sale Location", "Delivery Location Code", "Delivery Date", "Annotation", "Currency Code",
    "Exchange Rate", "Shipment Mode", "Payment Term", "Mode Of Payment", "Status",
    "Credit Card Transaction No.", "HEADER Discount Code", "HEADER Discount %", "HEADER Currency", "HEADER Basis",
    "HEADER Disc Value", "HEADER Expense Code", "HEADER Expense %", "HEADER Expense Currency", "HEADER Expense Basis",
    "HEADER Expense Value", "Subscription Id", "Billing Cycle Start Date", "Billing Cycle End Date",
    "ITEM Code", "ITEM Name", "UOM", "Grade code-1", "Grade code-2", "Quantity", "Qty Loose",
    "Rate Per Qty", "Gross Value", "ITEM Discount Code", "ITEM Discount %", "ITEM Discount Currency", "ITEM Discount Basis",
    "ITEM Disc Value", "ITEM Expense Code", "ITEM Expense %", "ITEM Expense Currency", "ITEM Expense Basis",
    "ITEM Expense Value", "ITEM Tax Code", "ITEM Tax %", "ITEM Tax Currency", "ITEM Tax Basis", "ITEM Tax Value",
    "LPO Number", "End User", "Cost"
]

# === Mappings ===
exchange_rate_map = {
    "UJ000": 1,
    "TC000": 0.272294078,
    "QA000": 0.274725274725,
    "OM000": 2.60078023407,
    "KA000": 0.2666666666
}
tax_code_map = {
    "WT000": "", "QA000": "", "TC000": "SLVAT5",
    "OM000": "SLVAT5", "UJ000": "SEVAT0", "KA000": "SLVAT15"
}
tax_percent_map = {
    "WT000": "", "QA000": "", "TC000": 5,
    "OM000": 5, "UJ000": 0, "KA000": 15
}
currency_map = {
    "WT000": "", "QA000": "", "TC000": "AED",
    "OM000": "OMR", "UJ000": "USD", "KA000": "SAR"
}
keyword_map = {
    ("windows server", "window server","MSPER-CNS"): "MSPER-CNS",
    ( "azure subscription","MSAZ-CNS"): "MSAZ-CNS",
    ("google workspace","GL-WSP-CNS"): "GL-WSP-CNS",
    ("m365", "microsoft 365", "office 365", "exchange online","Microsoft Defender for Endpoint P1","MS-CNS"): "MS-CNS",
    ("acronis","AS-CNS"): "AS-CNS",
    ("windows 11 pro","MSPER-CNS"): "MSPER-CNS",
    ("power bi","MS-CNS"): "MS-CNS",
    ("planner", "project plan","MS-CNS"): "MS-CNS",
    ("power automate premium","MS-CNS"): "MS-CNS",
    ("visio","MS-CNS"): "MS-CNS",
    ("Microsoft Entra ID Governance (Education Faculty Pricing)","MS-CNS"): "MS-CNS",
    ("MSRI-CNS"): "MSRI-CNS",
    ("dynamics 365","MS-CNS"): "MS-CNS"
}

def debug_log(message):
    """Log debug messages to both file and console"""
    print(f"DEBUG: {message}")
    logger.debug(message)

def fmt_date(value):
    try:
        dt = _parser.parse(str(value), dayfirst=False, fuzzy=True)
        return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
    except Exception:
        return str(value) if value is not None else ""

def extract_digits(s: str) -> str:
    return "".join(ch for ch in str(s) if ch.isdigit())

def normalize_guid(guid_str: str) -> str:
    """
    Normalizes a GUID by removing spaces and ensuring proper format.
    Returns normalized GUID or original string if invalid.
    """
    if not guid_str:
        return ""
    
    # Remove all whitespace and convert to lowercase
    clean_guid = re.sub(r'\s+', '', guid_str.lower())
    
    # Check if it's a valid GUID format (32 hex chars with optional hyphens)
    guid_pattern = r'^[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}$'
    if re.match(guid_pattern, clean_guid):
        # Format as standard GUID with hyphens
        if '-' not in clean_guid:
            clean_guid = f"{clean_guid[:8]}-{clean_guid[8:12]}-{clean_guid[12:16]}-{clean_guid[16:20]}-{clean_guid[20:]}"
        return clean_guid
    
    return guid_str  # Return original if not valid GUID format

def build_cloud_invoice_df(df: pd.DataFrame) -> pd.DataFrame:
    today = datetime.today()
    today_str = today.strftime("%d/%m/%Y")
    out_rows = []
    for _, row in df.iterrows():
        cost = row.get("Cost", "")
        try: cost_val = float(cost)
        except: cost_val = 0
        out_row = {}
        doc_loc = row.get("DocumentLocation", "")
        gross_value = row.get("GrossValue", 0)
        out_row["Invoice No."] = row.get("InvoiceNo", "")
        out_row["Customer Code"] = row.get("CustomerCode", "")
        out_row["Customer Name"] = row.get("CustomerName", "")
        out_row["Invoice Date"] = today_str
        out_row["Document Location"] = doc_loc
        out_row["Sale Location"] = row.get("SaleLocation", "")
        out_row["Delivery Location Code"] = row.get("DeliveryLocationCode", "")
        out_row["Delivery Date"] = today_str
        out_row["Annotation"] = ""
        out_row["Currency Code"] = row.get("CurrencyCode", "")
        out_row["Exchange Rate"] = exchange_rate_map.get(doc_loc, "")
        out_row["Shipment Mode"] = row.get("ShipmentMode", "")
        out_row["Payment Term"] = row.get("PaymentTerm", "")
        out_row["Mode Of Payment"] = row.get("ModeOfPayment", "")
        out_row["Status"] = "Unpaid"
        out_row["Credit Card Transaction No."] = ""
        for field in CLOUD_INVOICE_HEADER[16:26]:
            out_row[field] = ""
        item_code = str(row.get("ITEMCode", "")).strip().lower()
        item_desc_raw = str(row.get("ITEMDescription", ""))
        item_desc_lower = item_desc_raw.lower()
        item_name_raw = str(row.get("ITEMName", "")).strip()
        invoice_desc = str(row.get("InvoiceDescription", "")).strip()
        sub_id_raw = row.get("SubscriptionId", "")
        sub_id = str(sub_id_raw).strip() if pd.notna(sub_id_raw) else ""
        
        invoice_desc_clean = re.sub(r"^[#\s]+", "", invoice_desc)
        sub_id_clean = sub_id[:36] if sub_id else "Sub"
        
        # === Manual Logic Implementation ===
        # Step 1: Check if "manual" exists in item name
        debug_log(f"Processing item_name_raw: '{item_name_raw}'")
        debug_log(f"Processing item_desc_raw: '{item_desc_raw}'")
        debug_log(f"Checking for 'manual' in item name (lowercase): {'manual' in item_name_raw.lower()}")
        
        if "manual" in item_name_raw.lower():
            debug_log("Manual detected in item name")
            # Step 2: Try to find GUID pattern in item description (directly, not after "manual")
            # Create simple GUID pattern since item description starts with GUID directly
            guid_pattern = re.compile(rf'({GUID_WS})', re.IGNORECASE)
            m = guid_pattern.search(item_desc_raw)
            debug_log(f"GUID regex search result in item description: {m}")
            
            if m:
                debug_log(f"GUID match found: {m.group(0)}")
                debug_log(f"GUID captured group: {m.group(1)}")
                
                # Step 3: If GUID found, normalize and use it
                found_guid = m.group(1)
                normalized_guid = normalize_guid(found_guid)
                debug_log(f"Original GUID: '{found_guid}' -> Normalized: '{normalized_guid}'")
                out_row["Subscription Id"] = normalized_guid
                
                # Parse additional info using semicolon, colon, hash, and tab delimiters
                parts = item_desc_raw.split(';')
                if len(parts) == 1:  # No semicolons, try colons
                    parts = item_desc_raw.split(':')
                if len(parts) == 1:  # No colons either, try hash symbols
                    parts = item_desc_raw.split('#')
                if len(parts) == 1:  # No hash symbols either, try tabs
                    parts = item_desc_raw.split('\t')
                debug_log(f"Split item description by delimiters: {parts}")
                
                # Look for item code in parts or use keyword mapping
                item_code_found = False
                if len(parts) >= 2:
                    for i, part in enumerate(parts):
                        part_clean = part.strip()
                        # Check if this part looks like an item code
                        if re.match(r'^[A-Z]{2,4}[-\s]*[A-Z]*$', part_clean) and len(part_clean) >= 2:
                            out_row["ITEM Code"] = part_clean
                            debug_log(f"Found item code in description part {i}: '{part_clean}'")
                            item_code_found = True
                            break
                
                # If no item code found in parts, use keyword mapping
                if not item_code_found:
                    for keywords, code in keyword_map.items():
                        for keyword in keywords:  # Iterate through each keyword in the tuple
                            if keyword in item_desc_lower:
                                out_row["ITEM Code"] = code
                                debug_log(f"Found item code using keyword mapping: '{keyword}' -> '{code}'")
                                item_code_found = True
                                break
                        if item_code_found:
                            break
                
                # Look for dates in the entire item description
                date_pattern = r'\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b'
                dates_found = re.findall(date_pattern, item_desc_raw)
                debug_log(f"Dates found in item description: {dates_found}")
                
                if len(dates_found) >= 2:
                    start_date = fmt_date(dates_found[0])
                    end_date = fmt_date(dates_found[1])
                    out_row["Billing Cycle Start Date"] = start_date
                    out_row["Billing Cycle End Date"] = end_date
                    debug_log(f"Set billing dates - Start: '{start_date}', End: '{end_date}'")
                elif len(dates_found) == 1:
                    start_date = fmt_date(dates_found[0])
                    out_row["Billing Cycle Start Date"] = start_date
                    debug_log(f"Set billing start date: '{start_date}'")
                
                # Parse LPO from the item description
                # Multiple LPO patterns to handle different formats
                lpo_patterns = [
                    r'(?:^|-)?\s*LPO\s*:\s*([A-Z0-9]+)',     # "LPO: XXXXX" or "- LPO: XXXXX"
                    r'(?:^|-)?\s*LPO\s*-?\s*([A-Z0-9]+)',    # "LPO- XXXXX" or "- LPO- XXXXX"
                    r'LPO\s*-\s*(PO\d+)',                    # "LPO- PO00155411"
                    r'PO\s*#\s*(\d+)',                       # Pattern like "PO # 158068"
                    r'\b(PO-\d+)\b',                         # Pattern like "PO-01076"
                    r'\b([A-Z]\d{8})\b',                     # Pattern like P00040411
                    r'\b(PO\d+)\b',                          # Pattern like PO00159398
                    r'\b(APO\d+)\b',                         # Pattern like APO2503065
                    r'\b(DPO\d+)\b',                         # Pattern like DPO2500101
                    r'\b(T\d{4}PO\d+)\b',                    # Pattern like T2025PO20240
                ]
                
                lpo_found = False
                for pattern in lpo_patterns:
                    lpo_match = re.search(pattern, item_desc_raw, re.IGNORECASE)
                    if lpo_match:
                        lpo_value = lpo_match.group(1).strip()
                        out_row["LPO Number"] = lpo_value[:30]
                        debug_log(f"Found and set LPO Number: '{lpo_value}' using pattern: '{pattern}'")
                        lpo_found = True
                        break
                
                if not lpo_found:
                    debug_log(f"No LPO pattern found in item description")
                
                # Parse End User from the item description - multiple patterns to handle different formats
                end_user_found = False
                
                # Pattern 1: "EU:Name" or "EU: Name" in semicolon/colon parts
                if len(parts) >= 3:
                    remaining_parts = parts[2:]
                    for part in remaining_parts:
                        part_clean = part.strip()
                        if 'eu:' in part_clean.lower() or 'eu -' in part_clean.lower():
                            # Extract after EU: or EU-
                            eu_match = re.search(r'eu\s*[:\-]\s*([^;:]+)', part_clean, re.IGNORECASE)
                            if eu_match:
                                end_user_raw = eu_match.group(1).strip()
                                # Clean up - remove extra spaces, semicolons, and encoding issues
                                end_user_clean = re.sub(r'\s*[;:]\s*\w+$', '', end_user_raw)
                                end_user_clean = re.sub(r'_x000D_', '', end_user_clean).strip()
                                if end_user_clean:
                                    out_row["End User"] = end_user_clean
                                    debug_log(f"Set End User from description parts: '{end_user_clean}'")
                                    end_user_found = True
                                    break
                
                # Pattern 2: "EU -Name" or "EU- Name" in the full description (for cases without semicolons/colons)
                if not end_user_found:
                    eu_patterns = [
                        r'EU\s*-\s*([^:\-\n]+?)(?:\s*$|\s*:|$)',      # "EU- Name" or "EU -Name" 
                        r'EU\s*:\s*([^:\-\n]+?)(?:\s*$|\s*:|$)',      # "EU: Name"
                        r':\s*EU\s*-\s*([^:\-\n]+?)(?:\s*$|\s*:|$)',  # ": EU - alkhalij"
                        r':\s*([A-Za-z][^:\n]*?)\s*$',                # ": Fidu Properties" (colon followed by name at end)
                    ]
                    
                    for pattern in eu_patterns:
                        eu_match = re.search(pattern, item_desc_raw, re.IGNORECASE)
                        if eu_match:
                            end_user_clean = eu_match.group(1).strip()
                            # Remove trailing dashes, encoding issues and extra content
                            end_user_clean = re.sub(r'\s*-+\s*$', '', end_user_clean)
                            end_user_clean = re.sub(r'_x000D_', '', end_user_clean).strip()
                            if end_user_clean:
                                out_row["End User"] = end_user_clean
                                debug_log(f"Set End User from full description: '{end_user_clean}' using pattern: '{pattern}'")
                                end_user_found = True
                                break
                
                if not end_user_found:
                    debug_log(f"No End User pattern found in item description")
                
                # Additional item code extraction if not found from parts
                if not item_code_found and len(parts) < 2:
                    debug_log(f"No delimiters found, looking for item code patterns in full description")
                    # Look for item code patterns like "MS-PLATINUM-CNS", "MSRI-CNS", etc.
                    item_code_patterns = [
                        r'\b(MS-[A-Z]+-CNS)\b',  # MS-PLATINUM-CNS, MS-BASIC-CNS, etc.
                        r'\b(MSRI-CNS)\b',       # MSRI-CNS
                        r'\b(AS-CNS)\b',         # AS-CNS
                        r'\b([A-Z]{2,4}-CNS)\b', # Any XX-CNS pattern
                    ]
                    
                    for pattern in item_code_patterns:
                        item_match = re.search(pattern, item_desc_raw, re.IGNORECASE)
                        if item_match:
                            item_code = item_match.group(1).upper()
                            out_row["ITEM Code"] = item_code
                            debug_log(f"Set ITEM Code from full description: '{item_code}' using pattern: '{pattern}'")
                            break
                
                debug_log(f"Final manual processing result - Sub ID: '{out_row.get('Subscription Id')}', Item Code: '{out_row.get('ITEM Code')}', LPO: '{out_row.get('LPO Number')}', End User: '{out_row.get('End User')}'")
            else:
                debug_log(f"No GUID found after manual, using default sub_id_clean: '{sub_id_clean}'")
                # Step 4: No GUID found after manual, use default but still try to extract LPO and End User
                out_row["Subscription Id"] = sub_id_clean
                
                # Still parse LPO and End User even without GUID
                # Parse additional info using semicolon, colon, hash, and tab delimiters
                parts = item_desc_raw.split(';')
                if len(parts) == 1:  # No semicolons, try colons
                    parts = item_desc_raw.split(':')
                if len(parts) == 1:  # No colons either, try hash symbols
                    parts = item_desc_raw.split('#')
                if len(parts) == 1:  # No hash symbols either, try tabs
                    parts = item_desc_raw.split('\t')
                debug_log(f"Split item description by delimiters (no GUID): {parts}")
                
                # Look for item code in parts or use keyword mapping
                item_code_found = False
                if len(parts) >= 2:
                    for i, part in enumerate(parts):
                        part_clean = part.strip()
                        # Check if this part looks like an item code
                        if re.match(r'^[A-Z]{2,4}[-\s]*[A-Z]*$', part_clean) and len(part_clean) >= 2:
                            out_row["ITEM Code"] = part_clean
                            debug_log(f"Found item code in description part {i} (no GUID): '{part_clean}'")
                            item_code_found = True
                            break
                
                # If no item code found in parts, use keyword mapping
                if not item_code_found:
                    for keywords, code in keyword_map.items():
                        for keyword in keywords:  # Iterate through each keyword in the tuple
                            if keyword in item_desc_lower:
                                out_row["ITEM Code"] = code
                                debug_log(f"Found item code using keyword mapping (no GUID): '{keyword}' -> '{code}'")
                                item_code_found = True
                                break
                        if item_code_found:
                            break
                
                # Parse LPO from the item description
                lpo_patterns = [
                    r'(?:^|-)?\s*LPO\s*:\s*([A-Z0-9]+)',     # "LPO: XXXXX" or "- LPO: XXXXX"
                    r'(?:^|-)?\s*LPO\s*-?\s*([A-Z0-9]+)',    # "LPO- XXXXX" or "- LPO- XXXXX"
                    r'LPO\s*-\s*(PO\d+)',                    # "LPO- PO00155411"
                    r'PO\s*#\s*(\d+)',                       # Pattern like "PO # 158068"
                    r'\b(PO-\d+)\b',                         # Pattern like "PO-01076"
                    r'\b([A-Z]\d{8})\b',                     # Pattern like P00040411
                    r'\b(PO\d+)\b',                          # Pattern like PO00159398
                    r'\b(APO\d+)\b',                         # Pattern like APO2503065
                    r'\b(DPO\d+)\b',                         # Pattern like DPO2500101
                    r'\b(T\d{4}PO\d+)\b',                    # Pattern like T2025PO20240
                ]
                
                lpo_found = False
                for pattern in lpo_patterns:
                    lpo_match = re.search(pattern, item_desc_raw, re.IGNORECASE)
                    if lpo_match:
                        lpo_value = lpo_match.group(1).strip()
                        out_row["LPO Number"] = lpo_value[:30]
                        debug_log(f"Found and set LPO Number (no GUID): '{lpo_value}' using pattern: '{pattern}'")
                        lpo_found = True
                        break
                
                if not lpo_found:
                    debug_log(f"No LPO pattern found in item description (no GUID)")
                
                # Parse End User from the item description
                end_user_found = False
                
                # Pattern 1: "EU:Name" or "EU: Name" in semicolon/colon parts
                if len(parts) >= 3:
                    remaining_parts = parts[2:]
                    for part in remaining_parts:
                        part_clean = part.strip()
                        if 'eu:' in part_clean.lower() or 'eu -' in part_clean.lower():
                            # Extract after EU: or EU-
                            eu_match = re.search(r'eu\s*[:\-]\s*([^;:]+)', part_clean, re.IGNORECASE)
                            if eu_match:
                                end_user_raw = eu_match.group(1).strip()
                                # Clean up - remove extra spaces, semicolons, and encoding issues
                                end_user_clean = re.sub(r'\s*[;:]\s*\w+$', '', end_user_raw)
                                end_user_clean = re.sub(r'_x000D_', '', end_user_clean).strip()
                                if end_user_clean:
                                    out_row["End User"] = end_user_clean
                                    debug_log(f"Set End User from description parts (no GUID): '{end_user_clean}'")
                                    end_user_found = True
                                    break
                
                # Pattern 2: "EU -Name" or "EU- Name" in the full description
                if not end_user_found:
                    eu_patterns = [
                        r'EU\s*-\s*([^:\-\n]+?)(?:\s*$|\s*:|$)',      # "EU- Name" or "EU -Name" 
                        r'EU\s*:\s*([^:\-\n]+?)(?:\s*$|\s*:|$)',      # "EU: Name"
                        r':\s*EU\s*-\s*([^:\-\n]+?)(?:\s*$|\s*:|$)',  # ": EU - alkhalij"
                        r':\s*([A-Za-z][^:\n]*?)\s*$',                # ": Fidu Properties" (colon followed by name at end)
                    ]
                    
                    for pattern in eu_patterns:
                        eu_match = re.search(pattern, item_desc_raw, re.IGNORECASE)
                        if eu_match:
                            end_user_clean = eu_match.group(1).strip()
                            # Remove trailing dashes, encoding issues and extra content
                            end_user_clean = re.sub(r'\s*-+\s*$', '', end_user_clean)
                            end_user_clean = re.sub(r'_x000D_', '', end_user_clean).strip()
                            if end_user_clean:
                                out_row["End User"] = end_user_clean
                                debug_log(f"Set End User from full description (no GUID): '{end_user_clean}' using pattern: '{pattern}'")
                                end_user_found = True
                                break
                
                if not end_user_found:
                    debug_log(f"No End User pattern found in item description (no GUID)")
                
                debug_log(f"Final manual processing result (no GUID) - Sub ID: '{out_row.get('Subscription Id')}', Item Code: '{out_row.get('ITEM Code')}', LPO: '{out_row.get('LPO Number')}', End User: '{out_row.get('End User')}'")
        
        
        # === Existing Logic ===
        elif item_code == "az-cns":
            digits = extract_digits(invoice_desc_clean)
            out_row["Subscription Id"] = digits[-36:] if digits else sub_id_clean
        elif item_code == "msri-cns":
            out_row["Subscription Id"] = invoice_desc_clean[:36] if invoice_desc_clean else sub_id_clean
        elif "reserved vm instance" in item_desc_lower:
            out_row["Subscription Id"] = item_desc_raw[:38] if item_desc_raw else sub_id_clean
        else:
            # Step 5: Final fallback - Use default subscription ID
            out_row["Subscription Id"] = sub_id_clean
        
        # ✅ ONLY set input dates if manual processing didn't already extract them
        # Don't overwrite dates that manual processing successfully extracted
        if "Billing Cycle Start Date" not in out_row or out_row["Billing Cycle Start Date"] == "":
            input_start = fmt_date(row.get("BillingCycleStartDate", ""))
            # Only use if input actually has a valid date (not empty/nan)
            if input_start and str(input_start).strip() not in ["", "nan", "None"]:
                out_row["Billing Cycle Start Date"] = input_start
            elif "Billing Cycle Start Date" not in out_row:
                out_row["Billing Cycle Start Date"] = ""
        
        if "Billing Cycle End Date" not in out_row or out_row["Billing Cycle End Date"] == "":
            input_end = fmt_date(row.get("BillingCycleEndDate", ""))
            # Only use if input actually has a valid date (not empty/nan)
            if input_end and str(input_end).strip() not in ["", "nan", "None"]:
                out_row["Billing Cycle End Date"] = input_end
            elif "Billing Cycle End Date" not in out_row:
                out_row["Billing Cycle End Date"] = ""
        
        item_code = row.get("ITEMCode", "")
        if pd.notna(item_code) and str(item_code).strip():
            out_row["ITEM Code"] = item_code
        else:
            matched = False
            for keywords, code in keyword_map.items():
                for k in keywords:
                    if k in item_desc_lower:
                        out_row["ITEM Code"] = code
                        matched = True
                        break
                if matched:
                    break
            if not matched:
                out_row["ITEM Code"] = ""
        
        # ITEM Name merged description
        # Clean and prepare values
        item_code_upper = out_row.get("ITEM Code", "").strip().upper()
        item_desc_raw = str(row.get("ITEMDescription", "")).strip()
        item_name_raw = str(row.get("ITEMName", "")).strip()
        invoice_desc_clean = re.sub(r"^[#\s]+", "", str(row.get("InvoiceDescription", "")).strip())
        sub_id_raw = row.get("SubscriptionId", "")
        sub_id_full = str(sub_id_raw).strip() if pd.notna(sub_id_raw) else ""
        
        # Choose item_name_detail based on ITEM Code
        if item_code_upper == "MSRI-CNS":
            item_name_detail = invoice_desc_clean
        elif item_code_upper in ["MSAZ-CNS", "AS-CNS", "AWS-UTILITIES-CNS"]:
            item_name_detail = sub_id_full  # full SubscriptionId
        else:
            item_name_detail = out_row.get("Subscription Id", "").strip()
        
        # Billing cycle dates
        billing_start = out_row.get("Billing Cycle Start Date", "").strip()
        billing_end = out_row.get("Billing Cycle End Date", "").strip()
        
        # Format billing_end as MM/YYYY if applicable
        billing_info = ""
        if billing_end and billing_end.lower() != "nan":
            try:
                billing_end_dt = pd.to_datetime(billing_end, dayfirst=True)
                billing_info = billing_end_dt.strftime("%m/%Y")
            except Exception:
                billing_info = billing_end  # fallback
        
        ## Build parts list and skip empty or NaN values
        ##parts = [
         #   item_desc_raw,
         #   item_name_raw,
         #   item_name_detail,
        #]
        #    parts.append(f"{billing_start}-{billing_end}")
       # 
       # 
       # # Join non-empty parts with hyphen
       # out_row["ITEM Name"] = "-".join([p for p in parts if p and p.lower() != "nan"])
        # Special rule for KA000: use only column AE (ITEMName)
        if doc_loc == "KA000":
            out_row["ITEM Name"] = item_desc_raw.strip()
        else:
            # Build parts list and skip empty or NaN values
            parts = [
                item_desc_raw,
                item_name_raw,
                item_name_detail,
            ]
        
            
            if item_code_upper == "MSRI-CNS" and billing_info:
                parts.append(billing_info)
            elif item_code_upper in ["MSAZ-CNS", "AS-CNS", "AWS-UTILITIES-CNS"] and billing_info:
                parts.append(billing_info)
            elif billing_start and billing_end and billing_start.lower() != "nan" and billing_end.lower() != "nan":
                parts.append(f"{billing_start}-{billing_end}")
            
                # Join non-empty parts with hyphen
            out_row["ITEM Name"] = "-".join([p for p in parts if p and p.lower() != "nan"])

            # Join non-empty parts with hyphen
        
        #out_row["ITEM Name"] = f"{item_desc_raw}-{row.get('ITEMName','')}-{out_row['Subscription Id']}#{out_row['Billing Cycle Start Date']}-{out_row['Billing Cycle End Date']}"
        out_row["UOM"] = row.get("UOM", "")
        out_row["Grade code-1"] = "NA"
        out_row["Grade code-2"] = "NA"
        out_row["Quantity"] = row.get("Quantity", "")
        out_row["Qty Loose"] = row.get("QtyLoose", "")
        try:
            out_row["Rate Per Qty"] = float(gross_value) / float(row.get("Quantity", 1))
        except: out_row["Rate Per Qty"] = 0
        try:
            out_row["Gross Value"] = round(float(gross_value), 2)
        except: out_row["Gross Value"] = 0.00
        for field in CLOUD_INVOICE_HEADER[38:48]:
            out_row[field] = ""
        out_row["ITEM Tax Code"] = tax_code_map.get(doc_loc, "")
        out_row["ITEM Tax %"] = tax_percent_map.get(doc_loc, "")
        out_row["ITEM Tax Currency"] = currency_map.get(doc_loc, "")
        out_row["ITEM Tax Basis"] = "R" if doc_loc not in ["WT000", "QA000"] else ""
        try: gross_val = float(gross_value)
        except: gross_val = 0
        tax_value = {
            "TC000": round(gross_val * 0.05, 2),
            "OM000": round(gross_val * 0.05, 2),
            "KA000": round(gross_val * 0.15, 2),
            "UJ000": 0
        }.get(doc_loc, "")
        out_row["ITEM Tax Value"] = tax_value
        lpo = row.get("LPONumber", "")
        # Only set LPO from original data if manual parsing didn't already set it
        if not out_row.get("LPO Number"):
            out_row["LPO Number"] = "" if pd.isna(lpo) or str(lpo).strip().lower() in ["nan", "none"] else str(lpo)[:30]
        #end_user = str(row.get("EndUser", ""))
        #end_user_country = str(row.get("EndUserCountryCode", ""))
        
        #if end_user.strip().lower() in ["", "nan", "none"] or end_user_country.strip().lower() in ["", "nan", "none"]:
       #      out_row["End User"] = ""
        #else:
        #     out_row["End User"] = f"{end_user} ; {end_user_country}"
        # Normalize and clean inputs
        end_user = str(row.get("EndUser", "")).strip()
        end_user_country = str(row.get("EndUserCountryCode", "")).strip()
        
        # Define what values are considered invalid
        invalid_values = ["", "nan", "none"]
        
        # Only set End User from original data if manual parsing didn't already set it
        if not out_row.get("End User"):
            # Check for invalid End User or Country Code
            if end_user.lower() in invalid_values or end_user_country.lower() in invalid_values:
                out_row["End User"] = ""  # Show empty string if missing
                out_row["_highlight_end_user"] = True  # Flag for red highlight
            else:
                out_row["End User"] = f"{end_user} ; {end_user_country}"
                out_row["_highlight_end_user"] = False
        

        try: out_row["Cost"] = round(float(cost_val), 2)
        except: out_row["Cost"] = cost
        out_rows.append(out_row)

    result_df = pd.DataFrame(out_rows, columns=CLOUD_INVOICE_HEADER)

    # AS-CNS aggregation
    try:
        if not result_df.empty:
            is_as_cns = result_df["ITEM Code"].astype(str).str.strip().str.upper() == "AS-CNS"
            df_as = result_df[is_as_cns].copy()
            df_other = result_df[~is_as_cns].copy()
            if not df_as.empty:
                df_as["Gross Value"] = pd.to_numeric(df_as["Gross Value"], errors="coerce").fillna(0.0)
                df_as["Cost"] = pd.to_numeric(df_as["Cost"], errors="coerce").fillna(0.0)
                group_cols = ["Invoice No.", "End User", "LPO Number"]
                agg = df_as.groupby(group_cols, as_index=False).agg({"Gross Value": "sum"})
                merged = agg.merge(
                    df_as.drop(columns=["Gross Value"]).drop_duplicates(subset=group_cols),
                    on=group_cols,
                    how="left"
                )
                merged["Gross Value"] = merged["Gross Value"].round(2)
                merged["Quantity"] = 1
                merged["Rate Per Qty"] = merged["Gross Value"]
                merged["Cost"] = merged["Gross Value"]
    
                # <-- FIX: recalculate ITEM Tax Value based on summed Gross Value -->
                tax_rate_map = {"TC000": 0.05, "OM000": 0.05, "KA000": 0.15, "UJ000": 0}
                merged["ITEM Tax Value"] = merged.apply(
                    lambda x: round(x["Gross Value"] * tax_rate_map.get(x["Document Location"], 0), 2),
                    axis=1
                )
    
                result_df = pd.concat([df_other, merged], ignore_index=True)[CLOUD_INVOICE_HEADER]
    except Exception:
        pass

    return result_df

# === Versioning / Invoice Number Mapping ===

def create_summary_sheet(processed_df: pd.DataFrame) -> pd.DataFrame:
    sorted_df = processed_df.sort_values(by=processed_df.columns.tolist()).reset_index(drop=True)
    summary_df = sorted_df[["Invoice No.", "LPO Number", "End User"]].copy()
    summary_df["Combined (D)"] = (
        summary_df["Invoice No."].astype(str) +
        summary_df["LPO Number"].astype(str) +
        summary_df["End User"].astype(str)
    )
    summary_df = summary_df.drop_duplicates(subset=["Combined (D)"]).reset_index(drop=True)
    summary_df["V1"] = summary_df["Invoice No."].ne(summary_df["Invoice No."].shift()).astype(int).replace(0, "")
    v2 = []
    for i, v1 in enumerate(summary_df["V1"]):
        if i == 0:
            v2.append(1 if v1 == 1 else "")
        else:
            if v1 == 1:
                v2.append(1)
            else:
                prev_v2 = v2[-1]
                v2.append(prev_v2 + 1 if isinstance(prev_v2, int) else "")
    summary_df["V2"] = v2
    summary_df["V3"] = "-" + summary_df["V1"].astype(str) + summary_df["V2"].astype(str)
    summary_df["V4"] = summary_df["Invoice No."].astype(str) + summary_df["V3"]
    return summary_df

def map_invoice_numbers(processed_df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns processed_df with a NEW column 'Updated Invoice No.' instead of replacing the original.
    """
    summary_df = create_summary_sheet(processed_df)
    mapping = dict(zip(summary_df["Combined (D)"], summary_df["V4"]))
    processed_df["Combined (D)"] = (
        processed_df["Invoice No."].astype(str) +
        processed_df["LPO Number"].astype(str) +
        processed_df["End User"].astype(str)
    )
    processed_df.drop(columns=["Combined (D)"], inplace=True)
    return processed_df

def create_srcl_file(df):
   

    # --- Sheet 1 headers ---
    headers_head = [
        "S.No",
        "Date - (dd/MM/yyyy)",
        "Cust_Code",
        "Curr_Code",
        "FORM_CODE",
        "Doc_Src_Locn",
        "Location_Code",
        "Delivery_Location",
        "SalesmanID"
    ]

    # --- Sheet 2 headers ---
    headers_item = [
        "S.No",
        "Ref. Key",
        "Item_Code",
        "Item_Name",
        "Grade1",
        "Grade2",
        "UOM",
        "Qty",
        "Qty_Ls",
        "Rate",              # Unit rate
        "CI Number CL",
        "End User CL",
        "Subs ID CL",
        "MPC Billdate CL",   # Dynamic based on Document Location
        "Unit Cost CL",
        "Total"              # Moved to end
    ]

    wb = Workbook()

    # --- Sheet 1 ---
    ws_head = wb.active
    ws_head.title = "SALES_RET_HEAD"
    ws_head.append(headers_head)

    today_str = pd.Timestamp.today().strftime("%d/%m/%Y")

    # Sequential numeric S.No for header
    s_no = 1
    header_sno_map = {}
    for _, row in df.iterrows():
        versioned_inv = row.get("Versioned Invoice No.", row.get("Invoice No.", ""))
        if versioned_inv not in header_sno_map:
            header_sno_map[versioned_inv] = s_no
            ws_head.append([
                s_no,
                today_str,
                row.get("Customer Code", ""),
                row.get("Currency Code", ""),
                "0",  # FORM_CODE
                row.get("Document Location", ""),
                row.get("Document Location", ""),
                row.get("Delivery Location Code", ""),
                "ED068"
            ])
            s_no += 1

    # --- Sheet 2 ---
    ws_item = wb.create_sheet(title="SALES_RET_ITEM")
    ws_item.append(headers_item)

    item_counter = 1
    for _, row in df.iterrows():
        versioned_inv = row.get("Versioned Invoice No.", row.get("Invoice No.", ""))
        ref_key = header_sno_map.get(versioned_inv, "")
        doc_loc = str(row.get("Document Location", "")).strip().upper()

        # --- Clean and sanitize item name ---
        raw_item_name = str(row.get("ITEM Name", "")).strip()
        clean_item_name = re.sub(r"[\r\n]+", " ", raw_item_name)  # remove newlines (Ctrl+J)
        clean_item_name = clean_item_name.replace("'", "").replace('"', "")
        clean_item_name = re.sub(r"\s+", " ", clean_item_name).strip()[:240]

        # --- Numeric cleanup ---
        qty = abs(float(row.get("Quantity", 0) or 0))
        qty_ls = abs(float(row.get("Qty Loose", 0) or 0))
        rate = abs(float(row.get("Rate Per Qty", 0) or 0))
        total = abs(round(qty * rate, 2))
        unit_cost = abs(float(row.get("Cost", 0) or 0))

        # --- MPC Billdate CL mapping ---
        if doc_loc in ["TC000", "UJ000"]:
            mpc_billdate = "UAE - 28"
        elif doc_loc == "QA000":
            mpc_billdate = "QAR - 28"
        elif doc_loc == "WT000":
            mpc_billdate = "KWT - 28"
        else:
            mpc_billdate = ""

        # --- Append cleaned data ---
        ws_item.append([
            item_counter,
            ref_key,
            row.get("ITEM Code", ""),
            clean_item_name,
            row.get("Grade code-1", ""),
            row.get("Grade code-2", ""),
            row.get("UOM", ""),
            qty,
            qty_ls,
            rate,  # ✅ Unit rate only
            versioned_inv,
            row.get("End User", ""),
            row.get("Subscription Id", ""),
            mpc_billdate,  # ✅ new logic
            unit_cost,
            total  # moved to end
        ])

        item_counter += 1

    # --- Save to memory ---
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
