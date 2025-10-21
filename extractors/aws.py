# aws.py

from collections import defaultdict
import fitz  # PyMuPDF
import re
from datetime import datetime

AWS_OUTPUT_COLUMNS = [
    "DATE", "INV NUMBER", "WITHOUT VAT", "WITH VAT", "NARRATION",
    "Account Period", "A/C", "Due date", "Vat USD", "Vat AED",
    "Total USD", "Inv value", "Check", "Bill to"
]

def extract_due_date_fallback(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    normalized_text = re.sub(r"\s+", " ", full_text)
    match = re.search(r"TOTAL AMOUNT DUE ON\s+([A-Za-z]+)\s+(\d{1,2})\s*,?\s+(\d{4})", normalized_text)
    if match:
        month, day, year = match.groups()
        try:
            return datetime.strptime(f"{month} {day}, {year}", "%B %d, %Y").strftime("%d/%m/%Y")
        except ValueError:
            return ""
    return ""

def extract_value(pattern, text, default=""):
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else default

def extract_common_fields(text, is_credit_note=False, template="Unknown"):
  
    if template == "C":
        invoice_number = extract_value(r"Invoice Number:\s*([0-9]+)", text)
    else:
        invoice_number = extract_value(r"(EU[A-Z]+[0-9]{2}-\d+)", text)

    
    match = re.search(
    r"(This\s+(?:Tax\s+Invoice|Tax\s+Credit\s+Note|Document|invoice)?\s*is\s+for\s+the\s+billing\s+period\s+[A-Za-z]+\s+\d{1,2}\s*[-–]\s*[A-Za-z]+\s+\d{1,2}\s*,?\s*\d{4})",
    text,
    re.IGNORECASE
)
    formatted_period = match.group(1).strip() if match else ""
    
    net_charges_usd = ""
    if is_credit_note or template == "A"   :
       
        
        match = re.search(
            r"-?\s*USD\s*(-?[0-9,]+\.[0-9]{2})\s*-?\s*AED\s*-?[0-9,]+\.[0-9]{2}\s*Net Charges",
            text,
            re.IGNORECASE
)

        if match:
            net_charges_usd = match.group(1).replace(",", "")
    elif template in ["C", "D"]:
    
           normalized_text = re.sub(r"\s+", " ", text)
           match = re.search(
               r"TOTAL AMOUNT.*?(?:USD|\$)?\s*([0-9,]+\.[0-9]{2})",
               normalized_text,
               re.IGNORECASE
    )

           if match:
                net_charges_usd = match.group(1).replace(",", "")

    account_number = extract_value(r"(?:Account Number|رقم الحساب)[^\d]*?(\d{9,})", text)
    return invoice_number, net_charges_usd, account_number, formatted_period
def extract_bill_to(text, template):
    if template == "C":
        match = re.search(r"Bill to Address:\s*(.*?)\s+ATTN:", text, re.IGNORECASE)
        if match:
            bill_to_line = match.group(1).strip().upper()
            if "MINDWARE TECHNOLOGY" in bill_to_line:
                return "MINDWARE TECHNOLOGY TRADING L.L.C"
            elif "MINDWARE FZ" in bill_to_line:
                return "Mindware FZ LLC"
            return match.group(1).strip()
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "Address" in line or "العنوان" in line:
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                upper_line = next_line.upper()
                if upper_line.startswith("MINDWARE TECHNOLOGY"):
                    return "MINDWARE TECHNOLOGY TRADING L.L.C"
                elif "MINDWARE FZ" in upper_line:
                    return "Mindware FZ LLC"
                return next_line
    return ""

def detect_template(text):
    if "Tax Credit Note" in text:
        return "B"
    elif "Tax Invoice" in text:
        return "A"
    elif "Amazon Web Services, Inc. Invoice" in text and "Invoice Number:" in text:
        return "C"
    elif "AWS Marketplace Invoice" in text or "Marketplace Operator Invoicing" in text:
        return "D"
    return "Unknown"

def extract_supp_ref_date(text, template):
    if template == "A":
        date_str = extract_value(r"Tax Invoice Date:.*?([A-Za-z]+\s+\d{1,2},?\s+\d{4})", text)
    elif template == "B":
        date_str = extract_value(r"Original Tax Invoice Date:.*?([A-Za-z]+\s+\d{1,2},?\s+\d{4})", text)
    elif template == "C":
        date_str = extract_value(r"Invoice Date:.*?([A-Za-z]+\s+\d{1,2}\s*,?\s+\d{4})", text)
    elif template == "D":
        date_str = extract_value(r"Document Date:.*?([A-Za-z]+\s+\d{1,2},?\s+\d{4})", text)
    else:
        date_str = ""
    try:
        return datetime.strptime(date_str.replace(',', ''), "%B %d %Y").strftime("%d/%m/%Y")
    except Exception:
        return ""

def process_pdf_by_template(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)
    template = detect_template(text)
    invoice_number, net_charges_usd, account_number, formatted_period = extract_common_fields(text, template == "B", template)

    match = re.search(r"(?:TOTAL AMOUNT )?DUE ON\s+([A-Za-z]+)\s+(\d{1,2})\s*,?\s+(\d{4})", text)
    if match:
        month, day, year = match.groups()
        try:
            formatted_due_date = datetime.strptime(f"{month} {day}, {year}", "%B %d, %Y").strftime("%d/%m/%Y")
        except ValueError:
            formatted_due_date = ""
    else:
        formatted_due_date = extract_due_date_fallback(pdf_bytes)

    today_date = datetime.today().strftime("%d/%m/%Y")
    #narration = f"{'TAX CREDIT  if template == 'B' else 'TAX INVOICE'}#{invoice_number}-AMAZON WEB SERVICES - {formatted_period} - AC NO: {account_number}"
    if template == "A":
        narration_prefix ="TAX INVOICE#"
        label="AMAZON WEB SERVICES EMEA SARL (AWS)"
    elif template == "B":
        narration_prefix ="TAX CREDIT NOTE#"
        label="AMAZON WEB SERVICES EMEA SARL (AWS)"
    elif template == "C":
        narration_prefix ="INVOICE#"
        label="AMAZON WEB SERVICES, INC.INVOICE"
    elif template == "D":
        narration_prefix ="INVOICE#"
        label="AWS MARKETPLACE INVOICE"
    else:
        narration_prefix ="INVOICE#"
        label="AMAZON MARKETPLACE INVOICE"
    narration= f"{narration_prefix}{invoice_number} - {label} - {formatted_period} - AC NO: {account_number}"  
    #narration = f"{'TAX CREDIT NOTE' if template == 'B' else 'TAX INVOICE'}#{invoice_number}-AMAZON WEB SERVICES - {formatted_period} - AC NO: {account_number}"
    bill_to = extract_bill_to(text, template)

    try:
        if template in ["C", "D"]:
            vat_usd_str = ""
            vat_aed_calculated = ""
            total_with_vat = net_charges_usd
        else:
            vat_usd = float(net_charges_usd) * 0.05
            vat_usd_str = f"{vat_usd:.2f}"
            vat_aed_calculated = f"{(vat_usd * 3.6725):.2f}"
            total_with_vat = f"{(vat_usd + float(net_charges_usd)):.2f}"
    except ValueError:
        vat_usd_str = ""
        vat_aed_calculated = ""
        total_with_vat = ""

    if template in ["C", "D"]:
        without_vat = net_charges_usd
        with_vat = ""
    else:
        without_vat = ""
        with_vat = net_charges_usd

    row = [
        today_date, invoice_number, without_vat, with_vat, narration,
        formatted_period, account_number, formatted_due_date,
        vat_usd_str, vat_aed_calculated, total_with_vat, "", "", bill_to
    ]
    return [row], template, text

def process_multiple_aws_pdfs(uploaded_files):
    all_rows = []
    template_map = {}
    text_map = {}
    for file in uploaded_files:
        pdf_bytes = file.read()
        rows, template, text = process_pdf_by_template(pdf_bytes)
        all_rows.extend(rows)

        bill_to = rows[0][-1].strip().upper()
        invoice_number = rows[0][1].strip()
        key = f"{bill_to}__{invoice_number}"

        template_map[key] = template
        text_map[key] = text

    return all_rows, template_map, text_map

def build_dnts_cnts_rows(rows, template_map, text_map):
    grouped = defaultdict(list)
    for row in rows:
        bill_to = row[-1].strip().upper()
        invoice_number = row[1].strip()
        key = f"{bill_to}__{invoice_number}"
        template_type = template_map.get(key, "A")
        group_key = f"{bill_to}__{'CNTS' if template_type == 'B' else 'DNTS'}"
        grouped[group_key].append(row)

    output_files = {}
    for group_key, group_rows in grouped.items():
        bill_to, file_type = group_key.split("__")
        is_cnts = file_type == "CNTS"
        today = datetime.today().strftime("%d/%m/%Y")

        if bill_to == "MINDWARE FZ LLC":
            supp_code = "SDIA035"
            doc_src_locn = "UJ000"
            location_code = "UJ200"
            division = "PUHU"
        else:
            supp_code = "STIA007"
            doc_src_locn = "TC000"
            location_code = "TC200"
            division = "PTCK"

        header_rows = []
        item_rows = []
        for idx, row in enumerate(group_rows, 1):
            invoice_number = row[1].strip()
            text = text_map.get(f"{bill_to}__{invoice_number}", "")
            supp_ref_date = extract_supp_ref_date(text, template_map.get(f"{bill_to}__{invoice_number}", "A"))
            header_rows.append([
                idx, today, supp_code, "USD", 0, doc_src_locn, location_code,
                row[4], invoice_number, supp_ref_date
            ])
            rate = row[2] if row[2] else row[3]
            item_rows.append([
                idx, idx, "AWS-NS-SW", row[4], "NA", "NA", "NOS", 1, 0, rate,
                14807, "", division, "GEN" if not is_cnts else "ZZ-COMM", ""
            ])

        output_files[f"{bill_to}__{file_type}"] = {
            "header": header_rows,
            "item": item_rows,
            "is_cnts": is_cnts
        }

    return output_files

