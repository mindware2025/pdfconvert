import fitz  # PyMuPDF
import re
import io
from datetime import datetime
from collections import defaultdict

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
    match = re.search(r"TOTAL AMOUNT DUE ON\s+([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})", normalized_text)
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

def extract_values(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.groups() if match else ("", "")

def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%B %d, %Y").strftime("%d/%m/%Y")
    except Exception:
        return date_str

def extract_common_fields(text, is_credit_note=False):
    invoice_number = extract_value(r"(EU[A-Z]+[0-9]{2}-\d+)", text)
    if is_credit_note:
        net_charges_usd = extract_value(r"-USD\s*([0-9,]+\.[0-9]{2})", text)
    else:
        net_charges_usd = extract_value(r"USD\s*([0-9,]+\.[0-9]{2})", text)
    net_charges_usd = net_charges_usd.replace(",", "")
    try:
        net_charges_usd = f"{float(net_charges_usd):.2f}"
    except ValueError:
        net_charges_usd = ""
    account_number = extract_value(r"(?:Account Number|رقم الحساب)[^\d]*?(\d{9,})", text)
    billing_start, billing_end = extract_values(
        r"billing period\s*([A-Za-z]+ \d{1,2})\s*-\s*([A-Za-z]+ \d{1,2}, \d{4})", text
    )
    formatted_period = f"{billing_start} - {billing_end}" if billing_start and billing_end else ""
    return invoice_number, net_charges_usd, account_number, formatted_period

def extract_bill_to(text, template):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if template == "C" and "Bill to Address" in line:
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
        elif template == "D" and "Address:" in line:
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
        elif template != "C" and ("Address" in line or "العنوان" in line):
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
        else:
            continue
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

def process_pdf_by_template(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)
    template = detect_template(text)
    invoice_number, net_charges_usd, account_number, formatted_period = extract_common_fields(text, template == "B")
    due_date = extract_value(r"DUE ON\s*([A-Za-z]+\s+\d{1,2},?\s+\d{4})", text)
    formatted_due_date = format_date(due_date)
    if not formatted_due_date or formatted_due_date == due_date:
        formatted_due_date = extract_due_date_fallback(pdf_bytes)
    today_date = datetime.today().strftime("%d/%m/%Y")
    narration = f"{'TAX CREDIT NOTE' if template == 'B' else 'TAX INVOICE'}#{invoice_number}-AMAZON WEB SERVICES - THIS INVOICE IS FOR THE BILLING PERIOD {formatted_period} - AC NO: {account_number}"
    bill_to = extract_bill_to(text, template)
    try:
        vat_usd = float(net_charges_usd) * 0.05
        vat_usd_str = f"{vat_usd:.2f}"
        total_with_vat = f"{(vat_usd + float(net_charges_usd)):.2f}"
        vat_aed_calculated = f"{(vat_usd * 3.6725):.2f}"
    except ValueError:
        vat_usd_str = ""
        vat_aed_calculated = ""
        total_with_vat = ""
    row = [
        today_date, invoice_number, "", net_charges_usd, narration,
        formatted_period, account_number, formatted_due_date,
        vat_usd_str, vat_aed_calculated, total_with_vat, "", "", bill_to
    ]
    return [row], template

def process_multiple_aws_pdfs(uploaded_files):
    all_rows = []
    template_map = {}
    for uploaded_file in uploaded_files:
        pdf_bytes = uploaded_file.read()
        pdf_stream = io.BytesIO(pdf_bytes)
        rows, template = process_pdf_by_template(pdf_stream)
        if rows:
            for row in rows:
                all_rows.append(row)
                template_map[row[-1]] = template
    return all_rows, template_map

def build_dnts_cnts_rows(rows, template_map):
    grouped = defaultdict(list)
    for row in rows:
        bill_to = row[-1]
        grouped[bill_to].append(row)
    output_files = {}
    for bill_to, group_rows in grouped.items():
        template_type = template_map.get(bill_to, "A")
        is_cnts = template_type == "B"
        today = datetime.today().strftime("%d/%m/%Y")
        if bill_to == "Mindware FZ LLC":
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
            header_rows.append([
                idx, today, supp_code, "USD", 0, doc_src_locn, location_code,
                row[4], row[1], today
            ])
            rate = row[2] if row[2] else row[3]
            item_rows.append([
                idx, idx, "AWS-NS-SW", row[4], "NA", "NA", "NOS", 1, 0, rate,
                14807, "", division, "GEN" if not is_cnts else "ZZ-COMM", ""
            ])
        output_files[bill_to] = {
            "header": header_rows,
            "item": item_rows,
            "is_cnts": is_cnts
        }
    return output_files
