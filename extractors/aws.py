import fitz  # PyMuPDF
import re
import io
from datetime import datetime

# Define output columns
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
        net_charges_usd = extract_value(
            r"-USD\s*([0-9,]+\.[0-9]{2})\s+-AED\s*[0-9,]+\.[0-9]{2}\s+Net Charges", text
        )
    else:
        net_charges_usd = extract_value(
            r"USD\s*([0-9,]+\.[0-9]{2})\s+AED\s*[0-9,]+\.[0-9]{2}\s+Net Charges", text
        )

    net_charges_usd = net_charges_usd.replace(",", "")
    try:
        net_charges_usd = f"{float(net_charges_usd):.2f}"
    except ValueError:
        net_charges_usd = ""

    account_number = extract_value(r"(?:Account Number|رقم الحساب)[^\d]*?(\d{9,})", text)
    billing_start, billing_end = extract_values(
        r"This Tax (?:Invoice|Credit Note) is for the billing period\s*([A-Za-z]+ \d{1,2})\s*-\s*([A-Za-z]+ \d{1,2}, \d{4})",
        text
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

def process_template_a(pdf_bytes, template):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)

    invoice_number, net_charges_usd, account_number, formatted_period = extract_common_fields(text, is_credit_note=False)
    due_date = extract_value(r"DUE ON\s*([A-Za-z]+\s+\d{1,2},?\s+\d{4})", text)
    formatted_due_date = format_date(due_date)

    if not formatted_due_date or formatted_due_date == due_date:
       formatted_due_date = extract_due_date_fallback(pdf_bytes)

    today_date = datetime.today().strftime("%d/%m/%Y")
    narration = (
        f"TAX INVOICE#{invoice_number}-AMAZON WEB SERVICES EMEA SARL (AWS)  "
        f"THIS INVOICE IS FOR THE BILLING PERIOD {formatted_period} - AC NO: {account_number}"
    )
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
    return [row]

def process_template_b(pdf_bytes, template):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)

    credit_note_number, net_charges_usd, account_number, formatted_period = extract_common_fields(text, is_credit_note=True)
    formatted_due_date = ""
    today_date = datetime.today().strftime("%d/%m/%Y")
    narration = (
        f"TAX CREDIT NOTE#{credit_note_number}-AMAZON WEB SERVICES EMEA SARL (AWS) "
        f"THIS CREDIT NOTE IS FOR THE BILLING PERIOD {formatted_period} - AC NO: {account_number}"
    )
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
        today_date, credit_note_number, "", net_charges_usd, narration,
        formatted_period, account_number, formatted_due_date,
        vat_usd_str, vat_aed_calculated, total_with_vat, "", "", bill_to
    ]
    return [row]




def extract_total_due(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "TOTAL AMOUNT DUE ON" in line:
            if i + 1 < len(lines):
                amount_line = lines[i + 1]
                match = re.search(r"\$?([0-9,]+\.[0-9]{2})", amount_line)
                if match:
                    return match.group(1).replace(",", "")
    return ""

def extract_due_date(text):
    match = re.search(r"TOTAL AMOUNT DUE ON\s+([A-Za-z]+)\s+(\d{1,2})\s*,?\s*(\d{4})", text)
    if match:
        month, day, year = match.groups()
        try:
            return datetime.strptime(f"{month} {day} {year}", "%B %d %Y").strftime("%d/%m/%Y")
        except ValueError:
            return ""
    return ""


def extract_value(pattern, text):
    regex = re.compile(pattern, re.IGNORECASE)
    match = regex.search(text)
    return match.group(1).strip() if match else ""

def process_template_c(pdf_bytes, template):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)

    today_date = datetime.today().strftime("%d/%m/%Y")
    invoice_number = extract_value(r"Invoice Number:\s*(\d+)", text)

    # Total due
    total_due_raw = extract_total_due(text)
    try:
        total_due = f"{float(total_due_raw):.2f}"
    except ValueError:
        total_due = ""

    # Billing period
    billing_period = extract_value(
        r"billing period\s+([A-Za-z]+\s+\d{1,2}\s*-\s*[A-Za-z]+\s+\d{1,2}\s*,?\s*\d{4})",
        text
    )

    # Account number
    account_number = extract_value(r"Account number:\s*(\d+)", text)

    # Due date with fallback
    raw_due_date = extract_due_date(text)
    if not raw_due_date or raw_due_date == "":
        raw_due_date = extract_due_date_fallback(pdf_bytes)

    formatted_due_date = format_date(raw_due_date)

    # Bill to
    bill_to = extract_bill_to(text, template)

    # Narration
    narration = (
        f"INVOICE#{invoice_number}-AMAZON WEB SERVICES, INC. INVOICE - "
        f"THIS INVOICE IS FOR THE BILLING PERIOD {billing_period} - AC NO: {account_number}"
    )

    row = [
        today_date, invoice_number, total_due, "", narration,
        billing_period, account_number, formatted_due_date,
        "", "", total_due, "", "", bill_to
    ]
    return [row]



def process_template_d(pdf_bytes, template):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = "\n".join(page.get_text() for page in doc)

    today_date = datetime.today().strftime("%d/%m/%Y")
    document_number = extract_value(r"Document Number:\s*(\S+)", full_text)
    total_due_raw = extract_value(r"TOTAL AMOUNT DUE ON\s+[A-Za-z]+\s+\d{1,2},?\s+\d{4}\s+USD\s*([0-9,]+\.[0-9]{2})", full_text)

    try:
        total_due = f"{float(total_due_raw):.2f}"
    except ValueError:
        total_due = ""

    billing_period = extract_value(
        r"This Document is for the billing period\s*([A-Za-z]+ \d{1,2} - [A-Za-z]+ \d{1,2}, \d{4})",
        full_text
    )
    account_number = extract_value(r"Account number:\s*(\d+)", full_text)
    formatted_due_date = extract_due_date_fallback(pdf_bytes)
    bill_to = "Mindware FZ LLC" if "Mindware FZ LLC" in full_text else ""

    narration = f"INVOICE#{document_number} - AWS MARKETPLACE INVOICE - THIS INVOICE IS FOR THE BILLING PERIOD {billing_period} - AC NO: {account_number}"

    row = [
        today_date, document_number, total_due, "", narration,
        billing_period, account_number, formatted_due_date,
        "", "", total_due, "", "", bill_to
    ]
    return [row]

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

    if template == "A":
        return process_template_a(pdf_bytes, template)
    elif template == "B":
        return process_template_b(pdf_bytes, template)
    elif template == "C":
        return process_template_c(pdf_bytes, template)
    elif template == "D":
        return process_template_d(pdf_bytes, template)
    else:
        return []

def process_multiple_aws_pdfs(uploaded_files):
    all_rows = []
    for uploaded_file in uploaded_files:
        pdf_bytes = uploaded_file.read()
        pdf_stream = io.BytesIO(pdf_bytes)
        rows = process_pdf_by_template(pdf_stream)
        if rows:
            all_rows.extend(rows)
    return all_rows

