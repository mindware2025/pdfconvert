import pdfplumber
import re
from datetime import datetime

AWS_OUTPUT_COLUMNS = [
    "DATE",
    "INV NUMBER",
    "WITHOUT VAT",
    "WITH VAT",
    "NARRATION",
    "Account Period",
    "A/C",
    "Due date",
    "Vat USD",
    "Vat AED",
    "Total USD",
    "Inv value",
    "Check",
    "Bill to"
]

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
            r"-USD\s*([0-9,]+\.[0-9]{2})\s+-AED\s*[0-9,]+\.[0-9]{2}\s+Net Charges\s*\(After Credits/Discounts, excl\. Tax\)",
            text
        )
    else:
        net_charges_usd = extract_value(
            r"USD\s*([0-9,]+\.[0-9]{2})\s+AED\s*[0-9,]+\.[0-9]{2}\s+Net Charges\s*\(After Credits/Discounts, excl\. Tax\)",
            text
        )

    net_charges_usd = net_charges_usd.replace(",", "")
    account_number = extract_value(r"(?:Account Number|رقم الحساب)[^\d]*?(\d{9,})", text)
    billing_start, billing_end = extract_values(
        r"This Tax (?:Invoice|Credit Note) is for the billing period\s*([A-Za-z]+ \d{1,2})\s*-\s*([A-Za-z]+ \d{1,2}, \d{4})",
        text
    )
    formatted_period = f"{billing_start} - {billing_end}" if billing_start and billing_end else ""
    return invoice_number, net_charges_usd, account_number, formatted_period

def process_template_a(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    invoice_number, net_charges_usd, account_number, formatted_period = extract_common_fields(text, is_credit_note=False)
    due_date = extract_value(r"DUE ON\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    formatted_due_date = format_date(due_date)
    today_date = datetime.today().strftime("%d/%m/%Y")

    narration = (
        f"TAX INVOICE#{invoice_number}-AMAZON WEB SERVICES EMEA SARL (AWS) "
        f"THIS INVOICE IS FOR THE BILLING PERIOD {formatted_period} - AC NO: {account_number}"
    )
    bill_to = "Mindware FZ LLC"

    try:
        vat_usd = float(net_charges_usd) * 0.05
        vat_usd_str = str(vat_usd)
        total_with_vat = str(vat_usd + float(net_charges_usd))
        vat_aed_calculated = str(vat_usd * 3.6725)
    except ValueError:
        vat_usd_str = ""
        vat_aed_calculated = ""
        total_with_vat = ""

    row = [
        today_date,
        invoice_number,
        "",                 # WITHOUT VAT
        net_charges_usd,    # WITH VAT
        narration,
        formatted_period,
        account_number,
        formatted_due_date,
        vat_usd_str,
        vat_aed_calculated,
        total_with_vat,
        "",
        "",
        bill_to
    ]
    return [row]

def process_template_b(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    credit_note_number, net_charges_usd, account_number, formatted_period = extract_common_fields(text, is_credit_note=True)
    formatted_due_date = ""  # Always empty for credit notes
    today_date = datetime.today().strftime("%d/%m/%Y")

    narration = (
        f"TAX CREDIT NOTE#{credit_note_number}-AMAZON WEB SERVICES EMEA SARL (AWS) "
        f"THIS CREDIT NOTE IS FOR THE BILLING PERIOD {formatted_period} - AC NO: {account_number}"
    )
    bill_to = "Mindware FZ LLC"

    try:
        vat_usd = float(net_charges_usd) * 0.05
        vat_usd_str = str(vat_usd)
        total_with_vat = str(vat_usd + float(net_charges_usd))
        vat_aed_calculated = str(vat_usd * 3.6725)
    except ValueError:
        vat_usd_str = ""
        vat_aed_calculated = ""
        total_with_vat = ""

    row = [
        today_date,
        credit_note_number,
        "",                 # WITHOUT VAT
        net_charges_usd,    # WITH VAT
        narration,
        formatted_period,
        account_number,
        formatted_due_date,
        vat_usd_str,
        vat_aed_calculated,
        total_with_vat,
        "",
        "",
        bill_to
    ]
    return [row]

def process_template_c(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    today_date = datetime.today().strftime("%d/%m/%Y")
    invoice_number = extract_value(r"Invoice Number:\s*(\d+)", text)
    total_due = extract_value(r"TOTAL AMOUNT DUE ON [A-Za-z]+\s+\d{1,2}\s*,\s*\d{4}\s*\$?([0-9,]+\.[0-9]{2})", text).replace(",", "")
    narration = f"INVOICE#{invoice_number}-AMAZON WEB SERVICES, INC. INVOICE - THIS INVOICE IS FOR THE BILLING PERIOD"
    billing_period = extract_value(r"This invoice is for the billing period\s*([A-Za-z]+\s+\d{1,2}\s*-\s*[A-Za-z]+\s+\d{1,2}\s*,\s*\d{4})", text)
    account_number = extract_value(r"Account number:\s*(\d+)", text)
    due_date = extract_value(r"TOTAL AMOUNT DUE ON\s*([A-Za-z]+\s+\d{1,2}\s*,\s*\d{4})", text)
    formatted_due_date = format_date(due_date)
    bill_to = "Mindware FZ LLC"

    row = [
        today_date,
        invoice_number,
        total_due,
        "",                 # WITH VAT
        narration,
        billing_period,
        account_number,
        formatted_due_date,
        "",
        "",             # Vat USD, Vat AED
        total_due,
        "",
        "",
        bill_to
    ]
    return [row]

def detect_template(text):
    if "Tax Credit Note" in text:
        return "B"
    elif "Tax Invoice" in text:
        return "A"
    elif "Amazon Web Services, Inc. Invoice" in text and "Invoice Number:" in text:
        return "C"
    return "Unknown"

def process_pdf_by_template(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    template = detect_template(text)

    if template == "A":
        return process_template_a(uploaded_file)
    elif template == "B":
        return process_template_b(uploaded_file)
    elif template == "C":
        return process_template_c(uploaded_file)
    else:
        return []

def process_multiple_aws_pdfs(uploaded_files):
    all_rows = []
    for file in uploaded_files:
        rows = process_pdf_by_template(file)
        all_rows.extend(rows)
    return all_rows

