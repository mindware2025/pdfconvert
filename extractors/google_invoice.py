import pdfplumber
import re
from utils.helpers import normalize_line, format_invoice_date

GOOGLE_INVOICE_COLS = [
    "Domain name", "Customer ID", "Amount"
]

def extract_invoice_info(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        if not text:
            return None, None
        lines = text.splitlines()
        invoice_number = None
        invoice_date = None
        for line in lines:
            norm_line = normalize_line(line)
            if invoice_number is None and "Invoice number" in norm_line:
                match = re.search(r"Invoice number\s*:?\s*(\d{6,})", norm_line)
                if not match:
                    match = re.search(r"Invoice number\s*:?\s*([0-9]+)", norm_line)
                if match:
                    invoice_number = match.group(1)
            if invoice_date is None and "Invoice date" in norm_line:
                match = re.search(r"Invoice date\s*:?\s*([0-9]{1,2} [A-Za-z]+ [0-9]{4}|[0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", norm_line)
                if match:
                    invoice_date = match.group(1)
            if invoice_number and invoice_date:
                break
        return invoice_number, invoice_date

def extract_table_from_text(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.splitlines()
            in_table = False
            for i, line in enumerate(lines):
                if "Summary of costs by domain" in line:
                    in_table = True
                    continue
                if in_table:
                    if re.match(r"\d{1,2} \w+ \d{4} - \d{1,2} \w+ \d{4}", line):
                        continue
                    if all(h in line for h in ["Domain name", "Customer ID", "Amount"]):
                        continue
                    m = re.match(r"^([\w\-.]+)\s+(C\w+)\s+([\d,]+\.\d{2})$", line.strip(), re.IGNORECASE)
                    if m:
                        domain, customer_id, amount = m.groups()
                        rows.append([domain, customer_id, amount])
                    elif line.strip() == '' or 'Subtotal' in line:
                        in_table = False
    return rows