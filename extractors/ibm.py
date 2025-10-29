import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO
from openpyxl.styles import NamedStyle
from openpyxl import Workbook

def parse_number(value):
    """Convert European-style numbers to float."""
    try:
        return float(value.replace('.', '').replace(',', '.'))
    except:
        return None

def extract_ibm_data_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    lines = []
    for page in doc:
        text = page.get_text()
        lines.extend(text.split('\n'))

    extracted_data = []
    for line in lines:
        tokens = line.strip().split()
        if len(tokens) >= 15 and tokens[0].isdigit():
            try:
                part_number = tokens[1]
                description = ' '.join(tokens[2:7])
                coverage_start = tokens[7]
                coverage_end = tokens[8]
                quantity = int(tokens[9])
                unit_svp = parse_number(tokens[10])
                extended_svp = parse_number(tokens[11])
                discount = parse_number(tokens[12])
                bid_unit_svp = parse_number(tokens[13])
                bid_extended_svp = parse_number(tokens[14])
                line_total = bid_extended_svp

                extracted_data.append({
                    "Part Number": part_number,
                    "Description": description,
                    "Coverage Start": coverage_start,
                    "Coverage End": coverage_end,
                    "Quantity": quantity,
                    "Unit SVP": unit_svp,
                    "Extended SVP": extended_svp,
                    "Discount %": discount,
                    "Bid Unit SVP": bid_unit_svp,
                    "Bid Extended SVP": bid_extended_svp,
                    "Line Total": line_total
                })
            except Exception:
                continue

    return pd.DataFrame(extracted_data)

def generate_ibm_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='BoQ')
        workbook = writer.book
        worksheet = writer.sheets['BoQ']

        # Date formatting
        date_style = NamedStyle(name="date_style", number_format='MM/DD/YYYY')
        if "date_style" not in workbook.named_styles:
            workbook.add_named_style(date_style)

        for col_idx, col_name in enumerate(df.columns, start=1):
            if 'Coverage' in col_name:
                for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.style = date_style

        # Currency formatting
        currency_style = NamedStyle(name="currency_style", number_format='"$"#,##0.00')
        if "currency_style" not in workbook.named_styles:
            workbook.add_named_style(currency_style)

        for col_name in ["Unit SVP", "Extended SVP", "Bid Unit SVP", "Bid Extended SVP", "Line Total"]:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.style = currency_style

    output.seek(0)
    return output