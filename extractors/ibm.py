import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO
from openpyxl.styles import NamedStyle
from openpyxl import Workbook

def extract_ibm_data_from_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    lines = []
    for page in doc:
        text = page.get_text()
        lines.extend(text.split('\n'))

    extracted_data = []
    for line in lines:
        # Match lines that start with a number and a part number
        if line.strip().startswith(("1 D28AYLL", "2 E0R1HLL", "3 E0R1HLL")):
            parts = line.strip().split()
            try:
                extracted_data.append({
                    "Part Number": parts[1],
                    "Description": " ".join(parts[2:6]),
                    "Coverage Start": parts[6],
                    "Coverage End": parts[7],
                    "Quantity": int(parts[8]),
                    "Unit SVP": float(parts[9].replace('.', '').replace(',', '.')),
                    "Extended SVP": float(parts[10].replace('.', '').replace(',', '.')),
                    "Discount %": float(parts[11].replace('.', '').replace(',', '.')),
                    "Bid Unit SVP": float(parts[12].replace('.', '').replace(',', '.')),
                    "Bid Extended SVP": float(parts[13].replace('.', '').replace(',', '.')),
                    "Line Total": float(parts[13].replace('.', '').replace(',', '.'))
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

        date_style = NamedStyle(name="date_style", number_format='MM/DD/YYYY')
        if "date_style" not in workbook.named_styles:
            workbook.add_named_style(date_style)

        for col_idx, col_name in enumerate(df.columns, start=1):
            if 'Coverage' in col_name:
                for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.style = date_style

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