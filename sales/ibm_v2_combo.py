"""
Combo logic for IBM Excel-to-Excel (Template 1) and PDF-to-Excel (Template 2) in a single interface.
- Template 1: Excel-to-Excel logic from ibm_v2
- Template 2: PDF-to-Excel logic from ibm.py
"""

from extract_ibm_terms import extract_ibm_terms_text
from sales.ibm_v2 import (
    compare_mep_and_cost,
    check_bid_number_match,
    create_styled_excel_v2,
    parse_uploaded_excel
)
from ibm import (
    extract_ibm_data_from_pdf,
    create_styled_excel_template2,
    correct_descriptions,
    extract_last_page_text
)
from ibm_template2 import extract_ibm_template2_from_pdf
from template_detector import detect_ibm_template
from io import BytesIO
import logging


def process_ibm_combo(pdf_file, excel_file=None, master_csv=None, country="UAE"):
    """
    Unified processing for Template 1 (Excel-to-Excel) and Template 2 (PDF-to-Excel).
    - If excel_file is provided and template is 1: use Excel-to-Excel logic (ibm_v2)
    - If template is 2: use PDF-to-Excel logic (ibm.py)
    Returns: dict with keys: 'template', 'header_info', 'data', 'excel_bytes', 'mep_cost_msg', 'bid_number_error', 'error', 'ibm_terms_text'
    """
    result = {
        'template': None,
        'header_info': {},
        'data': [],
        'mep_cost_msg': None,
        'bid_number_error': None,
        'error': None,
        'ibm_terms_text': None,
        'columns': None,  # Add columns for DataFrame display
        'date_validation_msg': None  # Add date validation message
    }
    try:
        # Detect template
        template = detect_ibm_template(pdf_file)
        result['template'] = template
        pdf_file.seek(0)
        # Accept both '1' and 'template1' for template 1, and '2' and 'template2' for template 2
        if template in ('1', 'template1'):
            # Template 1: Excel-to-Excel logic
            header_info = {}
            data = []
            ibm_terms_text = ""
            # Extract header info from PDF
            try:
                _, extracted_header_info = extract_ibm_data_from_pdf(pdf_file)
                header_info.update(extracted_header_info)
                pdf_file.seek(0)
                ibm_terms_text = extract_ibm_terms_text(pdf_file)
            except Exception as e:
                result['error'] = f"Failed to extract header info or IBM Terms: {e}"
            # Extract data from Excel
            if excel_file:
                try:
                    data = parse_uploaded_excel(excel_file)
                except Exception as e:
                    result['error'] = f"Failed to extract data from Excel: {e}"
            result['header_info'] = header_info
            result['ibm_terms_text'] = ibm_terms_text

            # Date validation: Compare Excel dates with PDF dates for template 1
            if excel_file and data:
                try:
                    logging.info("Starting date validation for template 1")
                    pdf_file.seek(0)
                    pdf_data, _ = extract_ibm_data_from_pdf(pdf_file)
                    pdf_file.seek(0)
                    logging.info(f"Extracted {len(pdf_data)} rows from PDF")

                    # Create mapping of SKU to (start_date, end_date) from PDF
                    pdf_sku_dates = {}
                    for row in pdf_data:
                        if len(row) >= 5:
                            sku = str(row[0]).strip() if row[0] else ""
                            start_date = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                            end_date = str(row[4]).strip() if len(row) > 4 and row[4] else ""
                            if sku:
                                pdf_sku_dates[sku] = (start_date, end_date)
                    logging.info(f"Created PDF SKU mapping with {len(pdf_sku_dates)} SKUs")

                    # Validate dates for each Excel row
                    validation_messages = []
                    for i, row in enumerate(data, 1):
                        if len(row) >= 5:
                            sku = str(row[0]).strip() if row[0] else ""
                            excel_start = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                            excel_end = str(row[4]).strip() if len(row) > 4 and row[4] else ""
                            logging.info(f"Validating row {i}: SKU={sku}, Excel dates={excel_start}-{excel_end}")

                            if sku in pdf_sku_dates:
                                pdf_start, pdf_end = pdf_sku_dates[sku]
                                logging.info(f"Found PDF dates for SKU {sku}: {pdf_start}-{pdf_end}")
                                if excel_start == pdf_start and excel_end == pdf_end:
                                    validation_messages.append(f"Row {i} (SKU {sku}): Dates match between Excel and PDF")
                                else:
                                    validation_messages.append(f"Row {i} (SKU {sku}): Dates do NOT match - Excel: {excel_start}-{excel_end}, PDF: {pdf_start}-{pdf_end}")
                            else:
                                validation_messages.append(f"Row {i} (SKU {sku}): SKU not found in PDF data")
                                logging.warning(f"SKU {sku} not found in PDF data")

                    if validation_messages:
                        result['date_validation_msg'] = "\n".join(validation_messages)
                        logging.info(f"Generated {len(validation_messages)} validation messages")
                    else:
                        result['date_validation_msg'] = "No data to validate"
                        logging.info("No validation messages generated")

                except Exception as e:
                    logging.error(f"Date validation failed: {e}")
                    result['date_validation_msg'] = f"Failed to perform date validation: {e}"

            if country == "Qatar":
                columns = [
                    "SKU", "Product Description", "Quantity", "Start Date", "End Date",
                    "MEP Unit Price in USD", "Extended MEP Price USD", "Unit Partner Price USD", "Total Partner Price in USD"
                ]
                filtered_data = []
            
                excel_row = 18  # <-- Data starts on row 2 (change if your sheet starts later)
            
                for row in data:
                    sku = row[0] if len(row) > 0 else ""
                    desc = row[1] if len(row) > 1 else ""
                    qty  = row[2] if len(row) > 2 else 0
                    start_date = row[3] if len(row) > 3 else ""
                    end_date   = row[4] if len(row) > 4 else ""
                    raw_cost   = row[5] if len(row) > 5 else 0  # This is Extended MEP (total)
            
                    # Normalize numeric inputs so Excel formulas can compute correctly
                    def _to_float(x):
                        try:
                            return float(x) if x not in (None, "", "-",) else 0.0
                        except Exception:
                            return 0.0
            
                    qty_num  = _to_float(qty)
                    cost_num = _to_float(raw_cost)
            
                    # Build A1 references for this row
                    C = f"C{excel_row}"
                    F = f"F{excel_row}"
                    G = f"G{excel_row}"
                    H = f"H{excel_row}"
                    I = f"I{excel_row}"
                    J = f"J{excel_row}"
                    E= f"E{excel_row}"
            
                    # Visible Excel formulas
                    # unit_price_formula    = f"=IF({C}=0,0,{G}/{C})"     # F = unit price
                    unit_price_formula    = f"=ROUND({I}/{E},2)"        # F = unit price
                    unit_partner_formula  = f"=ROUND({H}*0.99,2)"       # H = 1% discount applied to unit price
                    total_partner_formula = f"={J}*{E}"                 # I = H * qty
            
                    filtered_data.append([
                        sku,                 # A
                        desc,                # B
                        qty_num,             # C (numeric)
                        start_date,          # D
                        end_date,            # E
                        unit_price_formula,  # F (formula)
                        cost_num,            # G (numeric)
                        unit_partner_formula,# H (formula)
                        total_partner_formula# I (formula)
                    ])
                    excel_row += 1
            
                result['data'] = filtered_data
                result['columns'] = columns
            else:
                result['data'] = data
                result['columns'] = ["SKU", "Description", "Quantity", "Start Date", "End Date", "Cost"]
        # MEP/cost check
            if header_info and data:
                result['mep_cost_msg'] = compare_mep_and_cost(header_info, data)
            # Bid number check
            if header_info and data and excel_file:
                pdf_bid_number = header_info.get('Bid Number', '')
                excel_file.seek(0)
                bid_number_match, bid_number_error = check_bid_number_match(excel_file, pdf_bid_number)
                if not bid_number_match:
                    result['bid_number_error'] = bid_number_error
            # Excel generation
            if header_info and not result['bid_number_error']:
                output = BytesIO()
                try:
                    create_styled_excel_v2(
                        data=result['data'] if result['data'] else [],
                        header_info=header_info,
                        logo_path="image.png",
                        output=output,
                        compliance_text="",
                        ibm_terms_text=ibm_terms_text,
                        country=country
                    )
                    result['excel_bytes'] = output.getvalue()
                except Exception as e:
                    result['error'] = f"Failed to create styled Excel: {e}"
        elif template in ('2', 'template2'):
            # Template 2: PDF-to-Excel logic (ibm_template2.py)
            try:
                data, header_info = extract_ibm_template2_from_pdf(pdf_file)
                pdf_file.seek(0)
                ibm_terms_text = extract_ibm_terms_text(pdf_file)
                result['header_info'] = header_info
                result['data'] = data
                result['ibm_terms_text'] = ibm_terms_text
                # Try to infer columns from data if available, else use generic
                if data and isinstance(data, list) and len(data) > 0:
                    if isinstance(data[0], (list, tuple)):
                        result['columns'] = [f"Col{i+1}" for i in range(len(data[0]))]
                    elif isinstance(data[0], dict):
                        result['columns'] = list(data[0].keys())
                    else:
                        result['columns'] = None
                else:
                    result['columns'] = None

                output = BytesIO()
                create_styled_excel_template2(
                    data=data,
                    header_info=header_info,
                    logo_path="image.png",
                    output=output,
                    compliance_text="",
                    ibm_terms_text=ibm_terms_text
                )
                result['excel_bytes'] = output.getvalue()
            except Exception as e:
                result['error'] = f"Failed to process Template 2: {e}"
        else:
            result['error'] = f"Unknown or unsupported template: {template}"
    except Exception as e:
        result['error'] = f"Failed to process file: {e}"
    return result
