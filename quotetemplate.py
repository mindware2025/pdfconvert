from io import BytesIO
import openpyxl

def detect_dell_template(input_bytes: bytes) -> str:
    """
    Detect Dell template type.
    Returns:
      - 'extended_services'
      - 'standard_quote'
    """
    # PDFs always go to existing logic
    if input_bytes.lstrip().startswith(b"%PDF"):
        return "standard_quote"

    try:
        wb = openpyxl.load_workbook(BytesIO(input_bytes), data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=80, max_col=10):
            for cell in row:
                if isinstance(cell.value, str) and \
                   "dell extended services details" in cell.value.lower():
                    return "extended_services"
    except Exception:
        pass

    return "standard_quote"
