"""Single entry point for the UAE rebate feature.

Kept deliberately independent of the quotation-generation pipeline
(sales/ibm_v2_combo.py, ibm.py, ibm_template2.py) so nothing here can affect
the existing quotation output. Any failure here is swallowed and returns
None -- the quotation flow must never be disrupted by this feature.
"""

from rebate.calculator import compute_rebate_rows
from rebate.extractor import extract_line_items
from rebate.workbook import build_rebate_workbook


def generate_rebate_excel(pdf_bytes, country):
    """Return .xlsx bytes for the rebate workbook, or None.

    None is returned when country isn't UAE, when no line items could be
    extracted, or if anything goes wrong during extraction/calculation --
    this feature must fail silently rather than break the IBM Quotation tool.
    """
    if country != "UAE":
        return None

    try:
        line_items = extract_line_items(pdf_bytes)
        if not line_items:
            return None
        rows, columns = compute_rebate_rows(line_items)
        if not columns:
            return None
        return build_rebate_workbook(rows, columns)
    except Exception:
        return None
