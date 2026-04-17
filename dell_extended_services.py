from io import BytesIO
from pathlib import Path
from typing import Optional, List, Dict
import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage

AED_RATE = 3.68


# ---------------- Helpers ----------------

def _text(v):
    return "" if v is None else str(v).strip()


def _usd_to_aed(val):
    try:
        return round(float(val) * AED_RATE, 2)
    except Exception:
        return 0.0


def _to_number(v):
    try:
        if v is None or v == "":
            return 0.0
        if isinstance(v, str):
            return float(v.replace(",", ""))
        return float(v)
    except Exception:
        return 0.0


def _get_local_logo_path() -> Optional[str]:
    base_dir = Path(__file__).resolve().parent
    for name in ("dell.png", "dell copy.png"):
        candidate = base_dir / name
        if candidate.exists():
            return str(candidate)
    return None


def _add_logo(ws, logo_bytes: Optional[bytes], anchor: str = "A1", width: int = 780, height: int = 52):
    if logo_bytes:
        try:
            img = XLImage(BytesIO(logo_bytes))
            img.width = width
            img.height = height
            ws.add_image(img, anchor)
            ws.merge_cells("A1:H4")
            ws.row_dimensions[1].height = 40
            ws.row_dimensions[2].height = 25
            return
        except Exception:
            pass

    local_logo = _get_local_logo_path()
    if local_logo:
        try:
            img = XLImage(local_logo)
            img.width = width
            img.height = height
            ws.add_image(img, anchor)
            ws.merge_cells("A1:H4")
            ws.row_dimensions[1].height = 40
            ws.row_dimensions[2].height = 25
        except Exception:
            pass


# ---------------- Metadata ----------------

def _extract_metadata(ws) -> Dict[str, str]:
    meta = {"quote_no": "", "date": "", "end_user": ""}

    for r in range(1, 60):
        row_values = [_text(ws.cell(r, c).value) for c in range(1, 11)]
        row_lower = [v.lower() for v in row_values]
        row_text = " ".join(row_lower)

        if "quote #:" in row_text:
            for idx, cell_text in enumerate(row_lower):
                if "quote #:" in cell_text:
                    for value in row_values[idx + 1:]:
                        if value:
                            meta["quote_no"] = value
                            break
                    if not meta["quote_no"] and r + 1 <= ws.max_row:
                        next_row_values = [_text(ws.cell(r + 1, c).value) for c in range(1, 11)]
                        for value in next_row_values:
                            if value:
                                meta["quote_no"] = value
                                break
                    break

        if any(cell_text.startswith("date") for cell_text in row_lower):
            for idx, cell_text in enumerate(row_lower):
                if cell_text.startswith("date"):
                    for value in row_values[idx + 1:]:
                        if value:
                            meta["date"] = value
                            break
                    break

        if "end user -" in row_text:
            meta["end_user"] = row_text.split("end user -", 1)[1].strip()

    return meta


# ---------------- Table ----------------

def _extract_extended_services_rows(ws) -> List[List]:
    rows = []
    start_row = None

    for r in range(1, ws.max_row + 1):
        row_text = " ".join(_text(ws.cell(r, c).value).lower() for c in range(1, 22))
        if "dell extended services details" in row_text:
            start_row = r
            break

    if not start_row:
        return rows

    header_row = None
    for r in range(start_row + 1, min(ws.max_row, start_row + 15) + 1):
        row_text = " ".join(_text(ws.cell(r, c).value).lower() for c in range(1, 22))
        if "asset" in row_text and "price after discount" in row_text:
            header_row = r
            break

    if not header_row:
        return rows

    target_headers = [
        "Asset", "Agreement ID", "Model", "Install At/Ship To",
        "Install At/Ship To City", "Install At/Ship To State",
        "Install At/Ship To Country", "LOB or Family", "Ship Date",
        "Service Contract Expiration", "Service Contract Description",
        "Services SKU", "New Contract Start Date", "New Contract End Date",
        "Quantity", "Price After Discount", "EOSS Date", "Product Type"
    ]

    header_cols = []
    for target in target_headers:
        found_col = 0
        target_lower = target.lower()
        for c in range(1, ws.max_column + 1):
            header_text = _text(ws.cell(header_row, c).value).lower()
            if not header_text:
                continue
            if target_lower in header_text or header_text in target_lower:
                found_col = c
                break
        header_cols.append(found_col)

    data_row = header_row + 1
    while data_row <= ws.max_row:
        row = [
            _text(ws.cell(data_row, c).value) if c else ""
            for c in header_cols
        ]
        if not any(row):
            data_row += 1
            continue

        if any("total" in cell.lower() for cell in row if cell):
            break

        # Keep USD value in column 16 for later AED conversion with margin.
        row[15] = _to_number(row[15])

        rows.append(row)
        data_row += 1

    return rows


# ---------------- Main Generator ----------------

def generate_dell_extended_services_quote(
    input_excel_bytes: bytes,
    logo_bytes: Optional[bytes] = None,
    margin_percent: float = 0.0,
) -> bytes:

    src_wb = openpyxl.load_workbook(BytesIO(input_excel_bytes), data_only=True)
    src_ws = src_wb.active

    meta = _extract_metadata(src_ws)
    rows = _extract_extended_services_rows(src_ws)

    wb = Workbook()
    ws = wb.active
    ws.title = "Quote"
    ws.sheet_view.showGridLines = False
    _add_logo(ws, None)

    # ----- Header -----
    ws["C5"] = "Quote No:"
    ws["D5"] = meta["quote_no"]

    ws["C6"] = "BDM:"
    ws["D6"] = "Omar Aboelawad"

    ws["C7"] = "Date:"
    ws["D7"] = meta["date"]

    ws["E6"] = "E-mail:"
    ws["F6"] = "o.aboelawad@mindware.net"

    ws["C8"] = "Quote Validity:"
    ws["D8"] = "30 days"

    ws["E7"] = "End User:"
    ws["F7"] = meta["end_user"]

    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for cell in ("C5", "C6", "C7", "C8", "D5", "D6", "D7", "D8", "E6", "F6", "E7", "F7"):
        ws[cell].font = Font(bold=True)
        ws[cell].border = border_thin

    for row in (5, 6, 7, 8):
        for col in (3, 4, 5, 6):
            cell = ws.cell(row=row, column=col)
            cell.border = border_thin

    ws.merge_cells("B11:M11")
    ws.merge_cells("B12:M12")
    ws.merge_cells("N12:R12")
    ws["B11"] = "Dell Extended Services Details"
    ws["B12"] = "Current Equipment Information"
    ws["N12"] = "Extended Service Information"
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for r, cols in [(11, range(2, 14)), (12, range(2, 14)), (12, range(14, 19))]:
        for c in cols:
            cell = ws.cell(row=r, column=c)
            cell.fill = header_fill
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for cell in ("B11", "B12", "N12"):
        ws[cell].font = Font(bold=True)

    # ----- Table -----
    headers = [
        "Asset", "Agreement ID", "Model", "Install At/Ship To",
        "Install At/Ship To City", "Install At/Ship To State",
        "Install At/Ship To Country", "LOB or Family", "Ship Date",
        "Service Contract Expiration", "Service Contract Description",
        "Services SKU", "New Contract Start Date", "New Contract End Date",
        "Quantity", "Price After Discount (AED)", "EOSS Date", "Product Type",
        "Margin"
    ]

    start = 13
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for c, h in enumerate(headers, 1):
        cell = ws.cell(start, c)
        cell.value = h
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = header_fill
        cell.border = border_thin

    r = start + 1
    for row in rows:
        price_usd = row[15]
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c)
            if c == 16:
                cell.value = f"=ROUND({price_usd}*(1+S{r})*{AED_RATE},2)"
                cell.number_format = '"AED" #,##0.00'
            else:
                cell.value = val
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="center", vertical="center")
        margin_cell = ws.cell(r, 19)
        margin_cell.value = margin_percent / 100.0
        margin_cell.number_format = '0.00%'
        margin_cell.border = border_thin
        margin_cell.alignment = Alignment(horizontal="center", vertical="center")
        r += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()