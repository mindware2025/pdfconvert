import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


PURPLE_FILL = PatternFill(fill_type="solid", fgColor="3F3D9E")
WHITE_BOLD_FONT = Font(color="FFFFFF", bold=True)
BOLD_FONT = Font(bold=True)
THIN_SIDE = Side(style="thin", color="000000")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=False)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=False)
TOP_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)

SUPPLIER_TEXT = (
    "Mindware Fz LLC\n"
    "P.O.Box 55609\n"
    "Jabel Ali, United Arab Emirates\n"
    "Tel : + 971 4500600\n"
    "Email: outboundjebelali@mindware.ae\n"
    "VAT TRN No : 100019912300003"
)

ADDRESS_MIN_ROWS = 2


def style_range(worksheet, cell_range: str, fill=None, font=None, border=None, alignment=None) -> None:
    for row in worksheet[cell_range]:
        for cell in row:
            if fill:
                cell.fill = fill
            if font:
                cell.font = font
            if border:
                cell.border = border
            if alignment:
                cell.alignment = alignment


def apply_border_to_range(worksheet, start_row: int, end_row: int, start_col: int, end_col: int) -> None:
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            worksheet.cell(row=row, column=col).border = THIN_BORDER


set_outer_border = apply_border_to_range


def to_number_if_possible(value):
    if isinstance(value, (int, float)):
        return value
    if not isinstance(value, str):
        return value
    candidate = value.replace(",", "").strip()
    if not candidate:
        return ""
    try:
        return float(candidate)
    except ValueError:
        return value


def count_display_lines(value: str) -> int:
    if not value:
        return 1
    return len([line for line in str(value).splitlines() if line.strip()]) or 1


# Approximate character capacity per address column block
BILL_TO_CHARS = 38    # cols A-D (~21+36+16+11 char widths)
SHIP_TO_CHARS = 26    # cols E-F (~15+14)
SUPPLIER_CHARS = 30   # cols G-H (~22+16)


def estimate_wrapped_lines(text: str, col_width_chars: int) -> int:
    """Count visual lines after word-wrap for a block of text in a merged cell."""
    if not text:
        return 1
    total = 0
    for line in text.splitlines():
        line = line.strip()
        if not line:
            total += 1
        else:
            total += max(1, -(-len(line) // col_width_chars))  # ceiling div
    return total or 1


def estimate_desc_row_height(desc: str, col_width_chars: int = 36) -> int:
    """Return row height in pts sufficient to show full desc text."""
    wrapped = estimate_wrapped_lines(desc, col_width_chars)
    return max(15, wrapped * 15)


def compute_address_rows(bill_to: str, ship_to: str) -> int:
    lines = max(
        estimate_wrapped_lines(bill_to, BILL_TO_CHARS),
        estimate_wrapped_lines(ship_to, SHIP_TO_CHARS),
        estimate_wrapped_lines(SUPPLIER_TEXT, SUPPLIER_CHARS),
    )
    return max(lines, ADDRESS_MIN_ROWS)


def set_merged_block_row_heights(worksheet, start_row: int, end_row: int, total_lines: int) -> None:
    num_rows = end_row - start_row + 1
    # Each line needs ~15 pts; spread evenly across the merged rows
    total_height = max(num_rows * 15, total_lines * 15)
    base_height = total_height / num_rows
    for row in range(start_row, end_row + 1):
        worksheet.row_dimensions[row].height = base_height


def get_layout(address_rows: int, item_count: int) -> dict:
    """Compute all dynamic row numbers from address block size and item count."""
    addr_start = 9
    addr_end = addr_start + address_rows - 1
    header_row = addr_end + 1
    items_start = header_row + 1
    items_end = items_start + max(item_count, 1) - 1
    freight_row = items_end + 1
    total_row = freight_row + 1
    net_total_row = total_row + 1
    words_row = net_total_row + 2   # one blank gap row between net total and words
    return {
        "addr_start": addr_start,
        "addr_end": addr_end,
        "header_row": header_row,
        "items_start": items_start,
        "items_end": items_end,
        "freight_row": freight_row,
        "total_row": total_row,
        "net_total_row": net_total_row,
        "words_row": words_row,
    }


# ---------------------------------------------------------------------------
# comm-inv sheet
# ---------------------------------------------------------------------------

def build_comm_inv_static(worksheet) -> None:
    """Write static structure: title, purple bars, rows 5-8 shells."""
    worksheet.title = "comm-inv"

    widths = {"A": 21, "B": 36, "C": 16, "D": 11, "E": 15, "F": 14, "G": 22, "H": 16}
    for col, w in widths.items():
        worksheet.column_dimensions[col].width = w

    worksheet.row_dimensions[2].height = 28
    worksheet.merge_cells("A2:H2")
    worksheet["A2"] = "Commercial Invoice"
    worksheet["A2"].font = Font(bold=True, size=16)
    worksheet["A2"].alignment = CENTER

    worksheet.row_dimensions[4].height = 18
    worksheet.merge_cells("A4:H4")
    style_range(worksheet, "A4:H4", fill=PURPLE_FILL, border=THIN_BORDER)

    for row in (5, 6, 7):
        worksheet.row_dimensions[row].height = 18

    worksheet.merge_cells("A5:F5")
    worksheet.merge_cells("A6:F6")
    worksheet.merge_cells("A7:F7")
    apply_border_to_range(worksheet, 5, 7, 1, 6)

    worksheet.merge_cells("G5:H5")
    worksheet.merge_cells("G6:H6")
    worksheet.merge_cells("G7:H7")
    apply_border_to_range(worksheet, 5, 7, 7, 8)

    for addr in ("A5", "A6", "A7"):
        worksheet[addr].font = BOLD_FONT
        worksheet[addr].alignment = LEFT
    for addr in ("G5", "G6", "G7"):
        worksheet[addr].font = BOLD_FONT
        worksheet[addr].alignment = LEFT

    worksheet.row_dimensions[8].height = 18
    worksheet.merge_cells("A8:D8")
    worksheet.merge_cells("E8:F8")
    worksheet.merge_cells("G8:H8")
    worksheet["A8"] = "Bill To"
    worksheet["E8"] = "Ship To"
    worksheet["G8"] = "Supplier"
    style_range(worksheet, "A8:H8", fill=PURPLE_FILL, font=WHITE_BOLD_FONT, border=THIN_BORDER)
    worksheet["A8"].alignment = LEFT
    worksheet["E8"].alignment = LEFT
    worksheet["G8"].alignment = LEFT


def fill_comm_inv_sheet(worksheet, fields: dict, item_count: int) -> None:
    bill_to = fields.get("bill_to", "")
    ship_to = fields.get("ship_to", "")
    address_rows = compute_address_rows(bill_to, ship_to)
    L = get_layout(address_rows, item_count)

    # Rows 5-7
    worksheet["A5"] = f"Payment Term: {fields.get('payment_term', '')}"
    worksheet["A6"] = f"Inco Terms: {fields.get('inco_terms', '')}"
    worksheet["A7"] = f"Customer PO: {fields.get('customer_po', '')}"
    for addr in ("A5", "A6", "A7"):
        worksheet[addr].font = BOLD_FONT
        worksheet[addr].alignment = LEFT

    worksheet["G5"] = f"Commercial Invoice No : {fields.get('commercial_invoice_no', '')}"
    worksheet["G6"] = f"Date: {fields.get('date', '')}"
    worksheet["G7"] = f"Currency: {fields.get('currency', '')}"
    for addr in ("G5", "G6", "G7"):
        worksheet[addr].font = BOLD_FONT
        worksheet[addr].alignment = LEFT

    # Address block — merged, sized to content
    addr_s, addr_e = L["addr_start"], L["addr_end"]
    worksheet.merge_cells(start_row=addr_s, start_column=1, end_row=addr_e, end_column=4)
    worksheet.merge_cells(start_row=addr_s, start_column=5, end_row=addr_e, end_column=6)
    worksheet.merge_cells(start_row=addr_s, start_column=7, end_row=addr_e, end_column=8)
    apply_border_to_range(worksheet, addr_s, addr_e, 1, 8)

    c = worksheet.cell(row=addr_s, column=1)
    c.value = bill_to
    c.alignment = TOP_LEFT

    c = worksheet.cell(row=addr_s, column=5)
    c.value = ship_to
    c.alignment = TOP_LEFT

    c = worksheet.cell(row=addr_s, column=7)
    c.value = SUPPLIER_TEXT
    c.font = BOLD_FONT
    c.alignment = TOP_LEFT

    set_merged_block_row_heights(worksheet, addr_s, addr_e, address_rows)

    # Column headers row — immediately after address block
    hdr = L["header_row"]
    worksheet.row_dimensions[hdr].height = 22
    headers = ["Item Code", "Desc", "Case#", "Origin", "HS Code", "Qty", "Unit Price", "Amount"]
    for idx, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=hdr, column=idx)
        cell.value = header
        cell.fill = PURPLE_FILL
        cell.font = WHITE_BOLD_FONT
        cell.border = THIN_BORDER
        cell.alignment = LEFT if header in {"Item Code", "Desc"} else CENTER

    # Footer rows
    fr = L["freight_row"]
    worksheet.merge_cells(start_row=fr, start_column=1, end_row=fr, end_column=6)
    apply_border_to_range(worksheet, fr, fr, 1, 8)
    worksheet.cell(row=fr, column=7).value = "Freight Charges"
    worksheet.cell(row=fr, column=7).font = BOLD_FONT
    worksheet.cell(row=fr, column=7).alignment = RIGHT
    worksheet.cell(row=fr, column=7).border = THIN_BORDER
    worksheet.cell(row=fr, column=8).alignment = RIGHT
    worksheet.cell(row=fr, column=8).border = THIN_BORDER
    if fields.get("freight_charges", "") != "":
        worksheet.cell(row=fr, column=8).value = to_number_if_possible(fields.get("freight_charges", ""))

    tr = L["total_row"]
    worksheet.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=6)
    apply_border_to_range(worksheet, tr, tr, 1, 8)
    worksheet.cell(row=tr, column=7).value = "Total Amount"
    worksheet.cell(row=tr, column=7).font = BOLD_FONT
    worksheet.cell(row=tr, column=7).alignment = RIGHT
    worksheet.cell(row=tr, column=7).border = THIN_BORDER
    worksheet.cell(row=tr, column=8).font = BOLD_FONT
    worksheet.cell(row=tr, column=8).alignment = RIGHT
    worksheet.cell(row=tr, column=8).border = THIN_BORDER
    worksheet.cell(row=tr, column=8).value = f"=SUM(H{L['items_start']}:H{L['items_end']})+H{fr}"

    nr = L["net_total_row"]
    worksheet.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=6)
    apply_border_to_range(worksheet, nr, nr, 1, 8)
    worksheet.cell(row=nr, column=7).value = "Net Total"
    worksheet.cell(row=nr, column=7).font = BOLD_FONT
    worksheet.cell(row=nr, column=7).alignment = RIGHT
    worksheet.cell(row=nr, column=7).border = THIN_BORDER
    worksheet.cell(row=nr, column=8).font = BOLD_FONT
    worksheet.cell(row=nr, column=8).alignment = RIGHT
    worksheet.cell(row=nr, column=8).border = THIN_BORDER
    worksheet.cell(row=nr, column=8).value = f"=H{tr}"

    wr = L["words_row"]
    worksheet.row_dimensions[wr - 1].height = 6  # small blank gap
    worksheet.merge_cells(start_row=wr, start_column=1, end_row=wr + 1, end_column=6)
    apply_border_to_range(worksheet, wr, wr + 1, 1, 6)
    words_text = (
        f"Total in Words : {fields.get('total_in_words', '')}"
        if fields.get("total_in_words")
        else "Total in Words :"
    )
    c = worksheet.cell(row=wr, column=1)
    c.value = words_text
    c.font = BOLD_FONT
    c.alignment = TOP_LEFT

    worksheet.merge_cells(start_row=wr, start_column=7, end_row=wr + 1, end_column=8)
    apply_border_to_range(worksheet, wr, wr + 1, 7, 8)
    c = worksheet.cell(row=wr, column=7)
    c.value = "Mindware FZ LLC"
    c.font = BOLD_FONT
    c.alignment = CENTER


def fill_comm_inv_items(worksheet, items: list[dict], address_rows: int) -> None:
    L = get_layout(address_rows, len(items))
    items_start = L["items_start"]

    for offset, item in enumerate(items):
        row = items_start + offset
        desc_val = str(item.get("desc", ""))
        needs_wrap = "\n" in desc_val or len(desc_val) > 36
        worksheet.row_dimensions[row].height = estimate_desc_row_height(desc_val)

        def w(col, value, align):
            c = worksheet.cell(row=row, column=col)
            c.value = value
            c.alignment = align
            c.border = THIN_BORDER

        w(1, item.get("item_code", ""), LEFT)
        w(2, item.get("desc", ""), TOP_LEFT if needs_wrap else LEFT)
        w(3, item.get("case_no", ""), CENTER)
        w(4, item.get("origin", ""), CENTER)
        w(5, item.get("hs_code", ""), CENTER)
        w(6, item.get("qty", ""), CENTER)
        w(7, item.get("unit_price", ""), RIGHT)
        w(8, item.get("amount", ""), RIGHT)

    # ensure borders on all item rows even if items list is empty
    apply_border_to_range(worksheet, L["items_start"], L["items_end"], 1, 8)


def fill_comm_inv_unmatched_items(worksheet, items: list[dict], address_rows: int, item_count: int) -> None:
    if not items:
        return

    L = get_layout(address_rows, item_count)
    net_total_row = L["net_total_row"]
    total_row = L["total_row"]
    words_row = L["words_row"]
    start_row = words_row + 4

    hdr = start_row - 1
    for col, label in [(6, "SOB No"), (7, "Other SOB Items"), (8, "Amount")]:
        c = worksheet.cell(row=hdr, column=col)
        c.value = label
        c.font = BOLD_FONT
        c.border = THIN_BORDER
        c.alignment = CENTER

    for offset, item in enumerate(items):
        row = start_row + offset
        worksheet.cell(row=row, column=6).value = item.get("sob_reference", "")
        worksheet.cell(row=row, column=6).alignment = CENTER
        worksheet.cell(row=row, column=6).border = THIN_BORDER
        worksheet.cell(row=row, column=7).value = item.get("item_code", "")
        worksheet.cell(row=row, column=7).alignment = LEFT
        worksheet.cell(row=row, column=7).border = THIN_BORDER
        worksheet.cell(row=row, column=8).value = item.get("amount", "")
        worksheet.cell(row=row, column=8).alignment = RIGHT
        worksheet.cell(row=row, column=8).border = THIN_BORDER

    other_total_row = start_row + len(items)
    worksheet.cell(row=other_total_row, column=7).value = "Total"
    worksheet.cell(row=other_total_row, column=7).font = BOLD_FONT
    worksheet.cell(row=other_total_row, column=7).border = THIN_BORDER
    worksheet.cell(row=other_total_row, column=7).alignment = RIGHT
    worksheet.cell(row=other_total_row, column=8).value = f"=SUM(H{start_row}:H{other_total_row - 1})"
    worksheet.cell(row=other_total_row, column=8).font = BOLD_FONT
    worksheet.cell(row=other_total_row, column=8).border = THIN_BORDER
    worksheet.cell(row=other_total_row, column=8).alignment = RIGHT

    apply_border_to_range(worksheet, hdr, other_total_row, 6, 8)

    # Update Net Total to include unmatched SOB items
    worksheet.cell(row=net_total_row, column=8).value = f"=H{total_row}+H{other_total_row}"


# ---------------------------------------------------------------------------
# pack_list sheet
# ---------------------------------------------------------------------------

def build_pack_list_sheet(worksheet) -> None:
    worksheet.title = "pack_list"

    widths = {"A": 21, "B": 34, "C": 15, "D": 15, "E": 15, "F": 14, "G": 13, "H": 10}
    for col, w in widths.items():
        worksheet.column_dimensions[col].width = w

    worksheet.row_dimensions[2].height = 28
    worksheet.row_dimensions[4].height = 18

    worksheet.merge_cells("A2:H2")
    worksheet["A2"] = "Packing List"
    worksheet["A2"].font = Font(bold=True, size=16)
    worksheet["A2"].alignment = CENTER

    worksheet.merge_cells("A4:H4")
    style_range(worksheet, "A4:H4", fill=PURPLE_FILL, border=THIN_BORDER)

    worksheet.merge_cells("A5:E6")
    worksheet.merge_cells("F5:H5")
    worksheet.merge_cells("F6:H6")
    apply_border_to_range(worksheet, 5, 6, 1, 8)
    worksheet["F5"] = "No. :"
    worksheet["F6"] = "Date :"
    worksheet["F5"].font = BOLD_FONT
    worksheet["F6"].font = BOLD_FONT
    worksheet["F5"].alignment = LEFT
    worksheet["F6"].alignment = LEFT

    worksheet["A7"] = "Bill To"
    worksheet["C7"] = "Ship To"
    worksheet["F7"] = "Supplier"
    style_range(worksheet, "A7:H7", fill=PURPLE_FILL, font=WHITE_BOLD_FONT, border=THIN_BORDER)
    worksheet.merge_cells("A7:B7")
    worksheet.merge_cells("C7:E7")
    worksheet.merge_cells("F7:H7")
    worksheet["A7"].alignment = LEFT
    worksheet["C7"].alignment = LEFT
    worksheet["F7"].alignment = LEFT


def fill_pack_list_sheet(worksheet, fields: dict, address_rows: int) -> None:
    addr_s = 8
    addr_e = addr_s + address_rows - 1

    worksheet.merge_cells(start_row=addr_s, start_column=1, end_row=addr_e, end_column=2)
    worksheet.merge_cells(start_row=addr_s, start_column=3, end_row=addr_e, end_column=5)
    worksheet.merge_cells(start_row=addr_s, start_column=6, end_row=addr_e, end_column=8)
    apply_border_to_range(worksheet, addr_s, addr_e, 1, 8)

    worksheet["F5"] = '="No. : " & \'comm-inv\'!G5'
    worksheet["F6"] = '="Date : " & \'comm-inv\'!G6'

    c = worksheet.cell(row=addr_s, column=1)
    c.value = "='comm-inv'!A9"
    c.alignment = TOP_LEFT

    c = worksheet.cell(row=addr_s, column=3)
    c.value = "='comm-inv'!E9"
    c.alignment = TOP_LEFT

    c = worksheet.cell(row=addr_s, column=6)
    c.value = SUPPLIER_TEXT
    c.font = BOLD_FONT
    c.alignment = TOP_LEFT

    set_merged_block_row_heights(
        worksheet,
        addr_s,
        addr_e,
        max(
            count_display_lines(fields.get("bill_to", "")),
            count_display_lines(fields.get("ship_to", "")),
            count_display_lines(SUPPLIER_TEXT),
        ),
    )

    hdr_row = addr_e + 1
    worksheet.row_dimensions[hdr_row].height = 22
    headers = ["Item Code", "Desc", "Case#", "Origin", "HS Code", "Qty", "Weight", "Package"]
    for idx, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=hdr_row, column=idx)
        cell.value = header
        cell.fill = PURPLE_FILL
        cell.font = WHITE_BOLD_FONT
        cell.border = THIN_BORDER
        cell.alignment = LEFT if header in {"Item Code", "Desc"} else CENTER


def fill_pack_list_items(worksheet, items: list[dict], address_rows: int) -> None:
    addr_e = 8 + address_rows - 1
    hdr_row = addr_e + 1
    items_start = hdr_row + 1
    total_row = items_start + max(len(items), 1)
    summary_start = total_row + 2
    case_hdr_row = summary_start + 3
    case_data_start = case_hdr_row + 1

    apply_border_to_range(worksheet, items_start, total_row - 1, 1, 8)
    apply_border_to_range(worksheet, total_row, total_row, 5, 8)
    apply_border_to_range(worksheet, summary_start, summary_start + 1, 2, 4)
    apply_border_to_range(worksheet, case_hdr_row, case_hdr_row, 2, 4)

    worksheet.cell(row=addr_e + 3, column=8).value = "Mindware FZ LLC"
    worksheet.cell(row=addr_e + 3, column=8).font = BOLD_FONT
    worksheet.cell(row=addr_e + 3, column=8).alignment = CENTER

    worksheet.cell(row=summary_start, column=2).value = "Total No of Cases"
    worksheet.cell(row=summary_start, column=2).font = BOLD_FONT
    worksheet.cell(row=summary_start + 1, column=2).value = "Total Gross Weight"
    worksheet.cell(row=summary_start + 1, column=2).font = BOLD_FONT

    worksheet.cell(row=case_hdr_row, column=2).value = "CASE #"
    worksheet.cell(row=case_hdr_row, column=2).font = BOLD_FONT
    worksheet.cell(row=case_hdr_row, column=4).value = "Dimension In Cms"
    worksheet.cell(row=case_hdr_row, column=4).font = BOLD_FONT

    for offset, item in enumerate(items):
        row = items_start + offset
        desc_val = str(item.get("desc", ""))
        needs_wrap = "\n" in desc_val or len(desc_val) > 36
        worksheet.row_dimensions[row].height = estimate_desc_row_height(desc_val)

        def w(col, value, align):
            c = worksheet.cell(row=row, column=col)
            c.value = value
            c.alignment = align
            c.border = THIN_BORDER

        w(1, item.get("item_code", ""), LEFT)
        w(2, item.get("desc", ""), TOP_LEFT if needs_wrap else LEFT)
        w(3, item.get("case_no", ""), CENTER)
        w(4, item.get("origin", ""), CENTER)
        w(5, item.get("hs_code", ""), CENTER)
        w(6, item.get("qty", ""), CENTER)
        w(7, item.get("gross_weight", ""), RIGHT)
        w(8, item.get("package", ""), CENTER)

    qty_total = round(sum(i["qty"] for i in items if isinstance(i.get("qty"), (int, float))), 2)
    weight_total = round(sum(i["gross_weight"] for i in items if isinstance(i.get("gross_weight"), (int, float))), 2)
    pkg_total = round(sum(i["package"] for i in items if isinstance(i.get("package"), (int, float))), 2)

    c = worksheet.cell(row=total_row, column=5)
    c.value = "Total"
    c.font = BOLD_FONT
    c.alignment = CENTER
    c.border = THIN_BORDER

    for col, val, align in [(6, qty_total, CENTER), (7, weight_total, RIGHT), (8, pkg_total, CENTER)]:
        c = worksheet.cell(row=total_row, column=col)
        c.value = val
        c.font = BOLD_FONT
        c.alignment = align
        c.border = THIN_BORDER

    total_packages = len({i.get("case_no") for i in items if i.get("case_no")})
    worksheet.cell(row=summary_start, column=4).value = total_packages
    worksheet.cell(row=summary_start + 1, column=4).value = weight_total

    for offset, item in enumerate(items):
        row = case_data_start + offset
        c = worksheet.cell(row=row, column=2)
        c.value = item.get("case_no", "")
        c.border = THIN_BORDER
        c.alignment = CENTER
        c = worksheet.cell(row=row, column=4)
        c.value = item.get("dimensions_cm", "")
        c.border = THIN_BORDER
        c.alignment = CENTER

    if items:
        apply_border_to_range(worksheet, case_data_start, case_data_start + len(items) - 1, 2, 4)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def write_dataframe_to_sheet(worksheet, df: pd.DataFrame) -> None:
    worksheet.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        worksheet.append(list(row))


def create_workbook_bytes(
    comm_inv_fields: dict,
    comm_inv_items: list[dict],
    comm_inv_unmatched_items: list[dict],
    pack_list_fields: dict,
    pack_list_items: list[dict],
    comm_inv_df: pd.DataFrame,
    pack_list_df: pd.DataFrame,
) -> io.BytesIO:
    bill_to = comm_inv_fields.get("bill_to", "")
    ship_to = comm_inv_fields.get("ship_to", "")
    address_rows = compute_address_rows(bill_to, ship_to)

    workbook = Workbook()

    comm_sheet = workbook.active
    build_comm_inv_static(comm_sheet)
    fill_comm_inv_items(comm_sheet, comm_inv_items, address_rows)
    fill_comm_inv_sheet(comm_sheet, comm_inv_fields, len(comm_inv_items))
    fill_comm_inv_unmatched_items(comm_sheet, comm_inv_unmatched_items, address_rows, len(comm_inv_items))

    pack_sheet = workbook.create_sheet("pack_list")
    build_pack_list_sheet(pack_sheet)
    fill_pack_list_sheet(pack_sheet, pack_list_fields, address_rows)
    fill_pack_list_items(pack_sheet, pack_list_items, address_rows)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------------------------
# Legacy shim — keeps old callers working
# ---------------------------------------------------------------------------

def get_comm_inv_footer_rows(item_count: int) -> tuple[int, int, int, int]:
    L = get_layout(6, item_count)
    return L["freight_row"], L["total_row"], L["net_total_row"], L["words_row"]