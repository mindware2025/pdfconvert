import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


PURPLE_FILL = PatternFill(fill_type="solid", fgColor="3F3D9E")
WHITE_BOLD_FONT = Font(color="FFFFFF", bold=True)
BOLD_FONT = Font(bold=True)
THIN_SIDE = Side(style="thin", color="000000")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
TOP_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
TOP_CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

SUPPLIER_TEXT = (
    "Mindware Fz LLC\n"
    "P.O.Box 55609\n"
    "Jabel Ali, United Arab Emirates\n"
    "Tel : + 971 4500600\n"
    "Email: outboundjebelali@mindware.ae\n"
    "VAT TRN No : 100019912300003"
)


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


def set_outer_border(worksheet, start_row: int, end_row: int, start_col: int, end_col: int) -> None:
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            worksheet.cell(row=row, column=col).border = THIN_BORDER


def get_comm_inv_footer_rows(item_count: int) -> tuple[int, int, int]:
    visible_item_count = max(item_count, 1)
    freight_row = 18 + visible_item_count
    total_row = freight_row + 1
    total_in_words_row = freight_row + 3
    return freight_row, total_row, total_in_words_row


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


def set_merged_block_row_heights(worksheet, start_row: int, end_row: int, total_lines: int) -> None:
    line_count = max(total_lines, end_row - start_row + 1)
    total_height = max(26 * (end_row - start_row + 1), line_count * 24)
    base_height = total_height / (end_row - start_row + 1)
    for row in range(start_row, end_row + 1):
        worksheet.row_dimensions[row].height = base_height


def count_display_lines(value: str) -> int:
    if not value:
        return 1
    return len([line for line in str(value).splitlines() if line.strip()]) or 1


def safe_unmerge(worksheet, cell_range: str) -> None:
    if cell_range in {str(rng) for rng in worksheet.merged_cells.ranges}:
        worksheet.unmerge_cells(cell_range)


def build_comm_inv_sheet(worksheet) -> None:
    worksheet.title = "comm-inv"

    widths = {
        "A": 21,
        "B": 36,
        "C": 16,
        "D": 11,
        "E": 15,
        "F": 14,
        "G": 15,
        "H": 16,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    for row, height in {
        2: 28,
        4: 18,
        8: 18,
        9: 22,
        10: 22,
        11: 22,
        12: 22,
        13: 22,
        14: 22,
        15: 22,
        16: 22,
        17: 30,
    }.items():
        worksheet.row_dimensions[row].height = height

    worksheet.merge_cells("A2:H2")
    worksheet["A2"] = "Commercial Invoice"
    worksheet["A2"].font = Font(bold=True, size=16)
    worksheet["A2"].alignment = CENTER

    worksheet.merge_cells("A4:H4")
    style_range(worksheet, "A4:H4", fill=PURPLE_FILL, border=THIN_BORDER)

    worksheet["A5"] = "Payment Term:"
    worksheet["A6"] = "Inco Terms:"
    worksheet["A7"] = "Customer PO:"
    for cell in ("A5", "A6", "A7"):
        worksheet[cell].font = BOLD_FONT

    worksheet.merge_cells("A5:F5")
    worksheet.merge_cells("A6:F6")
    worksheet.merge_cells("A7:F7")
    set_outer_border(worksheet, 5, 7, 1, 6)

    worksheet.merge_cells("G5:H5")
    worksheet.merge_cells("G6:H6")
    worksheet.merge_cells("G7:H7")
    set_outer_border(worksheet, 5, 7, 7, 8)

    worksheet["G5"] = "Commercial Invoice No :"
    worksheet["G6"] = "Date :"
    worksheet["G7"] = "Currency:"
    for cell in ("G5", "G6", "G7"):
        worksheet[cell].font = BOLD_FONT
        worksheet[cell].alignment = TOP_LEFT

    worksheet["A8"] = "Bill To"
    worksheet["E8"] = "Ship To"
    worksheet["G8"] = "Supplier"
    style_range(worksheet, "A8:H8", fill=PURPLE_FILL, font=WHITE_BOLD_FONT, border=THIN_BORDER)

    worksheet.merge_cells("A8:D8")
    worksheet.merge_cells("E8:F8")
    worksheet.merge_cells("G8:H8")

    worksheet.merge_cells("A9:D16")
    worksheet.merge_cells("E9:F16")
    worksheet.merge_cells("G9:H16")
    set_outer_border(worksheet, 9, 16, 1, 8)

    worksheet["G9"] = SUPPLIER_TEXT
    worksheet["G9"].font = BOLD_FONT
    worksheet["G9"].alignment = TOP_LEFT

    headers = ["Item Code", "Desc", "Case#", "Origin", "HS Code", "Qty", "Unit Price", "Amount"]
    for index, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=17, column=index)
        cell.value = header
        cell.fill = PURPLE_FILL
        cell.font = WHITE_BOLD_FONT
        cell.border = THIN_BORDER
        cell.alignment = LEFT if header in {"Item Code", "Desc"} else CENTER

    set_outer_border(worksheet, 18, 23, 1, 8)

    worksheet.merge_cells("F24:G24")
    worksheet["F24"] = "Freight Charges"
    worksheet["F24"].font = BOLD_FONT
    worksheet["F24"].alignment = CENTER
    set_outer_border(worksheet, 24, 24, 1, 8)

    worksheet.merge_cells("A25:F25")
    worksheet["G25"] = "Total Amount"
    worksheet["G25"].alignment = CENTER
    set_outer_border(worksheet, 25, 25, 1, 8)

    worksheet.merge_cells("A26:H26")
    worksheet["A26"] = "Total in Words :"
    worksheet["A26"].font = BOLD_FONT

    worksheet.merge_cells("A27:F29")
    worksheet.merge_cells("G27:H29")
    set_outer_border(worksheet, 27, 29, 1, 8)

    worksheet["G27"] = "Mindware FZ LLC"
    worksheet["G27"].font = BOLD_FONT
    worksheet["G27"].alignment = CENTER


def fill_comm_inv_sheet(worksheet, fields: dict, item_count: int) -> None:
    freight_row, total_row, total_in_words_row = get_comm_inv_footer_rows(item_count)

    for merge_range in ("F24:G24", "A25:F25", "A26:H26", "A27:F29", "G27:H29"):
        safe_unmerge(worksheet, merge_range)

    for cell_ref in ("F24", "G25", "A26", "G27"):
        worksheet[cell_ref] = None

    worksheet["A5"] = f"Payment Term: {fields.get('payment_term', '')}"
    worksheet["A6"] = f"Inco Terms: {fields.get('inco_terms', '')}"
    worksheet["A7"] = f"Customer PO: {fields.get('customer_po', '')}"
    for cell in ("A5", "A6", "A7"):
        worksheet[cell].font = BOLD_FONT

    worksheet["G5"] = f"Commercial Invoice No : {fields.get('commercial_invoice_no', '')}"
    worksheet["G6"] = f"Date : {fields.get('date', '')}"
    worksheet["G7"] = f"Currency: {fields.get('currency', '')}"
    for cell in ("G5", "G6", "G7"):
        worksheet[cell].alignment = TOP_LEFT

    worksheet["A9"] = fields.get("bill_to", "")
    worksheet["A9"].alignment = TOP_LEFT

    worksheet["E9"] = fields.get("ship_to", "")
    worksheet["E9"].alignment = TOP_LEFT

    address_block_lines = max(
        count_display_lines(fields.get("bill_to", "")),
        count_display_lines(fields.get("ship_to", "")),
        count_display_lines(SUPPLIER_TEXT),
    )
    set_merged_block_row_heights(worksheet, 9, 16, address_block_lines)

    worksheet.merge_cells(start_row=freight_row, start_column=6, end_row=freight_row, end_column=7)
    worksheet.cell(row=freight_row, column=6).value = "Freight Charges"
    worksheet.cell(row=freight_row, column=6).font = BOLD_FONT
    worksheet.cell(row=freight_row, column=6).alignment = CENTER

    worksheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
    worksheet.cell(row=total_row, column=7).value = "Total Amount"
    worksheet.cell(row=total_row, column=7).alignment = CENTER

    worksheet.merge_cells(start_row=total_in_words_row - 1, start_column=1, end_row=total_in_words_row - 1, end_column=8)
    worksheet.cell(row=total_in_words_row - 1, column=1).value = "Total in Words :"
    worksheet.cell(row=total_in_words_row - 1, column=1).font = BOLD_FONT

    worksheet.merge_cells(start_row=total_in_words_row, start_column=1, end_row=total_in_words_row + 2, end_column=6)
    worksheet.merge_cells(start_row=total_in_words_row, start_column=7, end_row=total_in_words_row + 2, end_column=8)
    set_outer_border(worksheet, total_in_words_row, total_in_words_row + 2, 1, 8)
    worksheet.cell(row=total_in_words_row, column=7).value = "Mindware FZ LLC"
    worksheet.cell(row=total_in_words_row, column=7).font = BOLD_FONT
    worksheet.cell(row=total_in_words_row, column=7).alignment = CENTER

    if fields.get("freight_charges", "") != "":
        worksheet.cell(row=freight_row, column=8).value = to_number_if_possible(fields.get("freight_charges", ""))

    item_end_row = 17 + max(item_count, 1)
    worksheet.cell(row=total_row, column=8).value = f"=SUM(H18:H{item_end_row})+H{freight_row}"

    if fields.get("total_in_words", "") != "":
        worksheet.cell(row=total_in_words_row, column=1).value = fields.get("total_in_words", "")
        worksheet.cell(row=total_in_words_row, column=1).alignment = TOP_LEFT


def ensure_comm_inv_item_rows(worksheet, item_count: int) -> None:
    item_start_row = 18
    default_item_rows = 1
    footer_start_row = 24

    if item_count > default_item_rows:
        worksheet.insert_rows(footer_start_row, item_count - default_item_rows)

    item_end_row = item_start_row + max(item_count, default_item_rows) - 1
    set_outer_border(worksheet, item_start_row, item_end_row, 1, 8)


def fill_comm_inv_items(worksheet, items: list[dict]) -> None:
    ensure_comm_inv_item_rows(worksheet, len(items))

    for offset, item in enumerate(items):
        row = 18 + offset
        worksheet.cell(row=row, column=1).value = item.get("item_code", "")
        worksheet.cell(row=row, column=2).value = item.get("desc", "")
        worksheet.cell(row=row, column=2).alignment = TOP_LEFT
        worksheet.row_dimensions[row].height = 34 if "\n" in str(item.get("desc", "")) or len(str(item.get("desc", ""))) > 45 else 22
        worksheet.cell(row=row, column=3).value = item.get("case_no", "")
        worksheet.cell(row=row, column=4).value = item.get("origin", "")
        worksheet.cell(row=row, column=5).value = item.get("hs_code", "")
        worksheet.cell(row=row, column=6).value = item.get("qty", "")
        worksheet.cell(row=row, column=7).value = item.get("unit_price", "")
        worksheet.cell(row=row, column=8).value = item.get("amount", "")


def fill_comm_inv_unmatched_items(worksheet, items: list[dict], item_count: int) -> None:
    if not items:
        return

    _, _, total_in_words_row = get_comm_inv_footer_rows(item_count)
    start_row = total_in_words_row + 5
    worksheet.cell(row=start_row - 1, column=7).value = "Other SOB Items"
    worksheet.cell(row=start_row - 1, column=7).font = BOLD_FONT
    worksheet.cell(row=start_row - 1, column=8).value = "Amount"
    worksheet.cell(row=start_row - 1, column=8).font = BOLD_FONT

    for offset, item in enumerate(items):
        row = start_row + offset
        worksheet.cell(row=row, column=7).value = item.get("item_code", "")
        worksheet.cell(row=row, column=8).value = item.get("amount", "")

    total_row = start_row + len(items)
    worksheet.cell(row=total_row, column=7).value = "Total"
    worksheet.cell(row=total_row, column=7).font = BOLD_FONT
    worksheet.cell(row=total_row, column=8).value = round(
        sum(item["amount"] for item in items if isinstance(item.get("amount"), (int, float))),
        2,
    )
    worksheet.cell(row=total_row, column=8).font = BOLD_FONT

    set_outer_border(worksheet, start_row - 1, total_row, 7, 8)


def build_pack_list_sheet(worksheet) -> None:
    worksheet.title = "pack_list"

    widths = {
        "A": 21,
        "B": 34,
        "C": 15,
        "D": 15,
        "E": 15,
        "F": 14,
        "G": 13,
        "H": 10,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    for row, height in {
        2: 28,
        4: 18,
        7: 18,
        8: 24,
        9: 24,
        10: 24,
        11: 24,
        12: 24,
        13: 24,
        14: 24,
        15: 24,
        16: 28,
    }.items():
        worksheet.row_dimensions[row].height = height

    worksheet.merge_cells("A2:H2")
    worksheet["A2"] = "Packing List"
    worksheet["A2"].font = Font(bold=True, size=16)
    worksheet["A2"].alignment = CENTER

    worksheet.merge_cells("A4:H4")
    style_range(worksheet, "A4:H4", fill=PURPLE_FILL, border=THIN_BORDER)

    worksheet.merge_cells("A5:E6")
    worksheet.merge_cells("F5:H5")
    worksheet.merge_cells("F6:H6")
    set_outer_border(worksheet, 5, 6, 1, 8)

    worksheet["F5"] = "No. :"
    worksheet["F6"] = "Date :"
    worksheet["F5"].font = BOLD_FONT
    worksheet["F6"].font = BOLD_FONT

    worksheet["A7"] = "Bill To"
    worksheet["C7"] = "Ship To"
    worksheet["F7"] = "Supplier"
    style_range(worksheet, "A7:H7", fill=PURPLE_FILL, font=WHITE_BOLD_FONT, border=THIN_BORDER)

    worksheet.merge_cells("A7:B7")
    worksheet.merge_cells("C7:E7")
    worksheet.merge_cells("F7:H7")

    worksheet.merge_cells("A8:B15")
    worksheet.merge_cells("C8:E15")
    worksheet.merge_cells("F8:H15")
    set_outer_border(worksheet, 8, 15, 1, 8)

    worksheet["F8"] = SUPPLIER_TEXT
    worksheet["F8"].font = BOLD_FONT
    worksheet["F8"].alignment = TOP_LEFT

    headers = ["Item Code", "Desc", "Case#", "Origin", "HS Code", "Qty", "Weight", "Package"]
    for index, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=16, column=index)
        cell.value = header
        cell.fill = PURPLE_FILL
        cell.font = WHITE_BOLD_FONT
        cell.border = THIN_BORDER
        cell.alignment = LEFT if header in {"Item Code", "Desc"} else CENTER

    set_outer_border(worksheet, 17, 18, 1, 8)
    worksheet["E19"] = "Total"
    worksheet["E19"].alignment = CENTER
    worksheet["E19"].font = BOLD_FONT
    set_outer_border(worksheet, 19, 19, 6, 8)

    set_outer_border(worksheet, 21, 22, 2, 4)
    worksheet["B21"] = "Total No of Cases"
    worksheet["B22"] = "Total Gross Weight"
    worksheet["B21"].font = BOLD_FONT
    worksheet["B22"].font = BOLD_FONT

    set_outer_border(worksheet, 24, 25, 2, 4)
    worksheet["B24"] = "CASE #"
    worksheet["D24"] = "Dimesion In Cms"
    worksheet["B24"].font = BOLD_FONT
    worksheet["D24"].font = BOLD_FONT

    worksheet["H20"] = "Mindware FZ LLC"
    worksheet["H20"].font = BOLD_FONT
    worksheet["H20"].alignment = CENTER


def ensure_pack_list_rows(worksheet, item_count: int) -> tuple[int, int]:
    default_item_rows = 2
    if item_count > default_item_rows:
        insert_count = item_count - default_item_rows
        worksheet.insert_rows(19, insert_count)

    total_row = 17 + max(item_count, default_item_rows)
    summary_start = total_row + 2
    case_header_row = summary_start + 3
    case_data_start = case_header_row + 1

    set_outer_border(worksheet, 17, total_row - 1, 1, 8)
    set_outer_border(worksheet, total_row, total_row, 6, 8)
    set_outer_border(worksheet, summary_start, summary_start + 1, 2, 4)
    set_outer_border(worksheet, case_header_row, case_header_row, 2, 4)
    return total_row, case_data_start


def fill_pack_list_sheet(worksheet, fields: dict) -> None:
    worksheet["F5"] = '="No. : " & \'comm-inv\'!G5'
    worksheet["F6"] = '="Date : " & \'comm-inv\'!G6'
    worksheet["A8"] = "='comm-inv'!A9"
    worksheet["C8"] = "='comm-inv'!E9"
    worksheet["A8"].alignment = TOP_LEFT
    worksheet["C8"].alignment = TOP_LEFT
    set_merged_block_row_heights(
        worksheet,
        8,
        15,
        max(count_display_lines(fields.get("bill_to", "")), count_display_lines(fields.get("ship_to", "")), count_display_lines(SUPPLIER_TEXT)),
    )
    worksheet["D21"] = fields.get("total_packages", "")
    worksheet["D22"] = fields.get("total_gross_weight", "")


def fill_pack_list_items(worksheet, items: list[dict]) -> None:
    total_row, case_data_start = ensure_pack_list_rows(worksheet, len(items))

    for offset, item in enumerate(items):
        row = 17 + offset
        worksheet.cell(row=row, column=1).value = item.get("item_code", "")
        worksheet.cell(row=row, column=2).value = item.get("desc", "")
        worksheet.cell(row=row, column=2).alignment = TOP_LEFT
        worksheet.row_dimensions[row].height = 34 if "\n" in str(item.get("desc", "")) or len(str(item.get("desc", ""))) > 45 else 22
        worksheet.cell(row=row, column=3).value = item.get("case_no", "")
        worksheet.cell(row=row, column=4).value = item.get("origin", "")
        worksheet.cell(row=row, column=5).value = item.get("hs_code", "")
        worksheet.cell(row=row, column=6).value = item.get("qty", "")
        worksheet.cell(row=row, column=7).value = item.get("gross_weight", "")
        worksheet.cell(row=row, column=8).value = item.get("package", "")

    qty_total = round(
        sum(item["qty"] for item in items if isinstance(item.get("qty"), (int, float))),
        2,
    )
    weight_total = round(
        sum(item["gross_weight"] for item in items if isinstance(item.get("gross_weight"), (int, float))),
        2,
    )
    package_total = round(
        sum(item["package"] for item in items if isinstance(item.get("package"), (int, float))),
        2,
    )

    worksheet.cell(row=total_row, column=6).value = qty_total
    worksheet.cell(row=total_row, column=7).value = weight_total
    worksheet.cell(row=total_row, column=8).value = package_total

    for offset, item in enumerate(items):
        row = case_data_start + offset
        worksheet.cell(row=row, column=2).value = item.get("case_no", "")
        worksheet.cell(row=row, column=4).value = item.get("dimensions_cm", "")

    if items:
        set_outer_border(worksheet, case_data_start, case_data_start + len(items) - 1, 2, 4)


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
    workbook = Workbook()

    comm_sheet = workbook.active
    build_comm_inv_sheet(comm_sheet)
    fill_comm_inv_items(comm_sheet, comm_inv_items)
    fill_comm_inv_sheet(comm_sheet, comm_inv_fields, len(comm_inv_items))
    fill_comm_inv_unmatched_items(comm_sheet, comm_inv_unmatched_items, len(comm_inv_items))

    pack_sheet = workbook.create_sheet("pack_list")
    build_pack_list_sheet(pack_sheet)
    fill_pack_list_sheet(pack_sheet, pack_list_fields)
    fill_pack_list_items(pack_sheet, pack_list_items)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output
