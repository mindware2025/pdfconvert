import xlsxwriter

def col_row_from_a1(a1):
    # simple A1 like "C34" -> (row0, col0)
    col = 0
    for ch in a1:
        if ch.isalpha():
            col = col * 26 + (ord(ch.upper()) - ord('A') + 1)
        else:
            break
    row = int(''.join([c for c in a1 if c.isdigit()])) - 1
    return row, col - 1

def write_terms_rich(worksheet, cell_a1, rich_parts, workbook):
    """
    rich_parts: list of tuples ("text"|"blue", string)
    worksheet: xlsxwriter worksheet
    workbook: xlsxwriter workbook (to create formats)
    """
    # create formats
    blue_fmt = workbook.add_format({'color': '#0000FF'})
    normal_fmt = workbook.add_format({'color': '#000000'})
    # build arguments for write_rich_string
    args = []
    for kind, txt in rich_parts:
        if kind == 'blue':
            args.append(blue_fmt)
            args.append(txt)
        else:
            args.append(normal_fmt)
            args.append(txt)
    # xlsxwriter expects alternating format, text, format, text... but you can start with text or format.
    # Use row/col
    row, col = col_row_from_a1(cell_a1)
    # write_rich_string signature: (row, col, *rich_items, cell_format=None)
    # combine args properly: xlsxwriter can accept format,text,format,text,...
    worksheet.write_rich_string(row, col, *args)