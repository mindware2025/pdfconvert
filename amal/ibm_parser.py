import re

import pdfplumber


CASE_NO_PATTERN = re.compile(r"^(970[A-Z0-9]{10})(.*)$")
ROW_START_PATTERN = re.compile(r"^\d+\s+\S+")
ITEM_ROW_START_PATTERN = re.compile(r"^\d+\s+\S+\s+970[A-Z0-9]{10}")
TAIL_PATTERN = re.compile(
    r"^(?P<body>.+?)(?P<coo>[A-Z]{2})\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<unit_price>[\d,]+(?:\.\d+)?)\s+(?P<total_price>[\d,]+(?:\.\d+)?)$"
)


def normalize_line(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def parse_decimal(value: str) -> float:
    
    return float(value.replace(",", "").strip())


def split_item_and_hs(body: str) -> tuple[str, str, str]:
    # Pattern 1: "4657-924 / 78E3R9W 84717098000000 STORAGE UNITS..."
    match1 = re.match(r"^(.+?\s/\s[A-Z0-9]{7,21})\s+(\d{6,14})\s+(.*)$", body)
    if match1:
        item_code = normalize_line(match1.group(1))
        hs_code = match1.group(2)
        description = normalize_line(match1.group(3))
        return item_code, hs_code, description

    # Pattern 2: "0000003LG589 8523510000 FLASH MEMORY..."
    match2 = re.match(r"^([A-Z0-9]+)\s+(\d{6,12})\s+(.*)$", body)
    if match2:
        item_code = normalize_line(match2.group(1))
        hs_code = match2.group(2)
        description = normalize_line(match2.group(3))
        return item_code, hs_code, description

    # Fallback: no hs code
    return normalize_line(body), "", ""


def parse_item_row(row_text: str) -> dict | None:
    row_text = normalize_line(row_text)
    match = re.match(r"^(?P<line_no>\d+)\s+(?P<order_no>\S+)\s+(?P<rest>.+)$", row_text)
    if not match:
        return None

    rest = match.group("rest")
    case_match = CASE_NO_PATTERN.match(rest)
    if not case_match:
        return None

    case_no = case_match.group(1)
    remaining = normalize_line(case_match.group(2))
    tail_match = TAIL_PATTERN.match(remaining)
    if not tail_match:
        return None

    item_body = normalize_line(tail_match.group("body"))
    item_code, hs_code, description = split_item_and_hs(item_body)
    
    qty = parse_decimal(tail_match.group("qty"))
    total_price = parse_decimal(tail_match.group("total_price"))
    unit_price = total_price / qty if qty else 0

    return {
        "line_no": match.group("line_no"),
        "order_no": match.group("order_no"),
        "case_no": case_no,
        "item_code": item_code,
        "hs_code": hs_code,
        "mibb_description": description,
        "origin": tail_match.group("coo"),
        "qty": qty,
        "temp_unit_price": unit_price,
        "mibb_total_price": total_price,
    }


def normalize_parts_for_value(value: str) -> str:
    normalized = normalize_line(value)
    normalized = normalized.replace("Parts for:", "").strip()
    normalized = re.sub(r"\s*-\s*", "-", normalized)
    normalized = re.sub(r"\s*/\s*", " / ", normalized)
    return normalized


def extract_item_rows_from_ibm_text(ibm_text: str) -> list[dict]:
    lines = [line.strip() for line in ibm_text.splitlines()]
    items: list[tuple[str, str]] = []
    current_row: list[str] = []
    current_parts_for = ""
    pending_parts_for: list[str] | None = None
    in_table = False

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        if "No. Order No Case No Part Number / Serial" in line:
            in_table = True
            continue

        if not in_table:
            continue

        if re.match(r"^[\d,]+(?:\.\d+)?$", line):
            if current_row:
                items.append((" ".join(current_row), current_parts_for))
                current_row = []
                current_parts_for = ""
            break

        if line.startswith("TOTAL AMOUNT") or line.startswith("Volumetric Weight") or line == "TOTAL":
            if current_row:
                items.append((" ".join(current_row), current_parts_for))
                current_row = []
                current_parts_for = ""
            break

        if line.startswith("Parts for:"):
            pending_parts_for = [line.split(":", 1)[1].strip()]
            continue

        if pending_parts_for is not None and not ITEM_ROW_START_PATTERN.match(line):
            pending_parts_for.append(line)
            continue

        if ITEM_ROW_START_PATTERN.match(line):
            if current_row:
                items.append((" ".join(current_row), current_parts_for))
            current_parts_for = normalize_parts_for_value(" ".join(pending_parts_for)) if pending_parts_for else ""
            pending_parts_for = None
            current_row = [line]
        elif current_row:
            current_row.append(line)

    if current_row:
        items.append((" ".join(current_row), current_parts_for))

    parsed_items = []
    for item_text, parts_for_value in items:
        parsed = parse_item_row(item_text)
        if parsed:
            if parts_for_value and "/" not in parsed["item_code"]:
                parsed["original_item_code"] = parsed["item_code"]
                parsed["item_code"] = parts_for_value
                parsed["parts_for_item_code"] = parts_for_value
            parsed_items.append(parsed)

    return parsed_items


def clean_cell(value) -> str:
    if value is None:
        return ""
    return normalize_line(str(value).replace("\n", " "))


def extract_item_rows_from_ibm_pdf(uploaded_file) -> list[dict]:
    uploaded_file.seek(0)
    parsed_items: list[dict] = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table or not table[0]:
                    continue

                header = [clean_cell(cell) for cell in table[0]]
                
                # MIBB format (Part Number / Serial, HS Code columns)
                if "Part Number / Serial" in header and "HS Code" in header:
                    parts_for_value = ""
                    for row in table[1:]:
                        cells = [clean_cell(cell) for cell in row]
                        if len(cells) < 10:
                            continue

                        if cells[0] == "Case No" or cells[0].startswith("Company name"):
                            break

                        if cells[3].startswith("Parts for:"):
                            parts_for_value = normalize_parts_for_value(cells[3])
                            continue

                        if not cells[0].isdigit() or not cells[2]:
                            continue

                        item_code = parts_for_value or cells[3]
                        parts_for_item_code = parts_for_value or ""
                        parts_for_value = ""

                        qty = parse_decimal(cells[7]) if cells[7] else 0
                        total_price = parse_decimal(cells[9]) if cells[9] else 0
                        unit_price = total_price / qty if qty else 0

                        parsed = {
                            "line_no": cells[0],
                            "order_no": cells[1],
                            "case_no": cells[2],
                            "item_code": item_code,
                            "original_item_code": cells[3],
                            "hs_code": cells[4],
                            "mibb_description": cells[5],
                            "origin": cells[6],
                            "qty": qty,
                            "temp_unit_price": unit_price,
                            "mibb_total_price": total_price,
                        }
                        if parts_for_item_code:
                            parsed["parts_for_item_code"] = parts_for_item_code
                        parsed_items.append(parsed)

                    if parsed_items:
                        uploaded_file.seek(0)
                        return parsed_items
                
                # SOB format (Item, Item Description columns)
                elif "Item" in header and "Item Description" in header:
                    item_idx = header.index("Item")
                    desc_idx = header.index("Item Description")
                    qty_idx = header.index("Order Qty") if "Order Qty" in header else -1
                    unit_price_idx = header.index("Unit Price") if "Unit Price" in header else -1
                    total_idx = header.index("Total (excl. VAT)") if "Total (excl. VAT)" in header else -1
                    
                    for row_num, row in enumerate(table[1:], 1):
                        cells = [clean_cell(cell) for cell in row]
                        if not cells or not cells[item_idx]:
                            continue
                        
                        item_code = cells[item_idx]
                        description = cells[desc_idx] if desc_idx < len(cells) else ""
                        qty = parse_decimal(cells[qty_idx]) if qty_idx >= 0 and qty_idx < len(cells) and cells[qty_idx] else 1
                        unit_price = parse_decimal(cells[unit_price_idx]) if unit_price_idx >= 0 and unit_price_idx < len(cells) and cells[unit_price_idx] else 0
                        total_price = parse_decimal(cells[total_idx]) if total_idx >= 0 and total_idx < len(cells) and cells[total_idx] else 0
                        
                        parsed = {
                            "line_no": str(row_num),
                            "order_no": "",
                            "case_no": "",
                            "item_code": item_code,
                            "original_item_code": item_code,
                            "hs_code": "",
                            "mibb_description": description,
                            "origin": "",
                            "qty": qty,
                            "temp_unit_price": unit_price,
                            "mibb_total_price": total_price,
                        }
                        parsed_items.append(parsed)
                    
                    if parsed_items:
                        uploaded_file.seek(0)
                        return parsed_items

    uploaded_file.seek(0)
    return []


def clean_numeric_token(value: str) -> str:
    match = re.search(r"\d+(?:\.\d+)?", value.replace(",", ""))
    return match.group(0) if match else ""


def parse_case_detail_segment(segment: str) -> dict | None:
    tokens = normalize_line(segment).split()
    if len(tokens) < 5 or not CASE_NO_PATTERN.match(tokens[0]):
        return None

    case_no = tokens[0]
    gross_weight = clean_numeric_token(tokens[1])
    if not gross_weight:
        return None

    x_positions = [index for index, token in enumerate(tokens) if token.upper() == "X"]
    if len(x_positions) < 2:
        return None

    first_x = x_positions[0]
    second_x = x_positions[1]
    if first_x - 1 < 0 or first_x + 1 >= len(tokens) or second_x + 1 >= len(tokens):
        return None

    dim_1 = clean_numeric_token(tokens[first_x - 1])
    dim_2 = clean_numeric_token(tokens[first_x + 1])
    dim_3 = clean_numeric_token(tokens[second_x + 1])
    if not all([dim_1, dim_2, dim_3]):
        return None

    return {
        "case_no": case_no,
        "gross_weight": float(gross_weight),
        "dimensions_cm": f"{dim_1} X {dim_2} X {dim_3}",
    }


def extract_case_details_from_ibm_text(ibm_text: str) -> list[dict]:
    normalized_text = re.sub(r"\s+", " ", ibm_text)
    segments = re.findall(r"(970[A-Z0-9]{10}.*?)(?=970[A-Z0-9]{10}|$)", normalized_text)

    details = []
    seen = set()
    for segment in segments:
        parsed = parse_case_detail_segment(segment)
        if not parsed:
            continue
        key = (parsed["case_no"], parsed["gross_weight"], parsed["dimensions_cm"])
        if key in seen:
            continue
        seen.add(key)
        details.append(parsed)

    return details
