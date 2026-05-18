import re


def normalize_whitespace(value: str) -> str:
    return re.sub(r"[ \t]+", " ", value).strip()


def extract_inline_value(text: str, label: str, next_labels: list[str] | None = None) -> str:
    if next_labels:
        next_pattern = "|".join(re.escape(next_label) for next_label in next_labels)
        pattern = rf"{re.escape(label)}\s*:\s*(.*?)(?=(?:{next_pattern})\s*:|$)"
    else:
        pattern = rf"{re.escape(label)}\s*:\s*(.+)"

    match = re.search(pattern, text, flags=re.S)
    if not match:
        return ""
    return normalize_whitespace(match.group(1))


def extract_block(text: str, start_marker: str, end_marker: str) -> str:
    pattern = rf"{re.escape(start_marker)}\s*(.*?)(?={re.escape(end_marker)})"
    match = re.search(pattern, text, flags=re.S)
    if not match:
        return ""

    lines = [line.strip() for line in match.group(1).splitlines()]
    clean_lines = [line for line in lines if line]
    return "\n".join(clean_lines).strip()


def split_bill_to_ship_to(block: str) -> tuple[str, str]:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    if not lines:
        return "", ""

    separator_index = None
    for index, line in enumerate(lines):
        if "GROUPEMENT INTERBANCAIRE" in line.upper():
            separator_index = index
            break

    if separator_index is None:
        midpoint = max(1, len(lines) // 2)
        return "\n".join(lines[:midpoint]), "\n".join(lines[midpoint:])

    bill_to_lines = lines[:separator_index]
    ship_to_lines = lines[separator_index:]
    if ship_to_lines and "GROUPEMENT INTERBANCAIRE" in ship_to_lines[0].upper():
        cleaned_first_line = re.sub(r"^.*?(GROUPEMENT INTERBANCAIRE)", r"\1", ship_to_lines[0], flags=re.I)
        ship_to_lines[0] = cleaned_first_line

    return "\n".join(bill_to_lines).strip(), "\n".join(ship_to_lines).strip()


def extract_comm_inv_fields_from_sob(sob_text: str) -> dict:
    bill_ship_block = extract_block(sob_text, "Bill To Ship To", "Forwarder")
    bill_to, ship_to = split_bill_to_ship_to(bill_ship_block)

    return {
        "payment_term": extract_inline_value(sob_text, "Credit Terms", ["Ship Via"]),
        "inco_terms": extract_inline_value(sob_text, "Inco Terms", ["Currency"]),
        "customer_po": extract_inline_value(sob_text, "Customer PO", ["Remarks"]),
        "commercial_invoice_no": extract_inline_value(sob_text, "Order No", ["Order Date"]),
        "currency": extract_inline_value(sob_text, "Currency", ["Customer PO"]),
        "freight_charges": extract_inline_value(sob_text, "Freight Charges ®", ["VAT"]),
        "total_in_words": extract_inline_value(sob_text, "Amount in Words", ["Bank Details"]),
        "bill_to": bill_to,
        "ship_to": ship_to,
    }


SOB_TAIL_PATTERN = re.compile(
    r"^(?P<body>.+?)(?P<del_loc>[A-Z]{2}\d{3})\s+(?P<uom>\S+)\s+(?P<qty>[\d,]+(?:\.\d+)?)\s+"
    r"(?P<unit_price>[\d,]+(?:\.\d+)?)\s+(?P<vat_pct>[\d,]+(?:\.\d+)?)\s+"
    r"(?P<vat>[\d,]+(?:\.\d+)?)\s+(?P<total>[\d,]+(?:\.\d+)?)$"
)


def parse_decimal(value: str) -> float:
    return float(value.replace(",", "").strip())


def normalize_item_code(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", value.upper())


def get_group_code(item_code: str) -> str:
    if item_code.upper().startswith("HS-IBM"):
        return "HS-IBM"

    if "-" in item_code:
        return item_code.split("-", 1)[0]

    if item_code.upper().endswith("IBM"):
        return item_code[:-3]

    return item_code


def extract_sob_line_items(sob_text: str) -> list[dict]:
    lines = [line.strip() for line in sob_text.splitlines()]
    row_chunks: list[str] = []
    current_row: list[str] = []
    in_table = False

    for line in lines:
        if not line:
            continue

        if line.startswith("Sl.No Item Item Description"):
            in_table = True
            continue

        if not in_table:
            continue

        if line.startswith("Gross Total") or line.startswith("Freight Charges"):
            if current_row:
                row_chunks.append(" ".join(current_row))
            break

        if re.match(r"^\d+\s+", line):
            if current_row:
                row_chunks.append(" ".join(current_row))
            current_row = [line]
        elif current_row and not line.startswith("Page "):
            current_row.append(line)

    parsed_rows = []
    for chunk in row_chunks:
        parsed = parse_sob_line_item(chunk)
        if parsed:
            parsed_rows.append(parsed)

    return parsed_rows


def parse_sob_line_item(row_text: str) -> dict | None:
    row_text = normalize_whitespace(row_text)
    match = re.match(r"^(?P<line_no>\d+)\s+(?P<rest>.+)$", row_text)
    if not match:
        return None

    tail_match = SOB_TAIL_PATTERN.match(match.group("rest"))
    if not tail_match:
        return None

    body = tail_match.group("body")
    body_match = re.match(r"^(?P<item_code>[A-Z0-9-]+)(?P<description>.*)$", body)
    if not body_match:
        return None

    item_code = body_match.group("item_code").strip()
    description = normalize_whitespace(body_match.group("description"))

    return {
        "line_no": match.group("line_no"),
        "item_code": item_code,
        "normalized_item_code": normalize_item_code(item_code),
        "description": description,
        "qty": parse_decimal(tail_match.group("qty")),
        "total": parse_decimal(tail_match.group("total")),
    }


def map_ibm_items_to_sob(ibm_items: list[dict], sob_items: list[dict]) -> tuple[list[dict], list[dict]]:
    mapped_items = []
    matched_sob_codes: set[str] = set()

    for ibm_item in ibm_items:
        base_code = ibm_item.get("item_code", "").split("/")[0].strip()
        base_normalized = normalize_item_code(base_code)
        prefix_matches = [
            sob_item
            for sob_item in sob_items
            if sob_item["normalized_item_code"].startswith(base_normalized)
        ]
        exact_match = next(
            (sob_item for sob_item in sob_items if sob_item["normalized_item_code"] == base_normalized),
            None,
        )

        description = ""
        if exact_match:
            description = exact_match["description"]
        elif prefix_matches:
            description = prefix_matches[0]["description"]

        is_parts_for_item = bool(ibm_item.get("parts_for_item_code"))
        if is_parts_for_item:
            amount = 0.0
            description = ibm_item.get("mibb_description", "")
        else:
            amount = round(sum(sob_item["total"] for sob_item in prefix_matches), 2)
            for sob_item in prefix_matches:
                matched_sob_codes.add(sob_item["normalized_item_code"])

        mapped_item = dict(ibm_item)
        mapped_item["desc"] = description
        mapped_item["amount"] = amount if (prefix_matches or is_parts_for_item) else ""
        if (prefix_matches or is_parts_for_item) and ibm_item.get("qty"):
            mapped_item["unit_price"] = round(amount / ibm_item["qty"], 2)
        else:
            mapped_item["unit_price"] = ""
        mapped_items.append(mapped_item)

    unmatched_groups: dict[str, dict] = {}
    for sob_item in sob_items:
        normalized_code = sob_item["normalized_item_code"]
        if normalized_code in matched_sob_codes:
            continue

        group_code = get_group_code(sob_item["item_code"])
        if group_code not in unmatched_groups:
            unmatched_groups[group_code] = {
                "item_code": group_code,
                "amount": 0.0,
            }
        unmatched_groups[group_code]["amount"] = round(
            unmatched_groups[group_code]["amount"] + sob_item["total"],
            2,
        )

    unmatched_items = [
        value for value in unmatched_groups.values() if value["amount"] != 0
    ]

    return mapped_items, unmatched_items
