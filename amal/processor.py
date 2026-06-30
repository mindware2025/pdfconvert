from dataclasses import dataclass, field
import io
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from .ibm_parser import (
        extract_case_details_from_ibm_pdf,
        extract_case_details_from_ibm_text,
        extract_item_rows_from_ibm_pdf,
        extract_item_rows_from_ibm_text,
    )
    from .pdf_utils import extract_text_from_pdf
    from .sob_parser import (
        extract_comm_inv_fields_from_sob,
        extract_sob_line_items,
        map_ibm_items_to_sob,
    )
    from .workbook_builder import create_workbook_bytes
except ImportError:
    from ibm_parser import extract_case_details_from_ibm_pdf, extract_case_details_from_ibm_text, extract_item_rows_from_ibm_pdf, extract_item_rows_from_ibm_text
    from pdf_utils import extract_text_from_pdf
    from sob_parser import extract_comm_inv_fields_from_sob, extract_sob_line_items, map_ibm_items_to_sob
    from workbook_builder import create_workbook_bytes


@dataclass
class ProcessingResult:
    sob_filename: str
    ibm_filename: str
    sob_text: str
    ibm_text: str
    comm_inv_fields: dict
    comm_inv_items: list[dict]
    comm_inv_unmatched_items: list[dict]
    pack_list_fields: dict
    pack_list_items: list[dict]
    comm_inv_df: pd.DataFrame
    pack_list_df: pd.DataFrame
    messages: list[str] = field(default_factory=list)


def normalize_file_identifier(file_name: str) -> str:
    return Path(str(file_name)).name.strip().lower()


def parse_decimal_string(value) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return 0.0

    candidate = value.replace(",", "").strip()
    if not candidate:
        return 0.0

    try:
        return float(candidate)
    except ValueError:
        return 0.0


ONES = [
    "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
]
TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
SCALES = [(1_000_000_000, "Billion"), (1_000_000, "Million"), (1_000, "Thousand"), (100, "Hundred")]


def number_to_words(value: int) -> str:
    if value < 20:
        return ONES[value]
    if value < 100:
        tens, remainder = divmod(value, 10)
        return TENS[tens] if remainder == 0 else f"{TENS[tens]} {ONES[remainder]}"

    for scale_value, scale_name in SCALES:
        if value >= scale_value:
            major, remainder = divmod(value, scale_value)
            major_words = number_to_words(major)
            return f"{major_words} {scale_name}" if remainder == 0 else f"{major_words} {scale_name} {number_to_words(remainder)}"

    return str(value)


def amount_to_words(amount: float, currency: str = "USD") -> str:
    normalized_amount = round(float(amount), 2)
    whole_part = int(normalized_amount)
    cents = int(round((normalized_amount - whole_part) * 100))
    if cents == 100:
        whole_part += 1
        cents = 0

    whole_words = number_to_words(whole_part)
    cents_words = number_to_words(cents)
    currency_code = (currency or "USD").strip().upper()
    return f"{currency_code} {whole_words} And Cents {cents_words} Only"


def pick_shared_value(values) -> str:
    distinct_values = []
    seen: set[str] = set()
    for value in values:
        clean_value = str(value).strip()
        if not clean_value:
            continue
        key = clean_value.upper()
        if key in seen:
            continue
        seen.add(key)
        distinct_values.append(clean_value)

    if len(distinct_values) == 1:
        return distinct_values[0]
    return ""


def process_uploaded_pdfs(sob_file, ibm_file) -> ProcessingResult:
    sob_text = extract_text_from_pdf(sob_file)
    ibm_text = extract_text_from_pdf(ibm_file)
    comm_inv_fields = extract_comm_inv_fields_from_sob(sob_text)
    comm_inv_fields["date"] = datetime.now().strftime("%d/%m/%Y")
    ibm_pdf_items = extract_item_rows_from_ibm_pdf(ibm_file)
    ibm_text_items = extract_item_rows_from_ibm_text(ibm_text)
    ibm_items = merge_ibm_item_sources(ibm_pdf_items, ibm_text_items)
    case_details = extract_case_details_from_ibm_pdf(ibm_file) or extract_case_details_from_ibm_text(ibm_text)
    sob_items = extract_sob_line_items(sob_text)
    comm_inv_items, comm_inv_unmatched_items = map_ibm_items_to_sob(ibm_items, sob_items)
    sob_reference = Path(str(sob_file.name)).stem
    for item in comm_inv_unmatched_items:
        item["sob_reference"] = sob_reference
    total_amount = round(
        sum(item["amount"] for item in comm_inv_items if isinstance(item.get("amount"), (int, float))),
        2,
    )
    comm_inv_fields["total_amount"] = total_amount
    comm_inv_fields["sob_total"] = parse_decimal_string(comm_inv_fields.get("sob_total", ""))
    if not comm_inv_fields.get("total_in_words") and comm_inv_fields["sob_total"]:
        comm_inv_fields["total_in_words"] = amount_to_words(
            comm_inv_fields["sob_total"],
            comm_inv_fields.get("currency", "USD"),
        )
    pack_list_fields, pack_list_items = build_pack_list_data(comm_inv_fields, comm_inv_items, case_details)

    messages = [
        f"SOB file loaded: {sob_file.name}",
        f"IBM file loaded: {ibm_file.name}",
        "Commercial invoice header fields are mapped from the SOB reference.",
        f"Commercial invoice item rows found in IBM file: {len(comm_inv_items)}",
        "Commercial invoice descriptions and amounts are mapped from SOB by normalized item prefix.",
        f"Commercial invoice unmatched SOB groups: {len(comm_inv_unmatched_items)}",
        f"Pack list case details found in IBM file: {len(case_details)}",
    ]

    comm_inv_df = pd.DataFrame(
        [
            {
                "source_file": ibm_file.name,
                "document_type": "commercial_invoice",
                "status": "pending_structure",
            }
        ]
    )

    pack_list_df = pd.DataFrame(
        [
            {
                "source_file": sob_file.name,
                "document_type": "sob",
                "status": "pending_structure",
            }
        ]
    )

    return ProcessingResult(
        sob_filename=sob_file.name,
        ibm_filename=ibm_file.name,
        sob_text=sob_text,
        ibm_text=ibm_text,
        comm_inv_fields=comm_inv_fields,
        comm_inv_items=comm_inv_items,
        comm_inv_unmatched_items=comm_inv_unmatched_items,
        pack_list_fields=pack_list_fields,
        pack_list_items=pack_list_items,
        comm_inv_df=comm_inv_df,
        pack_list_df=pack_list_df,
        messages=messages,
    )


def process_uploaded_pairs(file_pairs: list[tuple]) -> ProcessingResult:
    pair_results = [process_uploaded_pdfs(sob_file, ibm_file) for sob_file, ibm_file in file_pairs]
    if not pair_results:
        raise ValueError("At least one complete SOB + IBM pair is required.")

    first_result = pair_results[0]
    combined_comm_inv_items: list[dict] = []
    combined_comm_inv_unmatched_items: list[dict] = []
    combined_pack_list_items: list[dict] = []
    combined_messages: list[str] = []
    seen_ibm_files: set[str] = set()

    for pair_index, result in enumerate(pair_results, start=1):
        ibm_identifier = normalize_file_identifier(result.ibm_filename)
        if ibm_identifier not in seen_ibm_files:
            combined_comm_inv_items.extend(result.comm_inv_items)
            combined_pack_list_items.extend(result.pack_list_items)
            seen_ibm_files.add(ibm_identifier)

        combined_comm_inv_unmatched_items.extend(result.comm_inv_unmatched_items)
        combined_messages.append(
            f"Pair {pair_index}: {result.sob_filename} + {result.ibm_filename}"
        )

    comm_inv_fields = dict(first_result.comm_inv_fields)
    comm_inv_fields["date"] = datetime.now().strftime("%d/%m/%Y")
    comm_inv_fields["customer_po"] = join_distinct_values(
        result.comm_inv_fields.get("customer_po", "") for result in pair_results
    )
    comm_inv_fields["commercial_invoice_no"] = join_distinct_values(
        result.comm_inv_fields.get("commercial_invoice_no", "") for result in pair_results
    )
    comm_inv_fields["freight_charges"] = round(
        sum(parse_decimal_string(result.comm_inv_fields.get("freight_charges", "")) for result in pair_results),
        2,
    )
    summed_sob_total = round(
        sum(parse_decimal_string(result.comm_inv_fields.get("sob_total", "")) for result in pair_results),
        2,
    )
    comm_inv_fields["sob_total"] = summed_sob_total
    comm_inv_fields["total_amount"] = round(
        sum(item["amount"] for item in combined_comm_inv_items if isinstance(item.get("amount"), (int, float))),
        2,
    )
    comm_inv_fields["total_in_words"] = amount_to_words(
        summed_sob_total,
        comm_inv_fields.get("currency", "USD"),
    ) if summed_sob_total else pick_shared_value(
        result.comm_inv_fields.get("total_in_words", "") for result in pair_results
    )

    pack_list_fields = dict(first_result.pack_list_fields)
    pack_list_fields["commercial_invoice_no"] = comm_inv_fields.get("commercial_invoice_no", "")
    pack_list_fields["date"] = comm_inv_fields.get("date", "")
    pack_list_fields["total_packages"] = len(
        {item["case_no"] for item in combined_pack_list_items if item.get("case_no")}
    )
    pack_list_fields["total_gross_weight"] = round(
        sum(item["gross_weight"] for item in combined_pack_list_items if isinstance(item.get("gross_weight"), (int, float))),
        2,
    )
    pack_list_fields["total_qty"] = round(
        sum(item["qty"] for item in combined_pack_list_items if isinstance(item.get("qty"), (int, float))),
        2,
    )

    comm_inv_df = pd.DataFrame(
        [
            {
                "source_file": result.ibm_filename,
                "document_type": "commercial_invoice",
                "status": "pending_structure",
            }
            for index, result in enumerate(pair_results)
            if normalize_file_identifier(result.ibm_filename)
            not in {
                normalize_file_identifier(previous.ibm_filename)
                for previous in pair_results[:index]
            }
        ]
    )

    pack_list_df = pd.DataFrame(
        [
            {
                "source_file": result.sob_filename,
                "document_type": "sob",
                "status": "pending_structure",
            }
            for result in pair_results
        ]
    )

    return ProcessingResult(
        sob_filename=", ".join(result.sob_filename for result in pair_results),
        ibm_filename=", ".join(result.ibm_filename for result in pair_results),
        sob_text="\n\n".join(result.sob_text for result in pair_results if result.sob_text),
        ibm_text="\n\n".join(result.ibm_text for result in pair_results if result.ibm_text),
        comm_inv_fields=comm_inv_fields,
        comm_inv_items=combined_comm_inv_items,
        comm_inv_unmatched_items=merge_unmatched_items(combined_comm_inv_unmatched_items),
        pack_list_fields=pack_list_fields,
        pack_list_items=combined_pack_list_items,
        comm_inv_df=comm_inv_df,
        pack_list_df=pack_list_df,
        messages=combined_messages,
    )


def build_output_workbook(result: ProcessingResult) -> io.BytesIO:
    return create_workbook_bytes(
        comm_inv_fields=result.comm_inv_fields,
        comm_inv_items=result.comm_inv_items,
        comm_inv_unmatched_items=result.comm_inv_unmatched_items,
        pack_list_fields=result.pack_list_fields,
        pack_list_items=result.pack_list_items,
        comm_inv_df=result.comm_inv_df,
        pack_list_df=result.pack_list_df,
    )


def build_pack_list_data(comm_inv_fields: dict, comm_inv_items: list[dict], case_details: list[dict]) -> tuple[dict, list[dict]]:
    case_lookup = {detail["case_no"]: detail for detail in case_details}
    case_item_counts: dict[str, int] = {}
    for item in comm_inv_items:
        case_no = item.get("case_no", "")
        case_item_counts[case_no] = case_item_counts.get(case_no, 0) + 1

    pack_list_items: list[dict] = []

    for item in comm_inv_items:
        case_no = item.get("case_no", "")
        case_detail = case_lookup.get(case_no, {})
        raw_weight = case_detail.get("gross_weight", "")
        count_in_case = case_item_counts.get(case_no, 1)
        if isinstance(raw_weight, (int, float)) and raw_weight:
            weight_per_item = round(raw_weight / count_in_case, 2)
        else:
            weight_per_item = raw_weight

        pack_list_items.append(
            {
                "item_code": item.get("item_code", ""),
                "desc": item.get("desc", ""),
                "case_no": case_no,
                "origin": item.get("origin", ""),
                "hs_code": item.get("hs_code", ""),
                "qty": item.get("qty", ""),
                "gross_weight": weight_per_item,
                "package": 1.0,
                "dimensions_cm": case_detail.get("dimensions_cm", ""),
            }
        )

    total_packages = len({item["case_no"] for item in pack_list_items if item.get("case_no")})
    total_gross_weight = round(
        sum(item["gross_weight"] for item in pack_list_items if isinstance(item.get("gross_weight"), (int, float))),
        2,
    )
    total_qty = round(
        sum(item["qty"] for item in pack_list_items if isinstance(item.get("qty"), (int, float))),
        2,
    )

    pack_list_fields = {
        "commercial_invoice_no": comm_inv_fields.get("commercial_invoice_no", ""),
        "date": comm_inv_fields.get("date", ""),
        "bill_to": comm_inv_fields.get("bill_to", ""),
        "ship_to": comm_inv_fields.get("ship_to", ""),
        "total_packages": total_packages,
        "total_gross_weight": total_gross_weight,
        "total_qty": total_qty,
    }

    return pack_list_fields, pack_list_items


def merge_ibm_item_sources(pdf_items: list[dict], text_items: list[dict]) -> list[dict]:
    if not pdf_items:
        return text_items
    if not text_items:
        return pdf_items

    text_lookup = {
        (item.get("line_no", ""), item.get("case_no", ""), item.get("order_no", "")): item
        for item in text_items
    }

    merged_items: list[dict] = []
    for pdf_item in pdf_items:
        key = (pdf_item.get("line_no", ""), pdf_item.get("case_no", ""), pdf_item.get("order_no", ""))
        text_item = text_lookup.get(key, {})
        merged_item = dict(pdf_item)

        has_confident_text_split = bool(text_item.get("hs_code") or text_item.get("mibb_description"))
        if text_item.get("item_code") and has_confident_text_split:
            merged_item["item_code"] = text_item["item_code"]
        if text_item.get("parts_for_item_code"):
            merged_item["parts_for_item_code"] = text_item["parts_for_item_code"]
        if text_item.get("mibb_description"):
            merged_item["mibb_description"] = text_item["mibb_description"]

        merged_items.append(merged_item)

    return merged_items


def join_distinct_values(values) -> str:
    seen: set[str] = set()
    ordered_values: list[str] = []
    for value in values:
        clean_value = str(value).strip()
        if not clean_value:
            continue
        normalized_value = clean_value.upper()
        if normalized_value in seen:
            continue
        seen.add(normalized_value)
        ordered_values.append(clean_value)
    return " & ".join(ordered_values)


def merge_unmatched_items(items: list[dict]) -> list[dict]:
    grouped_items: dict[tuple[str, str], float] = {}
    for item in items:
        sob_reference = item.get("sob_reference", "")
        item_code = item.get("item_code", "")
        amount = item.get("amount", 0.0)
        if not item_code or not isinstance(amount, (int, float)):
            continue
        key = (sob_reference, item_code)
        grouped_items[key] = round(grouped_items.get(key, 0.0) + amount, 2)

    return [
        {"sob_reference": sob_reference, "item_code": item_code, "amount": amount}
        for (sob_reference, item_code), amount in grouped_items.items()
        if amount != 0
    ]
