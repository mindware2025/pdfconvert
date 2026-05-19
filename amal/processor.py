from dataclasses import dataclass, field
import io
from datetime import datetime

import pandas as pd

try:
    from .ibm_parser import (
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
    from ibm_parser import extract_case_details_from_ibm_text, extract_item_rows_from_ibm_pdf, extract_item_rows_from_ibm_text
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


def process_uploaded_pdfs(sob_file, ibm_file) -> ProcessingResult:
    sob_text = extract_text_from_pdf(sob_file)
    ibm_text = extract_text_from_pdf(ibm_file)
    comm_inv_fields = extract_comm_inv_fields_from_sob(sob_text)
    comm_inv_fields["date"] = datetime.now().strftime("%d/%m/%Y")
    ibm_pdf_items = extract_item_rows_from_ibm_pdf(ibm_file)
    ibm_text_items = extract_item_rows_from_ibm_text(ibm_text)
    ibm_items = merge_ibm_item_sources(ibm_pdf_items, ibm_text_items)
    case_details = extract_case_details_from_ibm_text(ibm_text)
    sob_items = extract_sob_line_items(sob_text)
    comm_inv_items, comm_inv_unmatched_items = map_ibm_items_to_sob(ibm_items, sob_items)
    total_amount = round(
        sum(item["amount"] for item in comm_inv_items if isinstance(item.get("amount"), (int, float))),
        2,
    )
    comm_inv_fields["total_amount"] = total_amount
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

    for pair_index, result in enumerate(pair_results, start=1):
        combined_comm_inv_items.extend(result.comm_inv_items)
        combined_comm_inv_unmatched_items.extend(result.comm_inv_unmatched_items)
        combined_pack_list_items.extend(result.pack_list_items)
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
    comm_inv_fields["total_amount"] = round(
        sum(item["amount"] for item in combined_comm_inv_items if isinstance(item.get("amount"), (int, float))),
        2,
    )
    comm_inv_fields["total_in_words"] = join_distinct_values(
        result.comm_inv_fields.get("total_in_words", "") for result in pair_results
    )

    pack_list_fields = dict(first_result.pack_list_fields)
    pack_list_fields["commercial_invoice_no"] = comm_inv_fields.get("commercial_invoice_no", "")
    pack_list_fields["date"] = comm_inv_fields.get("date", "")
    pack_list_fields["total_packages"] = round(
        sum(item["package"] for item in combined_pack_list_items if isinstance(item.get("package"), (int, float))),
        2,
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
            for result in pair_results
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
    case_occurrences: dict[str, int] = {}
    pack_list_items: list[dict] = []

    for item in comm_inv_items:
        case_no = item.get("case_no", "")
        case_occurrences[case_no] = case_occurrences.get(case_no, 0) + 1
        package_number = case_occurrences[case_no]
        case_detail = case_lookup.get(case_no, {})

        pack_list_items.append(
            {
                "item_code": item.get("item_code", ""),
                "desc": item.get("desc", ""),
                "case_no": case_no,
                "origin": item.get("origin", ""),
                "hs_code": item.get("hs_code", ""),
                "qty": item.get("qty", ""),
                "gross_weight": case_detail.get("gross_weight", ""),
                "package": float(package_number),
                "dimensions_cm": case_detail.get("dimensions_cm", ""),
            }
        )

    total_packages = round(
        sum(item["package"] for item in pack_list_items if isinstance(item.get("package"), (int, float))),
        2,
    )
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

        if text_item.get("item_code"):
            merged_item["item_code"] = text_item["item_code"]
        if text_item.get("parts_for_item_code"):
            merged_item["parts_for_item_code"] = text_item["parts_for_item_code"]
        if text_item.get("mibb_description") and not merged_item.get("mibb_description"):
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
    grouped_items: dict[str, float] = {}
    for item in items:
        item_code = item.get("item_code", "")
        amount = item.get("amount", 0.0)
        if not item_code or not isinstance(amount, (int, float)):
            continue
        grouped_items[item_code] = round(grouped_items.get(item_code, 0.0) + amount, 2)

    return [
        {"item_code": item_code, "amount": amount}
        for item_code, amount in grouped_items.items()
        if amount != 0
    ]
