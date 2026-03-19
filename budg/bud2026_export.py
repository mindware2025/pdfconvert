import io

import pandas as pd
import xlsxwriter


def _find_nth_occurrence(cols, target, n=1):
    count = 0
    for i, col in enumerate(cols):
        if col == target:
            count += 1
            if count == n:
                return i
    raise ValueError(f"{target!r} occurrence {n} not found")


def _normalize_header(value: str) -> str:
    return "".join(str(value).split()).lower()


def safe_col(headers, name, occ=1):
    normalized_target = _normalize_header(name)
    matches = []
    for i, header in enumerate(headers):
        if _normalize_header(header) == normalized_target:
            matches.append(i)
    if len(matches) < occ:
        return None
    return xlsxwriter.utility.xl_col_to_name(matches[occ - 1])


def _write_formula_if_present(ws, headers, header_name, row_idx, formula, occ=1):
    try:
        col_idx = _find_nth_occurrence(headers, header_name, occ)
    except ValueError:
        return
    ws.write_formula(row_idx, col_idx, formula)


PERIOD_SPECS = [
    ("Collections FC\n31-03-2026", "AR Provision FC at 31-03-2026"),
    ("Collections FC\n30-06-2026", "AR Provision FC at 30-06-2026"),
    ("Collections FC\n30-09-2026", "AR Provision FC at 30-09-2026"),
    ("Collections FC\n31-12-2026", "AR Provision FC at 31-12-2026"),
    ("Collections FC\n31-12-2027", "AR Provision FC at 31-12-2027"),
    ("Collections FC\n31-12-2028", "AR Provision FC at 31-12-2028"),
]


def export_bud2026_ordered(
    df_rows,
    headers,
    banner_anchors=None,
    header_gap_rows=1,
    freeze=True,
    autofilter=True,
    merge_banner=True,
):
    df = df_rows.copy()
    for col in headers:
        if col not in df.columns:
            df[col] = ""
    df = df[headers]

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    ws = wb.add_worksheet("ALL")

    header_row = max(int(header_gap_rows or 0), 0)
    data_start_row = header_row + 1

    header_fmt = wb.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 1,
            "bg_color": "#D9E2F3",
        }
    )
    banner_fmt = wb.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#B4C6E7",
        }
    )
    text_fmt = wb.add_format({"border": 1})
    num_fmt = wb.add_format({"border": 1, "num_format": "#,##0.00"})

    if merge_banner and banner_anchors:
        anchors = []
        for title, anchor_header, occurrence in banner_anchors:
            try:
                start_idx = _find_nth_occurrence(headers, anchor_header, occurrence)
                anchors.append((title, start_idx))
            except ValueError:
                continue
        for idx, (title, start_idx) in enumerate(anchors):
            end_idx = anchors[idx + 1][1] - 1 if idx + 1 < len(anchors) else len(headers) - 1
            if end_idx < start_idx:
                continue
            if start_idx == end_idx:
                ws.write(0, start_idx, title, banner_fmt)
            else:
                ws.merge_range(0, start_idx, 0, end_idx, title, banner_fmt)

    for col_idx, header in enumerate(headers):
        ws.write(header_row, col_idx, header, header_fmt)

    numeric_headers = {
        "Insurance",
        "On\nAccount",
        "Not Due\nAmount",
        "Aging\n1 to 60",
        "Aging\n61 to 90",
        "Aging\n91 to 120",
        "Aging\n121 to 150",
        "Aging\n>=151",
        " AR\nBalance",
        "AR Provision at\n31-08-2025",
        "AR Provision at\n31-12-2024",
        "Provision without any collection",
        "Provision after collection",
        "Provision after collection including Insurance/BG/LC",
        "Difference in Provision",
        "Collections FC\n31-03-2026",
        "Collections FC\n30-06-2026",
        "Collections FC\n30-09-2026",
        "Collections FC\n31-12-2026",
        "Collections FC\n31-12-2027",
        "Collections FC\n31-12-2028",
    }
    numeric_headers.update({h for h in headers if h == "Expected AR"})
    numeric_headers.update({h for h in headers if h == "Provision Effect"})
    numeric_headers.update({h for h in headers if str(h).startswith("AR Provision FC at ")})

    visible_periods = [
        spec for spec in PERIOD_SPECS if spec[0] in headers and spec[1] in headers
    ]

    # Base column references
    h = safe_col(headers, "Main Ac")
    j = safe_col(headers, "Insurance")
    k = safe_col(headers, "On Account")
    l = safe_col(headers, "Not Due Amount")
    m = safe_col(headers, "Aging 1 to 60")
    n = safe_col(headers, "Aging 61 to 90")
    o = safe_col(headers, "Aging 91 to 120")
    p = safe_col(headers, "Aging 121 to 150")
    q = safe_col(headers, "Aging >=151")
    r = safe_col(headers, "AR Balance")
    s = safe_col(headers, "AR Provision at 31-08-2025")

    for r_idx, row in enumerate(df.itertuples(index=False), start=data_start_row):
        excel_row = r_idx + 1

        for c_idx, value in enumerate(row):
            header = headers[c_idx]
            fmt = num_fmt if header in numeric_headers else text_fmt
            ws.write(r_idx, c_idx, value, fmt)

        if all([h, l, m, n, o, p, q, k]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision without any collection",
                r_idx,
                (
                    f"=IF(IF(IFERROR(VALUE({h}{excel_row}),0)=12301,"
                    f"(({l}{excel_row}*3%)+({m}{excel_row}*((3%*25%)/2))+"
                    f"({n}{excel_row}*50%)+({o}{excel_row}*75%)+"
                    f"{p}{excel_row}+{q}{excel_row}+{k}{excel_row}),0)>0,"
                    f"IF(IFERROR(VALUE({h}{excel_row}),0)=12301,"
                    f"(({l}{excel_row}*3%)+({m}{excel_row}*((3%*25%)/2))+"
                    f"({n}{excel_row}*50%)+({o}{excel_row}*75%)+"
                    f"{p}{excel_row}+{q}{excel_row}+{k}{excel_row}),0),0)"
                ),
            )

        u = safe_col(headers, "Provision without any collection")
        first_collection = safe_col(headers, visible_periods[0][0]) if visible_periods else None
        if all([u, r, first_collection]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision after collection",
                r_idx,
                f"=IFERROR({u}{excel_row}-({u}{excel_row}/{r}{excel_row}*{first_collection}{excel_row}),0)",
            )

        v = safe_col(headers, "Provision after collection")
        if all([j, v]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision after collection including Insurance/BG/LC",
                r_idx,
                (
                    f"=IFERROR(IF({j}{excel_row}>{v}{excel_row},"
                    f"{v}{excel_row}*5%,({j}{excel_row}*5%)+"
                    f"({v}{excel_row}-{j}{excel_row})),0)"
                ),
            )

        w = safe_col(headers, "Provision after collection including Insurance/BG/LC")
        if all([w, s]):
            _write_formula_if_present(
                ws,
                headers,
                "Difference in Provision",
                r_idx,
                f"={w}{excel_row}-{s}{excel_row}",
            )

        x = safe_col(headers, "Difference in Provision")
        prev_expected = None
        prev_ar_provision = None
        for occ, (collection_header, ar_header) in enumerate(visible_periods, start=1):
            collection_col = safe_col(headers, collection_header)
            expected_col = safe_col(headers, "Expected AR", occ)
            effect_col = safe_col(headers, "Provision Effect", occ)
            current_ar_col = safe_col(headers, ar_header)

            if occ == 1:
                if all([r, collection_col]):
                    _write_formula_if_present(
                        ws,
                        headers,
                        "Expected AR",
                        r_idx,
                        f"={r}{excel_row}-{collection_col}{excel_row}",
                        occ=occ,
                    )
                if x:
                    _write_formula_if_present(
                        ws, headers, "Provision Effect", r_idx, f"={x}{excel_row}", occ=occ
                    )
                if all([s, effect_col]):
                    _write_formula_if_present(
                        ws,
                        headers,
                        ar_header,
                        r_idx,
                        f"={s}{excel_row}+{effect_col}{excel_row}",
                    )
            else:
                if all([prev_expected, collection_col]):
                    _write_formula_if_present(
                        ws,
                        headers,
                        "Expected AR",
                        r_idx,
                        f"={prev_expected}{excel_row}-{collection_col}{excel_row}",
                        occ=occ,
                    )
                if all([prev_expected, collection_col, prev_ar_provision]):
                    _write_formula_if_present(
                        ws,
                        headers,
                        "Provision Effect",
                        r_idx,
                        (
                            f"=IF({prev_expected}{excel_row}>0,IFERROR(IF(({collection_col}{excel_row})>{prev_expected}{excel_row},"
                            f"-{prev_ar_provision}{excel_row},((-{collection_col}{excel_row})/{prev_expected}{excel_row}*{prev_ar_provision}{excel_row})),0),0)"
                        ),
                        occ=occ,
                    )
                if all([prev_ar_provision, effect_col]):
                    _write_formula_if_present(
                        ws,
                        headers,
                        ar_header,
                        r_idx,
                        f"={prev_ar_provision}{excel_row}+{effect_col}{excel_row}",
                    )

            prev_expected = expected_col
            prev_ar_provision = current_ar_col

    ws.set_row(header_row, 36)
    if merge_banner and banner_anchors:
        ws.set_row(0, 24)

    for idx, header in enumerate(headers):
        width = 14
        if header in {"Cust Name", "Sales Budget region", "Customer Status"}:
            width = 20
        elif header in {"CustCode", "Main Ac", "Focus List"}:
            width = 12
        elif "\n" in str(header) or str(header).startswith("AR Provision FC at "):
            width = 16
        elif header == "":
            width = 4
        ws.set_column(idx, idx, width)

    if freeze:
        ws.freeze_panes(data_start_row, 0)
    if autofilter:
        ws.autofilter(header_row, 0, max(data_start_row, len(df) + header_row), len(headers) - 1)

    wb.close()
    output.seek(0)
    return output.getvalue()
