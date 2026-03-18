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

    z = safe_col(headers, "Collections FC 31-03-2026")
    ae = safe_col(headers, "Collections FC 30-06-2026")
    aj = safe_col(headers, "Collections FC 30-09-2026")
    ao = safe_col(headers, "Collections FC 31-12-2026")
    at = safe_col(headers, "Collections FC 31-12-2027")
    ay = safe_col(headers, "Collections FC 31-12-2028")

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
                    f"=IF(IF({h}{excel_row}=12301,"
                    f"(({l}{excel_row}*3%)+({m}{excel_row}*((3%*25%)/2))+"
                    f"({n}{excel_row}*50%)+({o}{excel_row}*75%)+"
                    f"{p}{excel_row}+{q}{excel_row}+{k}{excel_row}),0)>0,"
                    f"IF({h}{excel_row}=12301,"
                    f"(({l}{excel_row}*3%)+({m}{excel_row}*((3%*25%)/2))+"
                    f"({n}{excel_row}*50%)+({o}{excel_row}*75%)+"
                    f"{p}{excel_row}+{q}{excel_row}+{k}{excel_row}),0),0)"
                ),
            )

        u = safe_col(headers, "Provision without any collection")
        if all([u, r, z]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision after collection",
                r_idx,
                f"=IFERROR({u}{excel_row}-({u}{excel_row}/{r}{excel_row}*{z}{excel_row}),0)",
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
        ac = safe_col(headers, "AR Provision FC at 31-03-2026")
        ah = safe_col(headers, "AR Provision FC at 30-06-2026")
        am = safe_col(headers, "AR Provision FC at 30-09-2026")
        ar = safe_col(headers, "AR Provision FC at 31-12-2026")
        aw = safe_col(headers, "AR Provision FC at 31-12-2027")

        aa = safe_col(headers, "Expected AR", 1)
        ab = safe_col(headers, "Provision Effect", 1)
        af = safe_col(headers, "Expected AR", 2)
        ag = safe_col(headers, "Provision Effect", 2)
        ak = safe_col(headers, "Expected AR", 3)
        al = safe_col(headers, "Provision Effect", 3)
        ap = safe_col(headers, "Expected AR", 4)
        aq = safe_col(headers, "Provision Effect", 4)
        au = safe_col(headers, "Expected AR", 5)
        av = safe_col(headers, "Provision Effect", 5)
        az = safe_col(headers, "Expected AR", 6)
        ba = safe_col(headers, "Provision Effect", 6)
        bb = safe_col(headers, "AR Provision FC at 31-12-2028")

        if all([r, z]):
            _write_formula_if_present(ws, headers, "Expected AR", r_idx, f"={r}{excel_row}-{z}{excel_row}", occ=1)
        if x:
            _write_formula_if_present(ws, headers, "Provision Effect", r_idx, f"={x}{excel_row}", occ=1)
        if all([s, ab]):
            _write_formula_if_present(
                ws, headers, "AR Provision FC at 31-03-2026", r_idx, f"={s}{excel_row}+{ab}{excel_row}"
            )

        if all([aa, ae]):
            _write_formula_if_present(ws, headers, "Expected AR", r_idx, f"={aa}{excel_row}-{ae}{excel_row}", occ=2)
        if all([aa, ae, ac]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision Effect",
                r_idx,
                (
                    f"=IF({aa}{excel_row}>0,IFERROR(IF(({ae}{excel_row})>{aa}{excel_row},"
                    f"-{ac}{excel_row},((-{ae}{excel_row})/{aa}{excel_row}*{ac}{excel_row})),0),0)"
                ),
                occ=2,
            )
        if all([ac, ag]):
            _write_formula_if_present(
                ws, headers, "AR Provision FC at 30-06-2026", r_idx, f"={ac}{excel_row}+{ag}{excel_row}"
            )

        if all([af, aj]):
            _write_formula_if_present(ws, headers, "Expected AR", r_idx, f"={af}{excel_row}-{aj}{excel_row}", occ=3)
        if all([ak, aj, ah, af]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision Effect",
                r_idx,
                (
                    f"=IF({ak}{excel_row}>0,IFERROR(IF(({aj}{excel_row})>{ak}{excel_row},"
                    f"-{ah}{excel_row},((-{aj}{excel_row})/{af}{excel_row}*{ah}{excel_row})),0),0)"
                ),
                occ=3,
            )
        if all([ah, al]):
            _write_formula_if_present(
                ws, headers, "AR Provision FC at 30-09-2026", r_idx, f"={ah}{excel_row}+{al}{excel_row}"
            )

        if all([ak, ao]):
            _write_formula_if_present(ws, headers, "Expected AR", r_idx, f"={ak}{excel_row}-{ao}{excel_row}", occ=4)
        if all([ap, ao, am, ak]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision Effect",
                r_idx,
                (
                    f"=IF({ap}{excel_row}>0,IFERROR(IF(({ao}{excel_row})>{ap}{excel_row},"
                    f"-{am}{excel_row},((-{ao}{excel_row})/{ak}{excel_row}*{am}{excel_row})),0),0)"
                ),
                occ=4,
            )
        if all([am, aq]):
            _write_formula_if_present(
                ws, headers, "AR Provision FC at 31-12-2026", r_idx, f"={am}{excel_row}+{aq}{excel_row}"
            )

        if all([ap, at]):
            _write_formula_if_present(ws, headers, "Expected AR", r_idx, f"={ap}{excel_row}-{at}{excel_row}", occ=5)
        if all([au, at, ar, ap]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision Effect",
                r_idx,
                (
                    f"=IF({au}{excel_row}>0,IFERROR(IF(({at}{excel_row})>{au}{excel_row},"
                    f"-{ar}{excel_row},((-{at}{excel_row})/{ap}{excel_row}*{ar}{excel_row})),0),0)"
                ),
                occ=5,
            )
        if all([ar, av]):
            _write_formula_if_present(
                ws, headers, "AR Provision FC at 31-12-2027", r_idx, f"={ar}{excel_row}+{av}{excel_row}"
            )

        if all([au, ay]):
            _write_formula_if_present(ws, headers, "Expected AR", r_idx, f"={au}{excel_row}-{ay}{excel_row}", occ=6)
        if all([az, ay, aw, au]):
            _write_formula_if_present(
                ws,
                headers,
                "Provision Effect",
                r_idx,
                (
                    f"=IF({az}{excel_row}>0,IFERROR(IF(({ay}{excel_row})>{az}{excel_row},"
                    f"-{aw}{excel_row},((-{ay}{excel_row})/{au}{excel_row}*{aw}{excel_row})),0),0)"
                ),
                occ=6,
            )
        if all([aw, ba]):
            _write_formula_if_present(
                ws, headers, "AR Provision FC at 31-12-2028", r_idx, f"={aw}{excel_row}+{ba}{excel_row}"
            )

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
