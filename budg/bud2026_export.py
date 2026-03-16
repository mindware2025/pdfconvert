# budg/bud2026_export.py
import io
import xlsxwriter
import pandas as pd
from typing import List, Optional, Sequence, Tuple


def _find_nth_occurrence(cols: List[str], target: str, n: int) -> int:
    """Find the 1-based nth occurrence of 'target' in the list; return its index (0-based)."""
    count = 0
    for i, c in enumerate(cols):
        if c == target:
            count += 1
            if count == n:
                return i
    raise ValueError(f"Header '{target}' occurrence #{n} not found in table header.")


def _compute_banner_positions(
    headers: List[str],
    banner_anchors: Sequence[Tuple[str, str, int]]
) -> List[Tuple[str, int]]:
    """
    For each (banner_title, anchor_header, nth), compute the column index where the banner must be placed.
    """
    positions = []
    for title, anchor, nth in banner_anchors:
        col_idx = _find_nth_occurrence(headers, anchor, nth)
        positions.append((title, col_idx))
    # Sort by column to make merging ranges easier
    positions.sort(key=lambda x: x[1])
    return positions


def _write_banner_row(
    ws,
    positions: List[Tuple[str, int]],
    total_cols: int,
    row_idx: int,
    fmt_banner,
    merge: bool = True
):
    """
    Write banner titles on 'row_idx' at given 'positions'.
    If merge=True, merge each banner title across its group's width
    (from its start col to the col before the next banner).
    """
    if not positions:
        return

    for i, (title, start_col) in enumerate(positions):
        end_col = total_cols - 1
        if i < len(positions) - 1:
            end_col = positions[i + 1][1] - 1
        end_col = max(end_col, start_col)  # safety

        if merge and end_col > start_col:
            ws.merge_range(row_idx, start_col, row_idx, end_col, title, fmt_banner)
        else:
            ws.write(row_idx, start_col, title, fmt_banner)


def build_empty_bud2026_workbook(
    headers: List[str],
    *,
    banner_anchors: Optional[Sequence[Tuple[str, str, int]]] = None,
    header_gap_rows: int = 1,
    freeze: bool = False,
    autofilter: bool = False,
    merge_banner: bool = True
) -> bytes:
    """
    Create an XLSX with:
      Row 1: banner aligned by anchors (optionally merged across group width)
      Row (1 + gap + 1): table header row
    """
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    ws = wb.add_worksheet("ALL")

    fmt_banner = wb.add_format({"bold": True, "align": "center"})
    fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2"})

    row = 0

    # Banner row by anchors
    if banner_anchors:
        positions = _compute_banner_positions(headers, banner_anchors)
        _write_banner_row(ws, positions, len(headers), row, fmt_banner, merge=merge_banner)
        row += 1

    # Gap rows
    for _ in range(header_gap_rows):
        row += 1

    # Table header on this row
    header_row_index = row
    for c, label in enumerate(headers):
        ws.write(header_row_index, c, "" if label is None else str(label), fmt_header)

    # UI niceties
    if freeze:
        ws.freeze_panes(header_row_index + 1, 0)
    if autofilter:
        ws.autofilter(header_row_index, 0, header_row_index, len(headers) - 1)

    wb.close()
    output.seek(0)
    return output.getvalue()


def export_bud2026_ordered(
    df_rows: pd.DataFrame,
    headers: List[str],
    *,
    banner_anchors: Optional[Sequence[Tuple[str, str, int]]] = None,
    header_gap_rows: int = 1,
    freeze: bool = True,
    autofilter: bool = True,
    merge_banner: bool = True
) -> bytes:
    """
    Write rows into 'ALL' with:
      Row 1: banner aligned by anchors
      Row 2..gap: blank
      Row 3: table header
      Row 4+: data
    """
    df = df_rows.copy()
    for col in headers:
        if col not in df.columns:
            df[col] = ""
    df = df[headers]

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"constant_memory": True})
    ws = wb.add_worksheet("ALL")

    fmt_banner = wb.add_format({"bold": True, "align": "center"})
    fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2"})

    row = 0

    # Banner row
    if banner_anchors:
        positions = _compute_banner_positions(headers, banner_anchors)
        _write_banner_row(ws, positions, len(headers), row, fmt_banner, merge=merge_banner)
        row += 1

    # Gap
    for _ in range(header_gap_rows):
        row += 1

    # Table header
    header_row_index = row
    for c, col in enumerate(headers):
        ws.write(header_row_index, c, str(col), fmt_header)

    # Data
    for r, tup in enumerate(df.itertuples(index=False), start=header_row_index + 1):
        ws.write_row(r, 0, list(tup))

    if freeze:
        ws.freeze_panes(header_row_index + 1, 0)
    if autofilter:
        last_data_row = max(header_row_index + 1, header_row_index + len(df))
        ws.autofilter(header_row_index, 0, last_data_row, len(headers) - 1)

    wb.close()
    output.seek(0)
    return output.getvalue()
