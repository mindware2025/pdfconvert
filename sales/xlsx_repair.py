"""Open Excel files that Excel accepts but openpyxl rejects.

Some Dell eQuote exports contain cells with an empty style index (s="").
Microsoft Excel ignores it, but openpyxl fails with
"list indices must be integers or slices, not str". repair_xlsx_bytes()
returns the bytes unchanged when the workbook opens normally; otherwise it
drops the empty style attributes from the sheet XML (cell values are not
touched) and returns the fixed copy.
"""
import re
import zipfile
from io import BytesIO

import openpyxl


def repair_xlsx_bytes(data: bytes) -> bytes:
    if not data or data[:2] != b"PK":
        return data
    try:
        openpyxl.load_workbook(BytesIO(data), data_only=True)
        return data
    except Exception:
        pass
    try:
        src = zipfile.ZipFile(BytesIO(data))
        out = BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                content = src.read(info.filename)
                if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".xml"):
                    content = re.sub(rb'\s+s=""', b"", content)
                dst.writestr(info, content)
        fixed = out.getvalue()
        openpyxl.load_workbook(BytesIO(fixed), data_only=True)
        return fixed
    except Exception:
        return data
