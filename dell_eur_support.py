import re
from typing import Optional, Tuple


EUR_ITEM_RE = re.compile(
    r"^(?P<desc>.*?)(?P<unit>€\s*[\d,]+(?:\.\d+)?)\s+(?P<qty>\d+)\s+(?P<total>€\s*[\d,]+(?:\.\d+)?)$"
)


def _parse_money(value) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    s = re.sub(r"[^\d,.-]", "", s)
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except Exception:
        return None


def is_eur_item_line(line: str) -> bool:
    """Return True only for EUR-style item lines that contain the euro symbol."""
    return bool(line) and ("€" in line or "eur" in line.lower())


def parse_eur_item_line(line: str) -> Optional[Tuple[str, int, float, float]]:
    """Parse a euro-formatted Dell item line into (desc, qty, unit, total)."""
    m = EUR_ITEM_RE.match(line)
    if not m:
        return None

    desc = m.group("desc").strip()
    qty = int(m.group("qty"))
    unit = _parse_money(m.group("unit")) or 0.0
    total = _parse_money(m.group("total")) or 0.0
    return desc, qty, unit, total
