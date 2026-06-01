import re

def _extract_part_number_from_description(text: str) -> str:
    """Extract common Dell part numbers from parentheses in a description.

    Returns the last matching candidate found, or empty string if none.
    """
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""

    matches = re.findall(r"\(([^()]+)\)", s)
    for candidate in reversed(matches):
        normalized = candidate.strip()
        normalized = normalized.replace("–", "-").replace("—", "-").replace("−", "-")
        normalized = re.sub(r"\s*-\s*", "-", normalized)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        if re.fullmatch(r"(?:\d{3}-[A-Z0-9]{4,5}|[A-Z]{2}\d{6,})", normalized, re.I):
            return normalized

    return ""
