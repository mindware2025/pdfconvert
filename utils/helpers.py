import re
from datetime import datetime

def normalize_line(line):
    line = re.sub(r"[.]+", "", line)
    return re.sub(r"\s+", " ", line).strip()

def format_invoice_date(date_str):
    try:
        dt = datetime.strptime(date_str, "%d %b %Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        dt = datetime.strptime(date_str, "%d %B %Y")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str

def format_amount(amount_str):
    try:
        amount = float(amount_str.replace(",", ""))
        if amount.is_integer():
            return str(int(amount))
        else:
            return str(amount).rstrip('0').rstrip('.') if '.' in str(amount) else str(amount)
    except Exception:
        return amount_str

def format_month_year(date_str):
    import calendar
    import re
    m = re.match(r"(\d{1,2}) ([A-Za-z]+) (\d{4})", date_str)
    if m:
        month = m.group(2)
        year = m.group(3)
        return f"{month} {year}"
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", date_str)
    if m:
        month_num = int(m.group(2))
        year = m.group(3)
        month = calendar.month_abbr[month_num]
        return f"{month} {year}"
    return date_str 