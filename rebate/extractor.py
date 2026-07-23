"""
Independent PDF text extraction for the UAE rebate feature.

Deliberately does NOT import or call anything from ibm.py / ibm_template2.py /
sales/ibm_v2_combo.py — it re-reads the raw PDF text on its own, with its own
copies of whatever regex it needs, so the existing quotation-generation code
can never be affected by anything in here.

Produces a flat list of line items: {part_number, commit_value_usd,
channel_pct, term_months}, regardless of which of the two PDF layouts
("Parts Information" flat table vs "Subscription License"/"Software as a
Service" block layout) the quote uses.
"""

import re
from datetime import datetime

import fitz
from dateutil.relativedelta import relativedelta

MONEY_RE = re.compile(r'^-?[\d.]+,\d{2}$')
DATE_RE = re.compile(r'^\d{2}-[A-Za-z]{3}-\d{4}$')
SKU_RE = re.compile(r'^[A-Z][A-Z0-9]{4,8}$')

MAX_ROW_WINDOW = 30


def _looks_like_sku(token):
    return bool(SKU_RE.fullmatch(token)) and any(c.isdigit() for c in token)


def _parse_money(token):
    token = token.strip()
    if token.upper().endswith('USD'):
        token = token[:-3].strip()
    token = token.rstrip('-')
    if not token:
        return None
    try:
        return float(token.replace('.', '').replace(',', '.'))
    except ValueError:
        return None


def _parse_pct(text):
    try:
        return float(text.replace(',', '.'))
    except ValueError:
        return None


def _full_text(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    try:
        return "\n".join(page.get_text('text') for page in doc)
    finally:
        doc.close()


def extract_line_items(pdf_bytes):
    """Return a list of {part_number, commit_value_usd, channel_pct, term_months}."""
    text = _full_text(pdf_bytes)
    lower = text.lower()
    if 'subscription part#' in lower or 'overage part#' in lower:
        return _extract_block_layout(text)
    if 'parts information' in lower:
        return _extract_flat_layout(text)
    return []


# ---------------------------------------------------------------------------
# Block layout: "Subscription License" / "Software as a Service" sections,
# each product block carrying "Channel Discount: X%" and
# "Subscription Length: X Months", with numbered (001, 002, ...) rows below.
# ---------------------------------------------------------------------------

_BLOCK_STOP_PREFIXES = (
    'Subtotal',
    'Subscription Part#:',
    'Overage Part#:',
)


def _extract_block_layout(text):
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    n = len(lines)

    current_part = None
    current_is_overage = False
    current_channel_pct = None
    current_term_months = None

    items = []
    i = 0
    while i < n:
        line = lines[i]

        m = re.match(r'Subscription Part#:\s*(\S+)', line)
        if m:
            current_part = m.group(1)
            current_is_overage = False
            i += 1
            continue

        m = re.match(r'Overage Part#:\s*(\S+)', line)
        if m:
            current_part = m.group(1)
            current_is_overage = True
            i += 1
            continue

        m = re.match(r'Channel Discount:\s*([\d.,]+)\s*%', line)
        if m:
            current_channel_pct = _parse_pct(m.group(1))
            i += 1
            continue

        m = re.match(r'Subscription Length:\s*(\d+)\s*Months?', line, re.IGNORECASE)
        if m:
            current_term_months = int(m.group(1))
            i += 1
            continue

        if re.fullmatch(r'\d{3}', line) and current_part is not None:
            j = i + 1
            window_lines = []
            while j < n and len(window_lines) < MAX_ROW_WINDOW:
                nxt = lines[j]
                if re.fullmatch(r'\d{3}', nxt) or any(nxt.startswith(p) for p in _BLOCK_STOP_PREFIXES):
                    break
                window_lines.append(nxt)
                j += 1

            if current_is_overage:
                # Overage lines are per-unit usage rates with no committed
                # total in the PDF at all -> no rebate base for these rows.
                commit_value = 0.0
            else:
                tokens = ' '.join(window_lines).split()
                money_vals = [v for t in tokens if MONEY_RE.match(t) and (v := _parse_money(t)) is not None]
                # Row shape: [entitled_rate, entitled_total, bid_rate,
                # bid_total_commit, ...]. bid_total_commit is always the 4th
                # money-shaped token, verified against real sample quotes.
                commit_value = money_vals[3] if len(money_vals) >= 4 else 0.0

            items.append({
                'part_number': current_part,
                'commit_value_usd': commit_value,
                'channel_pct': current_channel_pct,
                'term_months': current_term_months if current_term_months is not None else 12,
            })
            i = j
            continue

        i += 1

    return items


# ---------------------------------------------------------------------------
# Flat layout: a single "Parts Information" table, one bare SKU line per row,
# "Channel Margin: X%" embedded near the row, Coverage Start/End dates.
# ---------------------------------------------------------------------------

_FLAT_EXCLUDE_PREFIXES = (
    'Standard Price:',
    'IBM Opportunity Number:',
    'Channel Margin:',
    'Current Transaction',
)


def _extract_flat_layout(text):
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    n = len(lines)

    items = []
    i = 0
    while i < n:
        line = lines[i]
        if _looks_like_sku(line):
            part_number = line
            j = i + 1
            window_lines = []
            while j < n and len(window_lines) < MAX_ROW_WINDOW:
                nxt = lines[j]
                if _looks_like_sku(nxt):
                    break
                window_lines.append(nxt)
                if nxt.startswith('IBM Opportunity Number:'):
                    if j + 1 < n:
                        window_lines.append(lines[j + 1])
                        j += 1
                    j += 1
                    break
                j += 1

            window_text = ' '.join(window_lines)
            channel_pct = None
            cm = re.search(r'Channel Margin:\s*([\d.,]+)\s*%', window_text)
            if cm:
                channel_pct = _parse_pct(cm.group(1))

            all_tokens = window_text.split()
            dates = [t for t in all_tokens if DATE_RE.fullmatch(t)]
            term_months = 12
            if len(dates) >= 2:
                try:
                    start = datetime.strptime(dates[0], '%d-%b-%Y')
                    end = datetime.strptime(dates[1], '%d-%b-%Y')
                    delta = relativedelta(end, start)
                    term_months = delta.years * 12 + delta.months
                except ValueError:
                    pass

            filtered_lines = [
                l for l in window_lines
                if not any(l.startswith(p) for p in _FLAT_EXCLUDE_PREFIXES)
            ]
            filtered_tokens = ' '.join(filtered_lines).split()
            money_vals = [v for t in filtered_tokens if MONEY_RE.match(t) and (v := _parse_money(t)) is not None]
            # Row shape: [Unit Points, Ext Points, Entitled Unit SVP,
            # Entitled Ext SVP, Bid Unit SVP, Bid Ext SVP]. Bid Ext SVP (the
            # commit value) is always the last money-shaped token in the
            # row's own window, verified against real sample quotes.
            commit_value = money_vals[-1] if money_vals else 0.0

            items.append({
                'part_number': part_number,
                'commit_value_usd': commit_value,
                'channel_pct': channel_pct,
                'term_months': term_months,
            })
            i = j
            continue
        i += 1

    return items
