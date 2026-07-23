"""Turn extracted IBM quote line items into rebate rows.

Fully independent of the quotation-generation code — only consumes the
plain dicts produced by rebate.extractor, never touches ibm.py /
ibm_template2.py.
"""

from rebate.rate_card import RATE_CARD, MULTIYEAR_MIN_MONTHS, INCENTIVE_COLUMN_ORDER, bucket_for


def compute_rebate_rows(line_items):
    """Compute rebate rows and the set of columns that should be displayed.

    line_items: list of dicts, each with keys:
        part_number (str), commit_value_usd (float),
        channel_pct (float or None), term_months (int or None)

    Returns (rows, columns):
        rows: list of dicts, one per line item, each with:
            part_number, commit_value_usd, amounts (dict: incentive name -> $),
            total (float)
        columns: ordered list of incentive names that are structurally
            applicable to at least one row in this quote (i.e. the incentive
            has a nonzero rate for that row's bucket, regardless of whether
            the multiyear condition zeroed out the dollar amount for THIS
            particular row).
    """
    rows = []
    applicable_columns = set()

    for item in line_items:
        part_number = item.get("part_number", "")
        commit_value = item.get("commit_value_usd") or 0.0
        channel_pct = item.get("channel_pct")
        term_months = item.get("term_months") or 0

        bucket = bucket_for(channel_pct)
        amounts = {}

        if bucket is not None:
            for name, rate, requires_multiyear in RATE_CARD[bucket]:
                applicable_columns.add(name)
                if requires_multiyear and term_months < MULTIYEAR_MIN_MONTHS:
                    amounts[name] = 0.0
                else:
                    # Keep full precision here (only round to strip float
                    # noise) rather than rounding to cents per-incentive --
                    # the sample workbook's Total column is the sum of the
                    # unrounded amounts, not a sum of individually
                    # cents-rounded ones, so rounding early would drift the
                    # total by a cent or two.
                    amounts[name] = round(commit_value * rate, 4)

        total = round(sum(amounts.values()), 4)

        rows.append({
            "part_number": part_number,
            "commit_value_usd": commit_value,
            "amounts": amounts,
            "total": total,
        })

    columns = [name for name in INCENTIVE_COLUMN_ORDER if name in applicable_columns]
    return rows, columns
