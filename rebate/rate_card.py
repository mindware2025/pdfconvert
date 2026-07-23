"""
IBM back-end rebate rate card (UAE only).

Keyed by the Channel Discount / Channel Margin % already printed on the
quote PDF. Each entry lists the incentive components paid on top of the
front-end margin, as a percentage of the Bid Total Commit Value for that
line item.

Confirmed with the business owner against real sample quotes:
- 14% bucket -> New SaaS / New SL / SaaS Renewal / SL Renewal / SL Upgrade
- 6% bucket  -> NL on Prem / Trade up & Reinstatement
- 3% bucket  -> S&S Renewal

"New" vs "Renewal" never needs to be distinguished: within each bucket the
rate card gives identical rates regardless of that distinction.

SaaS Adopt (15%, would sit in the 14% bucket) is deliberately NOT included:
verified against a real SaaS line item (bid 21764696, part D0A51ZX) that
IBM does not apply it automatically from the quote data alone.
"""

# Each incentive tuple: (display_name, rate, requires_multiyear)
# requires_multiyear=True means the rate only applies when term_months >= 36.
RATE_CARD = {
    14: [
        ("Select Territory Accelerator", 0.10, False),
        ("Proficiency Incentives", 0.02, False),
        ("Multiyear Incentive", 0.03, True),
    ],
    6: [
        ("Base Sales Incentive", 0.06, False),
    ],
    3: [
        ("Base Sales Incentive", 0.03, False),
        ("On-time Incentive", 0.03, False),
        ("Multiyear Incentive", 0.01, True),
    ],
}

MULTIYEAR_MIN_MONTHS = 36

# Fixed column order for the rebate workbook (only columns with at least one
# applicable row are actually rendered).
INCENTIVE_COLUMN_ORDER = [
    "Base Sales Incentive",
    "Select Territory Accelerator",
    "Proficiency Incentives",
    "Multiyear Incentive",
    "On-time Incentive",
]

REBATE_NOTE = (
    "Note : All rebates are subject to IBM approval and upon completion of "
    "valid proficiency badges and approved DR."
)


def bucket_for(channel_pct):
    """Return the RATE_CARD key matching a Channel Discount/Margin % value.

    Accepts float/int percentages (e.g. 14, 14.0, 6.0, 3.0). Returns None if
    it doesn't match one of the confirmed buckets (14/6/3) — callers should
    treat that as "no rebate for this line" rather than guess.
    """
    if channel_pct is None:
        return None
    try:
        pct = round(float(channel_pct))
    except (TypeError, ValueError):
        return None
    return pct if pct in RATE_CARD else None
