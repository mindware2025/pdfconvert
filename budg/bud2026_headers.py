# ============================
# BUD2026 HEADERS + QUARTERS
# ============================

# Banner anchors (for merged section headers)
# (banner_title, anchor_header_in_row3, nth_occurrence)
BANNER_ANCHORS_BUD2026 = [
    ("Balance at 31-12-2025", "CustCode", 1),
    ("Q1 2026", "Collections FC\n31-03-2026", 1),
    ("Q2 2026", "Collections FC\n30-06-2026", 1),
    ("Q3 2026", "Collections FC\n30-09-2026", 1),
    ("Q4 2026", "Collections FC\n31-12-2026", 1),
    ("2027", "Collections FC\n31-12-2027", 1),
    ("2028", "Collections FC\n31-12-2028", 1),
]

# Quarter blocks (5 columns each: Collection FC, Expected AR, Provision Effect, AR Provision FC, Spacer)
QUARTER_BLOCKS = {
    "Q1": [
        "Collections FC\n31-03-2026",
        "Expected AR",
        "Provision Effect",
        "AR Provision FC at 31-03-2026",
        "",
    ],
    "Q2": [
        "Collections FC\n30-06-2026",
        "Expected AR",
        "Provision Effect",
        "AR Provision FC at 30-06-2026",
        "",
    ],
    "Q3": [
        "Collections FC\n30-09-2026",
        "Expected AR",
        "Provision Effect",
        "AR Provision FC at 30-09-2026",
        "",
    ],
    "Q4": [
        "Collections FC\n31-12-2026",
        "Expected AR",
        "Provision Effect",
        "AR Provision FC at 31-12-2026",
        "",
    ],
}

QUARTER_BLOCK_SIZE = 5

# ============================
# HEADER REMOVAL UTILITIES
# ============================

def remove_quarter_block(headers, quarter):
    """Remove the entire block of a specific quarter (5 columns: FC, Exp AR, Prov Eff, AR FC, spacer)"""
    if quarter not in QUARTER_BLOCKS:
        return headers
    start_header = QUARTER_BLOCKS[quarter][0]
    if start_header not in headers:
        return headers
    start_idx = headers.index(start_header)
    end_idx = start_idx + QUARTER_BLOCK_SIZE
    return headers[:start_idx] + headers[end_idx:]

def filter_headers_by_quarter(headers, selected_quarter):
    """
    Keep:
      - All base columns
      - Selected quarter block
      - Later quarter blocks
      - Always keep 2027 + 2028 blocks
    Remove:
      - Earlier quarter blocks
    """
    order = ["Q1", "Q2", "Q3", "Q4"]
    if selected_quarter == "Q1":
        return headers
    idx = order.index(selected_quarter)
    for q in order[:idx]:
        headers = remove_quarter_block(headers, q)
    return headers

# ============================
# FULL HEADERS (IN ORDER)
# ============================

HEADERS_BUD2026 = [
    '',
    'CustCode',
    'Cust Name',
    'BT',
    'Sales Budget region',
    'Cust Region',
    'Customer Status',
    'Main Ac',
    'Focus List',
    'Insurance',

    # Aging + Balances
    'On\nAccount',
    'Not Due\nAmount',
    'Aging\n1 to 60',
    'Aging\n61 to 90',
    'Aging\n91 to 120',
    'Aging\n121 to 150',
    'Aging\n>=151',
    ' AR\nBalance',

    # Starting balances
    'AR Provision at\n31-08-2025',
    'AR Provision at\n31-12-2024',
    'Provision without any collection',
    'Provision after collection',
    'Provision after collection including Insurance/BG/LC',
    'Difference in Provision',
    '',

    # ---- Quarterly blocks ----
    *QUARTER_BLOCKS["Q1"],
    *QUARTER_BLOCKS["Q2"],
    *QUARTER_BLOCKS["Q3"],
    *QUARTER_BLOCKS["Q4"],

    # ---- 2027 ----
    'Collections FC\n31-12-2027',
    'Expected AR',
    'Provision Effect',
    'AR Provision FC at 31-12-2027',
    '',

    # ---- 2028 ----
    'Collections FC\n31-12-2028',
    'Expected AR',
    'Provision Effect',
    'AR Provision FC at 31-12-2028',
]