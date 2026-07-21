
"""Preview the management dashboard with sample data — no Google Sheets needed.

Run:  streamlit run dashboard_preview.py          (in-app look)
      open with ?view=dashboard in the URL        (TV / wall look)

The catalog rows below mirror Book1t.xlsx (the manual "what we did this month"
Excel) so you can see how it looks once that data lives in the "Tool Catalog"
Sheets tab instead. Safe to delete once you're happy with the real dashboard.
"""
import random
from datetime import datetime, timedelta

import streamlit as st

from dashboard import CATALOG_HEADERS, RUN_LOG_HEADERS, render_dashboard

st.set_page_config(page_title="Dashboard preview", layout="wide")

TOOLS = [
    "Google Automation", "google Automation", "credit Automation ",
    "dell quotation-STD-USD", "dell orion-AED", "MIBB Quotations-TLS",
    "IBM Automation (UAE) temp 2", "Lenovo quotation", "Oracle Automation ",
    "IBM Credit Note Automation (KSA)", "Claims Automation",
]
TEAMS = ["Finance", "Operations", "Credit", "Sales"]


class SampleRunLogSheet:
    def get_all_values(self):
        random.seed(7)
        rows = [list(RUN_LOG_HEADERS)]
        now = datetime.now()
        for _ in range(300):
            ts = now - timedelta(days=random.randint(0, 40), hours=random.randint(0, 23))
            rows.append([
                ts.strftime("%Y-%m-%d %H:%M:%S"),
                random.choice(TOOLS),
                ts.strftime("%b-%Y"),
                random.choice(TEAMS),
                str(random.randint(0, 5)),
                random.choice(["", "0", "1"]),
            ])
        return rows


# Mirrors Book1t.xlsx (description column dropped, per the plan to skip it).
CATALOG_ROWS = [
    ["Lenovo quote", "Sales", "Generates a Lenovo quotation with margin applied.", "Live", "2026-07-14", 20, 0.5, "", 80],
    ["IBM Credit Note Automation (KSA)", "Finance", "Converts IBM KSA credit notes into a CNTS upload file.", "Live", "2026-07-14", "", 0.5, "", ""],
    ["CI and Packing list - IBM", "Operations", "Builds IBM commercial invoice and packing list documents.", "Live", "2026-07-14", "", 0.5, "", ""],
    ["Dell Quotation Southcomp Polaris", "Sales", "Generates Southcomp Polaris quotations from Dell BOQs.", "Live", "2026-07-14", "", 0.5, "", ""],
    ["MIBB Quotations", "Sales", "Generates MIBB quotations (Software, Hardware, TLS).", "Live", "2026-07-14", 25, 0.5, "", 30],
    ["Freight Forwarder JV Tool", "Finance", "Converting multiple pdfs to DNS entry.", "Live", "2026-04-25", 120, 0, "", 1],
    ["ms invoice tool", "Operations", "Convert the invoice report from IW platform into DN and srcl file.", "Live", "2026-07-01", "", "", "", ""],
    ["AR Backlog", "Credit", "Prepares AR backlog, collection and provision forecast.", "test", "", "", "", "", ""],
    ["Dell Quotation automatiom", "Sales", "Converts a Dell BOQ into a formatted quotation Excel.", "Live", "", 15, 0.5, "", 30],
    ["Dell Quotation (Orion)", "Sales", "Converts a Dell BOQ into the Orion export format.", "Live", "", "", "", "", ""],
    ["Lenovo CNTS Tool - KSA", "Finance", "Converts Lenovo credit note PDFs into CNTS format for KSA.", "Live", "2026-04-20", 350, 0.5, "", 1],
    ["Lenovo Credit Note Tool - UAE", "Finance", "Converts 50+ Lenovo credit note PDFs into one CNS Excel.", "Live", "2026-04-01", 250, 0.5, "", 1],
    ["Oracle invoices", "Finance", "Extracts all required fields from Oracle invoice PDFs.", "Live", "2026-03-04", 300, 0.5, "", 1],
    ["Google DNTS Extractor", "Finance", "Turns Google Cloud invoice PDFs into DNTS/Orion Excel.", "Live", "2025-07-06", 45, 0.5, "", 2],
    ["Google Invoice Extractor", "Finance", "Pulls domain-level costs from Google invoice PDFs.", "Live", "2025-08-06", 160, 0.5, "", 2],
    ["Claims Automation", "Finance", "Combines 4 claims files into one ready-to-use claims file.", "Live", "2025-08-14", 120, 0.5, "", 1],
    ["AWS Invoice Tool", "Finance", "Reads AWS invoice/credit-note PDFs into DNTS/CNTS files.", "Live", "2025-09-01", 60, 0.5, "", 1],
    ["Dell Invoice Extractor (Pre-Alert Upload)", "Operations", "Turns Dell invoice PDFs into the pre-alert upload sheet.", "Live", "2025-08-29", 60, 0.5, "", 6],
    ["Cloud Invoice Tool", "Operations", "Converts cloud billing export into invoice/SRCL files.", "Live", "2025-08-27", 60, 0.5, "", 5],
    ["Barcode PDF Generator", "Operations", "Generates pallet barcode labels from Pallet ID/IMEI list.", "Live", "2025-10-13", 60, 0.5, "", 4],
    ["AR to EDD File", "Credit", "Filters AR export to the ageing window needed for EDD.", "Live", "2025-10-15", 90, 0.5, "", 1],
    ["Coface CSV Uploader", "Credit", "Splits customer-invoice export into per-customer CSVs.", "Live", "2025-10-15", 90, 0.5, "", 1],
    ["IBM Quotation", "Sales", "Builds one styled IBM quotation Excel per country/currency.", "Live", "2026-02-06", 25, 0.5, "", 85],
]


class SampleCatalogSheet:
    def get_all_values(self):
        return [list(CATALOG_HEADERS)] + [[str(c) for c in row] for row in CATALOG_ROWS]


render_dashboard(
    SampleRunLogSheet(),
    SampleCatalogSheet(),
    tv_mode=st.query_params.get("view") == "dashboard",
)
