"""Preview the management dashboard with sample data — no Google Sheets needed.

Run:  streamlit run dashboard_preview.py          (in-app look)
      open with ?view=dashboard in the URL        (TV / wall look)

Safe to delete once you're happy with the real dashboard.
"""
import random
from datetime import datetime, timedelta

import streamlit as st

from dashboard import RUN_LOG_HEADERS, render_dashboard

st.set_page_config(page_title="Dashboard preview", layout="wide")

TOOLS = [
    "Google Automation", "google Automation", "credit Automation ",
    "dell quotation-STD-USD", "dell orion-AED", "MIBB Quotations-TLS",
    "IBM Automation (UAE) temp 2", "Lenovo quotation", "Oracle Automation ",
    "IBM Credit Note Automation (KSA)", "Claims Automation",
]
TEAMS = ["Finance", "Operations", "Credit", "Sales"]


class SampleSheet:
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


render_dashboard(SampleSheet(), tv_mode=st.query_params.get("view") == "dashboard")
