"""MINDBOT management dashboard.

Reads two Google Sheets worksheets, both read-only:
  - "Run Log" (written by log_run_details in app.py) — one row per automation run.
  - "Tool Catalog" (maintained by hand in Sheets, replacing the old manual Excel
    handoff) — one row per tool with its manual/automated time-per-run estimate,
    go-live date, status, and a monthly-runs estimate for tools not yet logged.

Rendered in two modes:
  - in-app view (behind the normal login), via render_dashboard(run_log_sheet, catalog_sheet)
  - TV / wall mode (?view=dashboard, no login), via render_dashboard(..., tv_mode=True)

This module must never import app.py (importing it would re-run the whole
Streamlit script) and never writes to Google Sheets.
"""

from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from gspread.exceptions import APIError

RUN_LOG_HEADERS = ["Timestamp", "Tool", "Month", "Team", "PDF Count", "Excel Count"]

TOOL_CATALOG_SHEET_NAME = "Tool Catalog"
CATALOG_HEADERS = [
    "Tool",
    "Owner Team",
    "Summary",
    "Status",
    "Go-Live Date",
    "Manual Time (min/run)",
    "Automated Time (min/run)",
    "Time Saved per Run (min)",
    "Runs per Month",
]

CACHE_TTL_SECONDS = 240   # at most one Sheets read per 4 min per client
TV_REFRESH_SECONDS = 300  # fragment rerun cadence; >= TTL so each refresh gets fresh data

PALETTE = {
    "navy": "#0c1930",
    "card": "#10223e",
    "steel": "#6ea8ff",
    "glow": "#7cc7ff",
    "white": "#f2f7ff",
    "dim": "rgba(157,193,255,0.65)",
}
GRID = "rgba(110,168,255,0.10)"

# Status colors reserved for state, never reused as a categorical series color.
STATUS_COLORS = {"live": "#0ca30c", "test": "#fab219", "default": "#6b7a94"}

# Categorical colors for the team donut — validated 4-slot set (dataviz
# validator, dark surface #10223e, all-pairs). Fixed per team, never by rank.
TEAM_ORDER = ["Finance", "Operations", "Credit", "Sales"]
TEAM_COLORS = {
    "Finance": "#3987e5",
    "Operations": "#008300",
    "Credit": "#d55181",
    "Sales": "#c98500",
}
TEAM_FALLBACK_COLOR = "#6b7a94"

# ---------------------------------------------------------------------------
# Tool-name normalization — the logged names are free-form (mixed casing,
# trailing spaces, dynamic f-string suffixes); collapse them for clean grouping.
# ---------------------------------------------------------------------------

_EXACT = {
    "google automation": "Google Extractor",
    "claims automation": "Claims Automation",
    "credit automation": "Credit Automation",
    "oracle automation": "Oracle Invoice",
    "aws automation": "AWS Invoice",
    "dell automation": "Dell Invoice",
    "barcode automation": "Barcode Generator",
    "freight_forwarder_expeditor": "Freight Forwarder",
    "credit format by customer": "Credit Format",
    "ibm packing list automation": "IBM Packing List",
    "ibm credit note automation (ksa)": "IBM Credit Note (KSA)",
    "lenovo credit note tool - uae": "Lenovo Credit Note (UAE)",
    "lenovo credit note tool - ksa": "Lenovo Credit Note (KSA)",
    "lenovo quotation": "Lenovo Quotation",
}
_PREFIXES = [
    ("mibb quotations", "MIBB Quotation"),
    ("dell quotation", "Dell Quotation"),
    ("dell orion", "Dell Quotation (Orion)"),
    ("ibm automation", "IBM Quotation"),
]


def normalize_tool_name(raw) -> str:
    key = str(raw).strip().casefold()
    if key in _EXACT:
        return _EXACT[key]
    for prefix, label in _PREFIXES:
        if key.startswith(prefix):
            return label
    return str(raw).strip() or "Unknown"


# The "Tool Catalog" sheet is filled in by hand with human-friendly labels that
# don't always match the strings update_usage() logs (e.g. "Lenovo quote" vs
# the logged "Lenovo quotation", or "Dell Quotation (Orion)" which would
# otherwise collide with the generic "dell quotation" prefix rule above).
# These aliases are checked first so the catalog joins onto real Run Log data;
# anything not listed here just falls back to normalize_tool_name's own rules.
CATALOG_ALIASES = {
    "lenovo quote": "Lenovo Quotation",
    "ci and packing list - ibm": "IBM Packing List",
    "freight forwarder jv tool": "Freight Forwarder",
    "lenovo cnts tool - ksa": "Lenovo Credit Note (KSA)",
    "oracle invoices": "Oracle Invoice",
    "dell quotation (orion)": "Dell Quotation (Orion)",
    "dell invoice extractor (pre-alert upload)": "Dell Invoice",
    "google dnts extractor": "Google Extractor",
    "google invoice extractor": "Google Extractor",
}


def normalize_catalog_tool(raw) -> str:
    key = str(raw).strip().casefold()
    if key in CATALOG_ALIASES:
        return CATALOG_ALIASES[key]
    return normalize_tool_name(raw)


# ---------------------------------------------------------------------------
# Data loading — Run Log
# ---------------------------------------------------------------------------


def _empty_run_log() -> pd.DataFrame:
    df = pd.DataFrame({col: pd.Series(dtype="object") for col in RUN_LOG_HEADERS})
    df["Timestamp"] = pd.Series(dtype="datetime64[ns]")
    df["PDF Count"] = pd.Series(dtype="int64")
    df["Excel Count"] = pd.Series(dtype="int64")
    df["date"] = pd.Series(dtype="datetime64[ns]")
    df["tool_clean"] = pd.Series(dtype="object")
    return df


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def load_run_log(_sheet) -> pd.DataFrame:
    values = _sheet.get_all_values()
    if len(values) < 2:
        return _empty_run_log()

    n = len(RUN_LOG_HEADERS)
    rows = [(row + [""] * n)[:n] for row in values[1:]]
    df = pd.DataFrame(rows, columns=RUN_LOG_HEADERS)

    df["Timestamp"] = pd.to_datetime(
        df["Timestamp"], format="%Y-%m-%d %H:%M:%S", errors="coerce"
    )
    df = df.dropna(subset=["Timestamp"])
    if df.empty:
        return _empty_run_log()

    for col in ("PDF Count", "Excel Count"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df["Team"] = df["Team"].astype(str).str.strip().str.title()
    df["date"] = df["Timestamp"].dt.normalize()
    df["tool_clean"] = df["Tool"].map(normalize_tool_name)
    return df


def compute_kpis(df: pd.DataFrame, now: datetime) -> dict:
    kpis = {
        "pdfs_today": 0,
        "runs_today": 0,
        "pdfs_month": 0,
        "runs_month": 0,
        "top_tool": "—",
        "top_tool_count": 0,
    }
    if df.empty:
        return kpis

    today_mask = df["Timestamp"].dt.date == now.date()
    month_mask = (df["Timestamp"].dt.month == now.month) & (
        df["Timestamp"].dt.year == now.year
    )
    mdf = df[month_mask]

    kpis["pdfs_today"] = int(df.loc[today_mask, "PDF Count"].sum())
    kpis["runs_today"] = int(today_mask.sum())
    kpis["pdfs_month"] = int(mdf["PDF Count"].sum())
    kpis["runs_month"] = int(len(mdf))

    if not mdf.empty:
        sums = mdf.groupby("tool_clean")["PDF Count"].sum()
        if sums.sum() == 0:
            sums = mdf["tool_clean"].value_counts()
        top = sums.sort_values(ascending=False)
        kpis["top_tool"] = str(top.index[0])
        kpis["top_tool_count"] = int(top.iloc[0])
    return kpis


# ---------------------------------------------------------------------------
# Data loading — Tool Catalog
# ---------------------------------------------------------------------------


def _empty_tool_catalog() -> pd.DataFrame:
    cols = CATALOG_HEADERS + ["tool_clean"]
    df = pd.DataFrame({c: pd.Series(dtype="object") for c in cols})
    df["Go-Live Date"] = pd.Series(dtype="datetime64[ns]")
    for c in (
        "Manual Time (min/run)",
        "Automated Time (min/run)",
        "Time Saved per Run (min)",
        "Runs per Month",
    ):
        df[c] = pd.Series(dtype="float64")
    return df


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def load_tool_catalog(_sheet) -> pd.DataFrame:
    values = _sheet.get_all_values()
    if len(values) < 2:
        return _empty_tool_catalog()

    n = len(CATALOG_HEADERS)
    rows = [(row + [""] * n)[:n] for row in values[1:]]
    df = pd.DataFrame(rows, columns=CATALOG_HEADERS)
    df = df[df["Tool"].astype(str).str.strip() != ""].reset_index(drop=True)
    if df.empty:
        return _empty_tool_catalog()

    df["Owner Team"] = df["Owner Team"].astype(str).str.strip().str.title()
    df["Status"] = (
        df["Status"].astype(str).str.strip().str.title().replace({"": "Unknown"})
    )
    df["Go-Live Date"] = pd.to_datetime(df["Go-Live Date"], errors="coerce")
    for col in (
        "Manual Time (min/run)",
        "Automated Time (min/run)",
        "Time Saved per Run (min)",
        "Runs per Month",
    ):
        df[col] = pd.to_numeric(df[col], errors="coerce")

    derived = df["Manual Time (min/run)"] - df["Automated Time (min/run)"]
    df["Time Saved per Run (min)"] = df["Time Saved per Run (min)"].fillna(derived)
    df["tool_clean"] = df["Tool"].map(normalize_catalog_tool)
    return df


def compute_time_saved(catalog_df: pd.DataFrame, run_log_df: pd.DataFrame, now: datetime) -> dict:
    """Combine the manual catalog assumptions with real Run Log volumes.

    A tool that has ever appeared in the Run Log uses its real count of runs
    *this month* (even if that's zero — more honest than a stale guess). A
    tool never wired into update_usage() falls back to the catalog's own
    "Runs per Month" estimate, flagged so the UI can mark it as such.

    Some catalog rows share one canonical bucket because the app currently
    logs them under the same generic name (e.g. "Google DNTS Extractor" and
    "Google Invoice Extractor" both log as "Google Automation"). Only the
    first such row claims the live count, so real runs are never counted
    twice under two different per-run rates; the rest fall back to their own
    catalog estimate.
    """
    result = {"total_hours": 0.0, "rows": []}
    if catalog_df.empty:
        return result

    if run_log_df.empty:
        tracked_tools = set()
        run_counts_this_month = pd.Series(dtype="int64")
    else:
        month_mask = (run_log_df["Timestamp"].dt.month == now.month) & (
            run_log_df["Timestamp"].dt.year == now.year
        )
        tracked_tools = set(run_log_df["tool_clean"].unique())
        run_counts_this_month = run_log_df.loc[month_mask, "tool_clean"].value_counts()

    total_minutes = 0.0
    rows = []
    claimed_canonicals = set()
    for _, row in catalog_df.iterrows():
        canonical = row["tool_clean"]
        per_run = row["Time Saved per Run (min)"]
        is_tracked = canonical in tracked_tools and canonical not in claimed_canonicals
        if is_tracked:
            runs = int(run_counts_this_month.get(canonical, 0))
            estimated = False
            claimed_canonicals.add(canonical)
        else:
            raw_runs = row["Runs per Month"]
            runs = 0 if pd.isna(raw_runs) else int(raw_runs)
            estimated = True

        saved_minutes = 0.0
        if pd.notna(per_run):
            saved_minutes = float(per_run) * runs
            total_minutes += saved_minutes

        rows.append(
            {
                "Tool": row["Tool"],
                "Team": row["Owner Team"],
                "Status": row["Status"],
                "Go-Live": row["Go-Live Date"],
                "Runs": runs,
                "Estimated": estimated,
                "PerRun": per_run,
                "SavedMinutes": saved_minutes,
            }
        )

    result["total_hours"] = total_minutes / 60
    result["rows"] = rows
    return result


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------


def _apply_dark(fig: go.Figure, title: str) -> go.Figure:
    fig.update_layout(
        template="plotly_dark",
        title=dict(
            text=title,
            x=0.01,
            xanchor="left",
            font=dict(size=15, color=PALETTE["dim"]),
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(
            family="Google Sans, Segoe UI, sans-serif",
            color=PALETTE["white"],
            size=13,
        ),
        margin=dict(t=44, l=8, r=8, b=8),
        hoverlabel=dict(
            bgcolor="#14294a",
            bordercolor="rgba(110,168,255,0.35)",
            font=dict(color=PALETTE["white"], family="Google Sans, Segoe UI, sans-serif"),
        ),
        showlegend=False,
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, linecolor="rgba(110,168,255,0.20)")
    fig.update_yaxes(gridcolor=GRID, zeroline=False, linecolor="rgba(0,0,0,0)")
    return fig


def build_trend_chart(df: pd.DataFrame, now: datetime, compact: bool = False) -> go.Figure:
    idx = pd.date_range(end=pd.Timestamp(now).normalize(), periods=30, freq="D")
    daily = df.groupby("date")["PDF Count"].sum().reindex(idx, fill_value=0)

    fig = go.Figure(
        go.Scatter(
            x=idx,
            y=daily.values,
            mode="lines",
            line=dict(color=PALETTE["glow"], width=3, shape="spline", smoothing=0.6),
            fill="tozeroy",
            fillcolor="rgba(124,199,255,0.15)",
            hovertemplate="%{x|%d %b}: <b>%{y}</b> PDFs<extra></extra>",
        )
    )
    _apply_dark(fig, "PDFs processed — last 30 days")
    fig.update_layout(height=185 if compact else 250, hovermode="x unified")
    fig.update_xaxes(
        showspikes=True,
        spikemode="across",
        spikecolor="rgba(124,199,255,0.35)",
        spikethickness=1,
        tickformat="%d %b",
    )
    fig.update_yaxes(rangemode="tozero")
    return fig


def build_top_tools_chart(df: pd.DataFrame, now: datetime, compact: bool = False) -> go.Figure:
    month_mask = (df["Timestamp"].dt.month == now.month) & (
        df["Timestamp"].dt.year == now.year
    )
    mdf = df[month_mask]
    sums = mdf.groupby("tool_clean")["PDF Count"].sum()
    if sums.sum() == 0:
        sums = mdf["tool_clean"].value_counts()
    top = sums.sort_values(ascending=True).tail(8)

    fig = go.Figure(
        go.Bar(
            x=top.values,
            y=top.index,
            orientation="h",
            marker=dict(color=PALETTE["steel"]),
            text=[f"{int(v):,}" for v in top.values],
            textposition="outside",
            textfont=dict(color=PALETTE["white"], size=12),
            cliponaxis=False,
            hovertemplate="%{y}: <b>%{x}</b> PDFs<extra></extra>",
        )
    )
    _apply_dark(fig, "Top tools this month")
    fig.update_layout(height=250 if compact else 320, bargap=0.35, barcornerradius=4)
    fig.update_xaxes(showgrid=True)
    fig.update_yaxes(showgrid=False, tickfont=dict(color=PALETTE["dim"], size=12))
    return fig


def build_team_chart(df: pd.DataFrame, now: datetime, compact: bool = False) -> go.Figure:
    month_mask = (df["Timestamp"].dt.month == now.month) & (
        df["Timestamp"].dt.year == now.year
    )
    mdf = df[month_mask]
    sums = mdf.groupby("Team")["PDF Count"].sum()
    if sums.sum() == 0:
        sums = mdf["Team"].value_counts()
    sums = sums[sums > 0]

    order = [t for t in TEAM_ORDER if t in sums.index] + [
        t for t in sums.index if t not in TEAM_ORDER
    ]
    colors = [TEAM_COLORS.get(t, TEAM_FALLBACK_COLOR) for t in order]
    total = int(sums.sum())

    fig = go.Figure(
        go.Pie(
            labels=order,
            values=[int(sums[t]) for t in order],
            hole=0.55,
            sort=False,
            direction="clockwise",
            marker=dict(colors=colors, line=dict(color=PALETTE["card"], width=2)),
            textinfo="label+percent",
            textfont=dict(color=PALETTE["white"], size=12),
            hovertemplate="%{label}: <b>%{value}</b> PDFs (%{percent})<extra></extra>",
        )
    )
    _apply_dark(fig, "Split by team — this month")
    fig.update_layout(
        height=250 if compact else 320,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.15,
            x=0.5,
            xanchor="center",
            font=dict(color=PALETTE["dim"], size=12),
        ),
        annotations=[
            dict(
                text=f"<b>{total:,}</b><br><span style='font-size:11px'>PDFs</span>",
                showarrow=False,
                font=dict(color=PALETTE["white"], size=22),
            )
        ],
    )
    return fig


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------


def _kpi_tile(label: str, value: str, sub: str = "") -> str:
    return (
        '<div class="mw-kpi">'
        f'<div class="mw-kpi-label">{label}</div>'
        f'<div class="mw-kpi-value">{value}</div>'
        f'<div class="mw-kpi-sub">{sub}</div>'
        "</div>"
    )


def _time_saved_hero(total_hours: float) -> str:
    return (
        '<div class="mw-hero">'
        '<div class="mw-hero-label">⚡ Estimated time saved this month</div>'
        f'<div class="mw-hero-value">{total_hours:,.1f}<span> hours</span></div>'
        "</div>"
    )


def _status_pill(status: str) -> str:
    key = str(status).strip().lower()
    color = STATUS_COLORS.get(key, STATUS_COLORS["default"])
    label = str(status).strip() or "Unknown"
    return (
        f'<span class="mw-pill" style="--pill-color:{color}">'
        f'<span class="mw-pill-dot"></span>{label}</span>'
    )


def _team_chip(team: str) -> str:
    color = TEAM_COLORS.get(str(team).strip().title(), TEAM_FALLBACK_COLOR)
    label = str(team).strip() or "—"
    return f'<span class="mw-chip" style="--chip-color:{color}">{label}</span>'


def build_catalog_table_html(rows: list, max_rows: int | None = None) -> str:
    ordered = sorted(rows, key=lambda r: r["SavedMinutes"], reverse=True)
    if max_rows:
        ordered = ordered[:max_rows]

    body_parts = []
    for r in ordered:
        go_live = r["Go-Live"]
        go_live_str = go_live.strftime("%d %b %Y") if pd.notna(go_live) else "—"
        runs_str = f'{r["Runs"]:,}' + (" (est.)" if r["Estimated"] else "")
        per_run_str = f'{r["PerRun"]:.0f} min' if pd.notna(r["PerRun"]) else "—"
        saved_str = f'{r["SavedMinutes"] / 60:.1f} hrs' if r["SavedMinutes"] else "—"
        body_parts.append(
            "<tr>"
            f'<td class="mw-tool-name">{r["Tool"]}</td>'
            f"<td>{_team_chip(r['Team'])}</td>"
            f"<td>{_status_pill(r['Status'])}</td>"
            f"<td>{go_live_str}</td>"
            f"<td>{runs_str}</td>"
            f"<td>{per_run_str}</td>"
            f"<td>{saved_str}</td>"
            "</tr>"
        )

    return (
        '<div class="mw-table-wrap"><table class="mw-table"><thead><tr>'
        "<th>Tool</th><th>Team</th><th>Status</th><th>Live since</th>"
        "<th>Runs/mo</th><th>Time saved/run</th><th>Saved this month</th>"
        "</tr></thead><tbody>" + "".join(body_parts) + "</tbody></table></div>"
    )


def _inject_css(tv_mode: bool) -> None:
    tv_extra = ""
    if tv_mode:
        tv_extra = """
        [data-testid="stHeader"] { display: none !important; }
        footer { display: none !important; }
        .block-container {
            padding: 1rem 2rem 0.5rem !important;
            max-width: 100% !important;
        }
        .mw-dash-head { margin: 0 0 0.6rem 0 !important; }
        .mw-hero {
            padding: 0.6rem 1rem 0.5rem !important;
            margin-bottom: 0.5rem !important;
        }
        .mw-hero-value { font-size: clamp(2rem, 4.5vw, 3rem) !important; }
        .mw-kpi { padding: 0.8rem 1.1rem 0.7rem !important; }
        .mw-kpi-value { font-size: clamp(1.8rem, 3.8vw, 3rem) !important; }
        """
    st.markdown(
        f"""
        <style>
        [data-testid="stApp"] {{
            background: radial-gradient(1200px 700px at 50% -10%,
                #14294a 0%, #0c1930 48%, #060e1c 100%) !important;
        }}
        [data-testid="stToolbar"], #MainMenu {{ display: none !important; }}
        html, body, [class*="css"] {{
            background: transparent !important;
            color: {PALETTE["white"]} !important;
        }}
        .mw-dash-head {{
            display: flex; align-items: baseline; justify-content: space-between;
            margin: 0 0 1rem 0;
        }}
        .mw-dash-title {{
            font-family: 'Google Sans', 'Segoe UI', sans-serif;
            font-size: 1.6rem; font-weight: 700; color: {PALETTE["white"]};
            letter-spacing: 2px;
        }}
        .mw-dash-title .mw-dot {{ color: {PALETTE["glow"]}; }}
        .mw-dash-updated {{
            font-size: 0.85rem; color: {PALETTE["dim"]}; letter-spacing: 1px;
        }}
        .mw-hero {{
            text-align: center;
            padding: 1.3rem 1rem 1.1rem;
            margin-bottom: 1rem;
            background: rgba(12, 163, 12, 0.08);
            border: 1px solid rgba(12, 163, 12, 0.35);
            border-radius: 18px;
        }}
        .mw-hero-label {{
            font-size: 0.8rem; letter-spacing: 2px; text-transform: uppercase;
            color: {PALETTE["dim"]}; margin-bottom: 0.3rem;
        }}
        .mw-hero-value {{
            font-family: 'Google Sans', 'Segoe UI', sans-serif;
            font-size: clamp(2.3rem, 5.5vw, 4.2rem); font-weight: 800;
            color: #baf7ba;
            text-shadow: 0 0 26px rgba(12, 163, 12, 0.55);
        }}
        .mw-hero-value span {{
            font-size: 1.1rem; font-weight: 500; color: {PALETTE["dim"]};
            margin-left: 0.4rem; text-shadow: none;
        }}
        .mw-section-title {{
            font-size: 0.95rem; letter-spacing: 1px; text-transform: uppercase;
            color: {PALETTE["dim"]}; margin: 1.2rem 0 0.6rem;
        }}
        .mw-kpi {{
            background: rgba(20, 41, 74, 0.55);
            border: 1px solid rgba(110, 168, 255, 0.25);
            border-radius: 18px;
            padding: 1.1rem 1.3rem 1rem;
            box-shadow: inset 0 0 40px rgba(124, 199, 255, 0.05);
        }}
        .mw-kpi-label {{
            font-size: 0.75rem; text-transform: uppercase; letter-spacing: 2px;
            color: {PALETTE["dim"]}; margin-bottom: 0.35rem;
        }}
        .mw-kpi-value {{
            font-family: 'Google Sans', 'Segoe UI', sans-serif;
            font-size: clamp(2.2rem, 4.5vw, 4rem); font-weight: 700; line-height: 1.05;
            color: {PALETTE["white"]};
            animation: mw-kpi-glow 6s ease-in-out infinite;
        }}
        .mw-kpi-sub {{ font-size: 0.8rem; color: {PALETTE["dim"]}; margin-top: 0.3rem; }}
        @keyframes mw-kpi-glow {{
            0%, 100% {{ text-shadow: 0 0 18px rgba(124, 199, 255, 0.45); }}
            50% {{ text-shadow: 0 0 32px rgba(124, 199, 255, 0.85); }}
        }}
        .mw-empty {{
            border: 1px dashed rgba(110, 168, 255, 0.3); border-radius: 18px;
            padding: 3rem 1rem; text-align: center; color: {PALETTE["dim"]};
        }}
        .mw-banner {{
            background: rgba(20, 41, 74, 0.55);
            border: 1px solid rgba(110, 168, 255, 0.25); border-radius: 12px;
            padding: 1rem 1.2rem; color: {PALETTE["dim"]};
        }}
        .mw-table-wrap {{
            overflow-x: auto;
            border: 1px solid rgba(110, 168, 255, 0.2);
            border-radius: 14px;
        }}
        .mw-table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
        .mw-table th {{
            text-align: left; padding: 0.6rem 0.9rem; color: {PALETTE["dim"]};
            font-size: 0.7rem; text-transform: uppercase; letter-spacing: 1px;
            border-bottom: 1px solid rgba(110, 168, 255, 0.2); white-space: nowrap;
        }}
        .mw-table td {{
            padding: 0.55rem 0.9rem; border-bottom: 1px solid rgba(110, 168, 255, 0.08);
            color: {PALETTE["white"]}; white-space: nowrap;
        }}
        .mw-table tr:last-child td {{ border-bottom: none; }}
        .mw-tool-name {{ font-weight: 600; }}
        .mw-pill {{
            display: inline-flex; align-items: center; gap: 0.35rem;
            padding: 0.2rem 0.6rem; border-radius: 999px;
            background: rgba(255, 255, 255, 0.06);
            color: var(--pill-color); font-size: 0.75rem; font-weight: 600;
        }}
        .mw-pill-dot {{ width: 6px; height: 6px; border-radius: 50%; background: var(--pill-color); }}
        .mw-chip {{
            padding: 0.15rem 0.55rem; border-radius: 8px;
            background: rgba(255, 255, 255, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.08);
            color: var(--chip-color); font-size: 0.78rem; font-weight: 600;
        }}
        .stButton > button {{
            background: rgba(20, 41, 74, 0.8) !important;
            color: {PALETTE["white"]} !important;
            border: 1px solid rgba(110, 168, 255, 0.35) !important;
            box-shadow: none !important;
        }}
        {tv_extra}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _dashboard_body(run_log_sheet, catalog_sheet, tv_mode: bool) -> None:
    now = datetime.now()
    st.markdown(
        '<div class="mw-dash-head">'
        '<div class="mw-dash-title">MINDBOT <span class="mw-dot">·</span> Automation Dashboard</div>'
        f'<div class="mw-dash-updated">Last updated {now.strftime("%H:%M")}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    try:
        df = load_run_log(run_log_sheet)
    except APIError:
        st.markdown(
            '<div class="mw-banner">📡 Live data temporarily unavailable — '
            "retrying automatically.</div>",
            unsafe_allow_html=True,
        )
        return

    catalog_df = None
    if catalog_sheet is not None:
        try:
            catalog_df = load_tool_catalog(catalog_sheet)
        except APIError:
            catalog_df = None

    if catalog_df is not None and not catalog_df.empty:
        ts = compute_time_saved(catalog_df, df, now)
        st.markdown(_time_saved_hero(ts["total_hours"]), unsafe_allow_html=True)

    kpis = compute_kpis(df, now)

    tiles = st.columns(4)
    tiles[0].markdown(
        _kpi_tile("PDFs today", f"{kpis['pdfs_today']:,}", f"{kpis['runs_today']} runs"),
        unsafe_allow_html=True,
    )
    tiles[1].markdown(
        _kpi_tile("PDFs this month", f"{kpis['pdfs_month']:,}", now.strftime("%B %Y")),
        unsafe_allow_html=True,
    )
    tiles[2].markdown(
        _kpi_tile("Automation runs", f"{kpis['runs_month']:,}", "this month"),
        unsafe_allow_html=True,
    )
    tiles[3].markdown(
        _kpi_tile(
            "Top tool",
            kpis["top_tool"],
            f"{kpis['top_tool_count']:,} PDFs this month" if kpis["top_tool_count"] else "",
        ),
        unsafe_allow_html=True,
    )

    st.markdown(
        f"<div style='height:{'0.4rem' if tv_mode else '0.75rem'}'></div>",
        unsafe_allow_html=True,
    )

    if df.empty:
        st.markdown(
            '<div class="mw-empty">No runs logged yet — numbers will appear here '
            "as soon as the team starts processing PDFs.</div>",
            unsafe_allow_html=True,
        )
    else:
        st.plotly_chart(
            build_trend_chart(df, now, compact=tv_mode),
            width="stretch",
            config={"displayModeBar": False},
        )

        col_tools, col_teams = st.columns([3, 2])
        with col_tools:
            st.plotly_chart(
                build_top_tools_chart(df, now, compact=tv_mode),
                width="stretch",
                config={"displayModeBar": False},
            )
        with col_teams:
            st.plotly_chart(
                build_team_chart(df, now, compact=tv_mode),
                width="stretch",
                config={"displayModeBar": False},
            )

    # The catalog table is detail meant to be read at a desk, not from across a
    # room — TV mode stops at the hero number + charts above so everything
    # still fits on one screen without scrolling.
    if catalog_sheet is not None and not tv_mode:
        if catalog_df is None:
            pass  # APIError already implied core data is flaky; don't pile on a second banner
        elif catalog_df.empty:
            st.markdown(
                '<div class="mw-empty">No tools in the catalog yet — add rows to the '
                '"Tool Catalog" sheet tab (Tool, Owner Team, Summary, Status, Go-Live Date, '
                "Manual/Automated Time, Runs per Month) to see time-saved estimates here.</div>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="mw-section-title">Automation Catalog</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                build_catalog_table_html(ts["rows"]),
                unsafe_allow_html=True,
            )


def render_dashboard(run_log_sheet, catalog_sheet=None, tv_mode: bool = False) -> None:
    _inject_css(tv_mode)
    if tv_mode:
        st.fragment(run_every=TV_REFRESH_SECONDS)(_dashboard_body)(
            run_log_sheet, catalog_sheet, tv_mode
        )
    else:
        _dashboard_body(run_log_sheet, catalog_sheet, tv_mode)
