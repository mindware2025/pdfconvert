"""MINDBOT management dashboard.

Read-only view over the "Run Log" worksheet (written by log_run_details in
app.py). Rendered in two modes:
  - in-app view (behind the normal login), via render_dashboard(sheet)
  - TV / wall mode (?view=dashboard, no login), via render_dashboard(sheet, tv_mode=True)

This module must never import app.py (importing it would re-run the whole
Streamlit script) and never writes to Google Sheets.
"""

from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from gspread.exceptions import APIError

RUN_LOG_HEADERS = ["Timestamp", "Tool", "Month", "Team", "PDF Count", "Excel Count"]

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


# ---------------------------------------------------------------------------
# Data loading
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


def build_trend_chart(df: pd.DataFrame, now: datetime) -> go.Figure:
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
    fig.update_layout(height=250, hovermode="x unified")
    fig.update_xaxes(
        showspikes=True,
        spikemode="across",
        spikecolor="rgba(124,199,255,0.35)",
        spikethickness=1,
        tickformat="%d %b",
    )
    fig.update_yaxes(rangemode="tozero")
    return fig


def build_top_tools_chart(df: pd.DataFrame, now: datetime) -> go.Figure:
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
    fig.update_layout(height=320, bargap=0.35, barcornerradius=4)
    fig.update_xaxes(showgrid=True)
    fig.update_yaxes(showgrid=False, tickfont=dict(color=PALETTE["dim"], size=12))
    return fig


def build_team_chart(df: pd.DataFrame, now: datetime) -> go.Figure:
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
        height=320,
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


def _dashboard_body(run_log_sheet, tv_mode: bool) -> None:
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

    st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)

    if df.empty:
        st.markdown(
            '<div class="mw-empty">No runs logged yet — numbers will appear here '
            "as soon as the team starts processing PDFs.</div>",
            unsafe_allow_html=True,
        )
        return

    st.plotly_chart(
        build_trend_chart(df, now), width="stretch", config={"displayModeBar": False}
    )

    col_tools, col_teams = st.columns([3, 2])
    with col_tools:
        st.plotly_chart(
            build_top_tools_chart(df, now),
            width="stretch",
            config={"displayModeBar": False},
        )
    with col_teams:
        st.plotly_chart(
            build_team_chart(df, now),
            width="stretch",
            config={"displayModeBar": False},
        )


def render_dashboard(run_log_sheet, tv_mode: bool = False) -> None:
    _inject_css(tv_mode)
    if tv_mode:
        st.fragment(run_every=TV_REFRESH_SECONDS)(_dashboard_body)(run_log_sheet, tv_mode)
    else:
        _dashboard_body(run_log_sheet, tv_mode)
