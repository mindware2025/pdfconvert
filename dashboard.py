"""MINDBOT management dashboard.

Reads three Google Sheets worksheets, all read-only:
  - "Sheet1" (written by update_usage in app.py) — monthly run counts per tool
    and team since Oct-2025; the source for the month-by-month numbers.
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
    "navy": "#0f1724",
    "card": "#141d2b",
    "steel": "#7fa6d6",
    "glow": "#8db8e0",
    "white": "#e6ebf2",
    "dim": "#8b97a8",
}
GRID = "rgba(255,255,255,0.06)"

# Status colors reserved for state, never reused as a categorical series color.
STATUS_COLORS = {"live": "#5db39e", "test": "#d4a35c", "default": "#7d8796"}

# Categorical colors for the team donut — validated 4-slot set (dataviz
# validator, dark surface #10223e, all-pairs). Fixed per team, never by rank.
TEAM_ORDER = ["Finance", "Operations", "Credit", "Sales"]
TEAM_COLORS = {
    "Finance": "#d4a35c",
    "Operations": "#5db39e",
    "Credit": "#b08bc9",
    "Sales": "#6c9bd2",
}
TEAM_FALLBACK_COLOR = "#7d8796"

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
    ("southcomp polaris", "Southcomp Polaris"),
    ("joint_line", "Joint Line OCR"),
    ("ibm rebate", "IBM Rebate"),
    ("cloud automation", "Cloud Invoice"),
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
    "dell quotation southcomp polaris": "Southcomp Polaris",
    "southcomp polaris": "Southcomp Polaris",
    "dell quotation automatiom": "Dell Quotation",
    "dell quotation automation": "Dell Quotation",
    "aws invoice tool": "AWS Invoice",
    "cloud invoice tool": "Cloud Invoice",
    "barcode pdf generator": "Barcode Generator",
    "ibm credit note automation (ksa)": "IBM Credit Note (KSA)",
    "joint line ocr": "Joint Line OCR",
}

# Tools that produce a customer quotation (counted in the "Quotations" KPI).
QUOTATION_TOOLS = {
    "Dell Quotation",
    "Dell Quotation (Orion)",
    "Southcomp Polaris",
    "IBM Quotation",
    "MIBB Quotation",
    "Lenovo Quotation",
}
HOURS_PER_FTE_MONTH = 160


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
# Data loading — monthly usage history (Sheet1) and management metrics
# ---------------------------------------------------------------------------


def _empty_usage() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "month": pd.Series(dtype="datetime64[ns]"),
            "tool_clean": pd.Series(dtype="object"),
            "Team": pd.Series(dtype="object"),
            "runs": pd.Series(dtype="int64"),
        }
    )


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def load_usage_history(_sheet) -> pd.DataFrame:
    """Monthly run counts from "Sheet1" (Tool, Month, Usage Count, Team…).

    Sheet1 goes back to Oct-2025, before the Run Log existed, so it is the
    source for every month-by-month number on the dashboard.
    """
    values = _sheet.get_all_values()
    if len(values) < 2:
        return _empty_usage()
    rows = [(row + [""] * 4)[:4] for row in values[1:]]
    df = pd.DataFrame(rows, columns=["Tool", "Month", "Usage Count", "Team"])
    df = df[df["Tool"].astype(str).str.strip() != ""]
    df["month"] = pd.to_datetime(df["Month"].astype(str).str.strip(), format="%b-%Y", errors="coerce")
    df["runs"] = pd.to_numeric(df["Usage Count"], errors="coerce").fillna(0).astype(int)
    df = df.dropna(subset=["month"])
    if df.empty:
        return _empty_usage()
    df["Team"] = df["Team"].astype(str).str.strip().str.title()
    df["tool_clean"] = df["Tool"].map(normalize_tool_name)
    return df[["month", "tool_clean", "Team", "runs"]].reset_index(drop=True)


def usage_from_run_log(run_log_df: pd.DataFrame) -> pd.DataFrame:
    """Fallback when Sheet1 can't be read: monthly counts from the Run Log."""
    if run_log_df.empty:
        return _empty_usage()
    df = run_log_df.assign(month=run_log_df["Timestamp"].dt.to_period("M").dt.to_timestamp())
    return (
        df.groupby(["month", "tool_clean", "Team"]).size().rename("runs").reset_index()
    )


def minutes_saved_per_tool(catalog_df) -> dict:
    """Minutes saved per run for each canonical tool (first catalog row wins)."""
    rates = {}
    if catalog_df is None or catalog_df.empty:
        return rates
    for _, row in catalog_df.iterrows():
        per_run = row["Time Saved per Run (min)"]
        if pd.notna(per_run) and row["tool_clean"] not in rates:
            rates[row["tool_clean"]] = float(per_run)
    return rates


def compute_management(usage_df: pd.DataFrame, run_log_df: pd.DataFrame, rates: dict, now: datetime) -> dict:
    df = usage_df.copy()
    df["hours"] = df["tool_clean"].map(rates).fillna(0) * df["runs"] / 60

    this_month = pd.Timestamp(now.year, now.month, 1)
    last_month = this_month - pd.DateOffset(months=1)
    year_start = pd.Timestamp(now.year, 1, 1)
    ytd = df[df["month"] >= year_start]

    def hours(mask):
        return float(df.loc[mask, "hours"].sum())

    in_use = 0
    if not run_log_df.empty:
        recent = run_log_df["Timestamp"] >= pd.Timestamp(now) - pd.Timedelta(days=60)
        in_use = int(run_log_df.loc[recent, "tool_clean"].nunique())

    unrated = sorted(
        t for t in df.loc[df["month"] >= year_start, "tool_clean"].unique() if t not in rates
    )
    return {
        "df": df,
        "hours_ytd": float(ytd["hours"].sum()),
        "runs_ytd": int(ytd["runs"].sum()),
        "hours_month": hours(df["month"] == this_month),
        "hours_last_month": hours(df["month"] == last_month),
        "last_month_label": last_month.strftime("%B"),
        "quotes_ytd": int(ytd.loc[ytd["tool_clean"].isin(QUOTATION_TOOLS), "runs"].sum()),
        "quotes_month": int(
            df.loc[(df["month"] == this_month) & df["tool_clean"].isin(QUOTATION_TOOLS), "runs"].sum()
        ),
        "tools_in_use": in_use,
        "unrated": unrated,
    }


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
            bgcolor="#1c2636",
            bordercolor="rgba(255,255,255,0.12)",
            font=dict(color=PALETTE["white"], family="Google Sans, Segoe UI, sans-serif"),
        ),
        showlegend=False,
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, linecolor="rgba(255,255,255,0.10)")
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


def build_monthly_hours_chart(mgmt_df: pd.DataFrame, now: datetime, compact: bool = False) -> go.Figure:
    df = mgmt_df[mgmt_df["month"] <= pd.Timestamp(now)]
    months = pd.date_range(df["month"].min(), pd.Timestamp(now.year, now.month, 1), freq="MS") if not df.empty else []
    pivot = (
        df.groupby(["month", "Team"])["hours"].sum().unstack(fill_value=0).reindex(months, fill_value=0)
        if not df.empty
        else pd.DataFrame()
    )
    teams = [t for t in TEAM_ORDER if t in pivot.columns and pivot[t].sum() > 0] + [
        t for t in pivot.columns if t not in TEAM_ORDER and pivot[t].sum() > 0
    ]
    labels = [m.strftime("%b %y") for m in months]
    fig = go.Figure()
    for team in teams:
        fig.add_trace(
            go.Bar(
                x=labels,
                y=pivot[team].round(0),
                name=team,
                marker=dict(color=TEAM_COLORS.get(team, TEAM_FALLBACK_COLOR), line=dict(color=PALETTE["card"], width=1)),
                hovertemplate=f"{team}: <b>%{{y:,.0f}}</b> hrs<extra></extra>",
            )
        )
    totals = pivot[teams].sum(axis=1) if teams else []
    if len(labels):
        fig.add_trace(
            go.Scatter(
                x=labels,
                y=list(totals),
                mode="text",
                text=[f"{v:,.0f}" if v else "" for v in totals],
                textposition="top center",
                textfont=dict(color=PALETTE["dim"], size=11),
                hoverinfo="skip",
                showlegend=False,
            )
        )
    _apply_dark(fig, "Hours saved per month, by team")
    fig.update_layout(
        barmode="stack",
        bargap=0.3,
        height=260 if compact else 340,
        hovermode="x unified",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=1, xanchor="right", font=dict(color=PALETTE["dim"], size=12)),
    )
    fig.update_xaxes(showgrid=False)
    peak = float(max(totals)) if len(labels) and len(totals) else 0
    fig.update_yaxes(range=[0, peak * 1.15 if peak else 1])
    return fig


def build_top_tools_hours_chart(mgmt_df: pd.DataFrame, now: datetime, compact: bool = False) -> go.Figure:
    ytd = mgmt_df[mgmt_df["month"] >= pd.Timestamp(now.year, 1, 1)]
    grouped = ytd.groupby("tool_clean").agg(hours=("hours", "sum"), runs=("runs", "sum"))
    top = grouped[grouped["hours"] > 0].sort_values("hours").tail(8)
    fig = go.Figure(
        go.Bar(
            x=top["hours"].round(0),
            y=top.index,
            orientation="h",
            marker=dict(color=PALETTE["steel"]),
            text=[f"{v:,.0f} hrs" for v in top["hours"]],
            textposition="outside",
            textfont=dict(color=PALETTE["white"], size=12),
            cliponaxis=False,
            customdata=top["runs"],
            hovertemplate="%{y}: <b>%{x:,.0f}</b> hrs · %{customdata:,} runs<extra></extra>",
        )
    )
    _apply_dark(fig, f"Top tools by hours saved — {now.year}")
    fig.update_layout(height=260 if compact else 340, bargap=0.35, barcornerradius=4, margin=dict(t=44, l=8, r=60, b=8))
    fig.update_yaxes(showgrid=False, tickfont=dict(color=PALETTE["dim"], size=12))
    return fig


def build_runs_trend_chart(df: pd.DataFrame, now: datetime, compact: bool = False) -> go.Figure:
    """Month-to-date runs, cumulative. Hover shows that day's runs and the running total."""
    month_start = pd.Timestamp(now.year, now.month, 1)
    idx = pd.date_range(month_start, pd.Timestamp(now).normalize(), freq="D")
    month_df = df[df["date"] >= month_start] if not df.empty else df
    daily = month_df.groupby("date").size().reindex(idx, fill_value=0) if not month_df.empty else pd.Series(0, index=idx)
    cumulative = daily.cumsum()
    x = [month_start - pd.Timedelta(hours=12)] + list(idx)
    y = [0] + list(cumulative.values)
    custom = [[0, 0]] + [[int(d), int(c)] for d, c in zip(daily.values, cumulative.values)]
    fig = go.Figure(
        go.Scatter(
            x=x,
            y=y,
            mode="lines",
            line=dict(color=PALETTE["glow"], width=3, shape="spline", smoothing=0.5),
            fill="tozeroy",
            fillcolor="rgba(141,184,224,0.10)",
            customdata=custom,
            hovertemplate="<b>%{x|%d %b}</b><br>%{customdata[0]} runs that day"
            "<br>%{customdata[1]:,} month to date<extra></extra>",
        )
    )
    total = int(cumulative.iloc[-1]) if len(cumulative) else 0
    _apply_dark(fig, f"Runs this month — {total:,} so far")
    fig.update_layout(height=260 if compact else 340, hovermode="x unified")
    fig.update_xaxes(tickformat="%d %b", range=[month_start - pd.Timedelta(hours=12), pd.Timestamp(now).normalize() + pd.Timedelta(hours=12)])
    fig.update_yaxes(rangemode="tozero", title=dict(text="Cumulative runs", font=dict(color=PALETTE["dim"], size=11)))
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


def _time_saved_hero(total_hours: float, year: int, runs: int, months: int = 12) -> str:
    fte = total_hours / HOURS_PER_FTE_MONTH / max(1, months)
    return (
        '<div class="mw-hero">'
        f'<div class="mw-hero-label">⚡ Time saved by automation in {year}</div>'
        f'<div class="mw-hero-value">{total_hours:,.0f}<span> hours</span></div>'
        f'<div class="mw-kpi-sub">{runs:,} automated runs · ≈ {fte:,.1f} full-time employees every month</div>'
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
        per_run_str = f'{r["PerRun"]:g} min' if pd.notna(r["PerRun"]) else "—"
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
            background: #0f1724 !important;
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
            background: #141d2b;
            border: 1px solid rgba(255, 255, 255, 0.07);
            border-radius: 14px;
        }}
        .mw-hero-label {{
            font-size: 0.8rem; letter-spacing: 2px; text-transform: uppercase;
            color: {PALETTE["dim"]}; margin-bottom: 0.3rem;
        }}
        .mw-hero-value {{
            font-family: 'Google Sans', 'Segoe UI', sans-serif;
            font-size: clamp(2.3rem, 5vw, 3.8rem); font-weight: 600;
            color: #9fd3c3;
        }}
        .mw-hero-value span {{
            font-size: 1.1rem; font-weight: 500; color: {PALETTE["dim"]};
            margin-left: 0.4rem;
        }}
        .mw-section-title {{
            font-size: 0.95rem; letter-spacing: 1px; text-transform: uppercase;
            color: {PALETTE["dim"]}; margin: 1.2rem 0 0.6rem;
        }}
        .mw-kpi {{
            background: #141d2b;
            border: 1px solid rgba(255, 255, 255, 0.07);
            border-radius: 14px;
            padding: 1.1rem 1.3rem 1rem;
        }}
        .mw-kpi-label {{
            font-size: 0.75rem; text-transform: uppercase; letter-spacing: 2px;
            color: {PALETTE["dim"]}; margin-bottom: 0.35rem;
        }}
        .mw-kpi-value {{
            font-family: 'Google Sans', 'Segoe UI', sans-serif;
            font-size: clamp(2rem, 4vw, 3.4rem); font-weight: 600; line-height: 1.05;
            color: {PALETTE["white"]};
        }}
        .mw-kpi-sub {{ font-size: 0.8rem; color: {PALETTE["dim"]}; margin-top: 0.3rem; }}
        .mw-empty {{
            border: 1px dashed rgba(255, 255, 255, 0.08); border-radius: 18px;
            padding: 3rem 1rem; text-align: center; color: {PALETTE["dim"]};
        }}
        .mw-banner {{
            background: #141d2b;
            border: 1px solid rgba(255, 255, 255, 0.08); border-radius: 12px;
            padding: 1rem 1.2rem; color: {PALETTE["dim"]};
        }}
        .mw-table-wrap {{
            overflow-x: auto;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 14px;
        }}
        .mw-table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
        .mw-table th {{
            text-align: left; padding: 0.6rem 0.9rem; color: {PALETTE["dim"]};
            font-size: 0.7rem; text-transform: uppercase; letter-spacing: 1px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.08); white-space: nowrap;
        }}
        .mw-table td {{
            padding: 0.55rem 0.9rem; border-bottom: 1px solid rgba(255, 255, 255, 0.08);
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
            background: #141d2b !important;
            color: {PALETTE["white"]} !important;
            border: 1px solid rgba(255, 255, 255, 0.08) !important;
            box-shadow: none !important;
        }}
        {tv_extra}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _dashboard_body(run_log_sheet, catalog_sheet, tv_mode: bool, usage_sheet=None) -> None:
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

    usage_df = None
    if usage_sheet is not None:
        try:
            usage_df = load_usage_history(usage_sheet)
        except APIError:
            usage_df = None
    if usage_df is None or usage_df.empty:
        usage_df = usage_from_run_log(df)

    rates = minutes_saved_per_tool(catalog_df)
    mgmt = compute_management(usage_df, df, rates, now)

    st.markdown(_time_saved_hero(mgmt["hours_ytd"], now.year, mgmt["runs_ytd"], now.month), unsafe_allow_html=True)

    tiles = st.columns(4)
    tiles[0].markdown(
        _kpi_tile(
            "Hours saved this month",
            f"{mgmt['hours_month']:,.0f}",
            f"{now.strftime('%B')} so far · ≈ {mgmt['hours_month'] / HOURS_PER_FTE_MONTH:,.1f} FTE",
        ),
        unsafe_allow_html=True,
    )
    tiles[1].markdown(
        _kpi_tile(
            f"Hours saved in {mgmt['last_month_label']}",
            f"{mgmt['hours_last_month']:,.0f}",
            f"≈ {mgmt['hours_last_month'] / HOURS_PER_FTE_MONTH:,.1f} full-time employees",
        ),
        unsafe_allow_html=True,
    )
    tiles[2].markdown(
        _kpi_tile(
            "Documents processed",
            f"{mgmt['runs_ytd']:,}",
            f"{now.year} · {mgmt['quotes_ytd']:,} of them quotations",
        ),
        unsafe_allow_html=True,
    )
    tiles[3].markdown(
        _kpi_tile("Tools in use", f"{mgmt['tools_in_use']:,}", "used in the last 60 days"),
        unsafe_allow_html=True,
    )

    st.markdown(
        f"<div style='height:{'0.4rem' if tv_mode else '0.75rem'}'></div>",
        unsafe_allow_html=True,
    )

    if mgmt["df"].empty:
        st.markdown(
            '<div class="mw-empty">No runs logged yet — numbers will appear here '
            "as soon as the team starts using the tools.</div>",
            unsafe_allow_html=True,
        )
    else:
        st.plotly_chart(
            build_monthly_hours_chart(mgmt["df"], now, compact=tv_mode),
            width="stretch",
            config={"displayModeBar": False},
        )
        col_tools, col_trend = st.columns([3, 2])
        with col_tools:
            st.plotly_chart(
                build_top_tools_hours_chart(mgmt["df"], now, compact=tv_mode),
                width="stretch",
                config={"displayModeBar": False},
            )
        with col_trend:
            st.plotly_chart(
                build_runs_trend_chart(df, now, compact=tv_mode),
                width="stretch",
                config={"displayModeBar": False},
            )

    # The catalog table is detail meant to be read at a desk, not from across a
    # room — TV mode stops at the numbers and charts above.
    if catalog_sheet is not None and not tv_mode:
        if catalog_df is None:
            pass
        elif catalog_df.empty:
            st.markdown(
                '<div class="mw-empty">No tools in the catalog yet — add rows to the '
                '"Tool Catalog" sheet tab (Tool, Owner Team, Summary, Status, Go-Live Date, '
                "Manual/Automated Time, Runs per Month) to see time-saved estimates here.</div>",
                unsafe_allow_html=True,
            )
        else:
            ts = compute_time_saved(catalog_df, df, now)
            st.markdown(
                '<div class="mw-section-title">Automation Catalog</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                build_catalog_table_html(ts["rows"]),
                unsafe_allow_html=True,
            )
            if mgmt["unrated"]:
                st.caption(
                    "Not counted in hours saved yet (no minutes in the Tool Catalog): "
                    + ", ".join(mgmt["unrated"])
                )


def render_dashboard(run_log_sheet, catalog_sheet=None, tv_mode: bool = False, usage_sheet=None) -> None:
    _inject_css(tv_mode)
    if tv_mode:
        st.fragment(run_every=TV_REFRESH_SECONDS)(_dashboard_body)(
            run_log_sheet, catalog_sheet, tv_mode, usage_sheet
        )
    else:
        _dashboard_body(run_log_sheet, catalog_sheet, tv_mode, usage_sheet)
