# aap.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import altair as alt
from io import BytesIO

# ---------- Page config ----------
st.set_page_config(page_title="TFC – R1–R3 Dashboard", layout="wide")

st.title("The Fresh Connection – Rounds 1–3 Dashboard")
st.caption("Sales • Supply Chain • Purchasing • Operations (metrics & visuals)")

# ---------- Helpers ----------
@st.cache_data
def load_workbook(path: str) -> dict:
    """Load all sheets; return dict of DataFrames keyed by sheet name."""
    dfs = pd.read_excel(path, sheet_name=None, engine="openpyxl")
    # Clean each sheet
    cleaned = {}
    for name, df in dfs.items():
        df = df.copy()
        # Normalize column names (strip)
        df.columns = [c.strip() for c in df.columns]
        # Ensure a Round column exists
        if "Round" not in df.columns:
            raise ValueError(f"Sheet '{name}' must contain a 'Round' column.")
        # Clean numeric cells (handle %, €, commas)
        for col in df.columns:
            if col == "Round":
                continue
            # Convert to numeric: remove percent/currency/commas/spaces
            df[col] = (
                df[col]
                .astype(str)
                .str.replace("%", "", regex=False)
                .str.replace("€", "", regex=False)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")
        cleaned[name] = df
    return cleaned

def is_percent(colname: str) -> bool:
    tokens = ["%", "service", "availability", "reliab", "rejection", "obsolete",
              "utilization", "adherence", "cost", "shelf", "osa"]
    col = colname.lower()
    return any(t in col for t in tokens)

def is_currency(colname: str) -> bool:
    col = colname.lower()
    return "gross margin" in col or "€" in col or "margin" in col

def fmt_value(colname: str, val: float) -> str:
    if pd.isna(val):
        return "–"
    if is_currency(colname):
        return f"€{val:,.0f}"
    if is_percent(colname):
        return f"{val:.1f}%"
    return f"{val:,.2f}"

def metric_cards(df: pd.DataFrame, round_sel: int):
    """Show metrics as cards with delta vs previous round."""
    cols = [c for c in df.columns if c != "Round"]
    row = df.loc[df["Round"] == round_sel]
    prev_row = df.loc[df["Round"] == (round_sel - 1)]
    if row.empty:
        st.info(f"No data for Round {round_sel}")
        return
    row = row.iloc[0]
    prev = prev_row.iloc[0] if not prev_row.empty else None

    # Layout in 3 columns per row
    n = len(cols)
    per_row = 3
    for start in range(0, n, per_row):
        cset = st.columns(per_row)
        for idx, colname in enumerate(cols[start:start+per_row]):
            val = row[colname]
            delta = None
            if prev is not None and pd.notna(prev[colname]) and pd.notna(val):
                diff = val - prev[colname]
                if is_percent(colname):
                    delta = f"{diff:+.1f}%"
                elif is_currency(colname):
                    delta = f"{diff:+,.0f}"
                else:
                    delta = f"{diff:+.2f}"
            with cset[idx]:
                st.metric(label=colname, value=fmt_value(colname, val), delta=delta)

def line_chart_each_metric(df: pd.DataFrame, title_prefix: str):
    """One small line per metric to avoid mixed scales."""
    metrics = [c for c in df.columns if c != "Round"]
    for i in range(0, len(metrics), 2):
        c1, c2 = st.columns(2)
        for col, container in zip(metrics[i:i+2], (c1, c2)):
            with container:
                melted = df[["Round", col]].rename(columns={col: "Value"})
                # Decide y-axis formatting
                yaxis_title = "%"
                if is_currency(col):
                    yaxis_title = "€"
                elif not is_percent(col):
                    yaxis_title = "Value"
                fig = px.line(
                    melted,
                    x="Round",
                    y="Value",
                    markers=True,
                    title=f"{title_prefix} – {col}",
                )
                fig.update_layout(hovermode="x unified",
                                  xaxis=dict(dtick=1),
                                  yaxis_title=yaxis_title,
                                  margin=dict(l=10, r=10, t=50, b=10))
                st.plotly_chart(fig, use_container_width=True)

def download_df(df: pd.DataFrame, label: str):
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label=f"Download {label} CSV",
        data=csv,
        file_name=f"{label.replace(' ', '_').lower()}.csv",
        mime="text/csv",
        use_container_width=True,
    )

# ---------- Load data ----------
try:
    book = load_workbook("metrics.xlsx")
except FileNotFoundError:
    st.error("`metrics.xlsx` not found in the repository. Please add it to the repo root.")
    st.stop()
except Exception as e:
    st.exception(e)
    st.stop()

# Expected sheet keys
expected_sheets = ["Sales", "SupplyChain", "Purchasing", "Operations"]
missing = [s for s in expected_sheets if s not in book]
if missing:
    st.warning(f"Missing sheets in `metrics.xlsx`: {', '.join(missing)}")
    st.write("Found sheets:", list(book.keys()))

# Sidebar controls
with st.sidebar:
    st.header("Controls")
    round_sel = st.slider("Select Round", min_value=1, max_value=3, value=3, step=1)
    st.caption("Metric cards show deltas vs. previous round.")

# Overview – quick key charts (one metric per function)
st.subheader("Overview")
overview_cols = st.columns(4)
try:
    with overview_cols[0]:
        df = book["Sales"][["Round", "ROI (%)"]].dropna()
        fig = px.line(df, x="Round", y="ROI (%)", markers=True, title="Sales – ROI (%)")
        fig.update_layout(xaxis=dict(dtick=1), hovermode="x unified", margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with overview_cols[1]:
        df = book["SupplyChain"][["Round", "Availability components (%)"]].dropna()
        fig = px.line(df, x="Round", y="Availability components (%)", markers=True, title="SC – Availability components (%)")
        fig.update_layout(xaxis=dict(dtick=1), hovermode="x unified", margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with overview_cols[2]:
        df = book["Operations"][["Round", "Production plan adherence (%)"]].dropna()
        fig = px.line(df, x="Round", y="Production plan adherence (%)", markers=True, title="Ops – Production plan adherence (%)")
        fig.update_layout(xaxis=dict(dtick=1), hovermode="x unified", margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with overview_cols[3]:
        df = book["Purchasing"][["Round", "Delivery reliability suppliers (%)"]].dropna()
        fig = px.line(df, x="Round", y="Delivery reliability suppliers (%)", markers=True, title="Purch – Delivery reliability (%)")
        fig.update_layout(xaxis=dict(dtick=1), hovermode="x unified", margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
except Exception:
    st.info("If an overview chart is blank, ensure the corresponding column exists in metrics.xlsx.")

st.markdown("---")

# Tabs for each function
tabs = st.tabs(["📈 Sales", "🚚 Supply Chain", "🏭 Operations", "🛒 Purchasing"])

# ----- Sales tab -----
with tabs[0]:
    st.header("Sales")
    df = book.get("Sales")
    if df is not None:
        st.caption("Metric cards (with Δ vs previous round)")
        metric_cards(df, round_sel)

        st.markdown("#### Visualizations")
        line_chart_each_metric(df, "Sales")

        st.markdown("#### Data (Rounds 1–3)")
        st.dataframe(df, use_container_width=True)
        download_df(df, "Sales")
    else:
        st.warning("Sales sheet not found.")

# ----- Supply Chain tab -----
with tabs[1]:
    st.header("Supply Chain")
    df = book.get("SupplyChain")
    if df is not None:
        st.caption("Metric cards (with Δ vs previous round)")
        metric_cards(df, round_sel)

        st.markdown("#### Visualizations")
        line_chart_each_metric(df, "Supply Chain")

        st.markdown("#### Data (Rounds 1–3)")
        st.dataframe(df, use_container_width=True)
        download_df(df, "SupplyChain")
    else:
        st.warning("SupplyChain sheet not found.")

# ----- Operations tab -----
with tabs[2]:
    st.header("Operations")
    df = book.get("Operations")
    if df is not None:
        st.caption("Metric cards (with Δ vs previous round)")
        metric_cards(df, round_sel)

        st.markdown("#### Visualizations")
        line_chart_each_metric(df, "Operations")

        st.markdown("#### Data (Rounds 1–3)")
        st.dataframe(df, use_container_width=True)
        download_df(df, "Operations")
    else:
        st.warning("Operations sheet not found.")

# ----- Purchasing tab -----
with tabs[3]:
    st.header("Purchasing")
    df = book.get("Purchasing")
    if df is not None:
        st.caption("Metric cards (with Δ vs previous round)")
        metric_cards(df, round_sel)

        st.markdown("#### Visualizations")
        line_chart_each_metric(df, "Purchasing")

        st.markdown("#### Data (Rounds 1–3)")
        st.dataframe(df, use_container_width=True)
        download_df(df, "Purchasing")
    else:
        st.warning("Purchasing sheet not found.")

st.markdown("---")
st.caption("Tip: If a metric doesn’t appear, check its exact column name in `metrics.xlsx` (case/spacing must match).")
