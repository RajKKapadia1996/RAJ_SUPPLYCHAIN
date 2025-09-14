# app.py
# The Fresh Connection – Dashboard (R1 & R2)
# Robust KPI visuals + quick explorers for the other areas

import os
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

# --------------------- Page ---------------------
st.set_page_config(
    page_title="The Fresh Connection – Dashboard (R1 & R2)",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.title("The Fresh Connection – Dashboard (R1 & R2)")

DATA_FILE = "Dashboard_Metrics_R1_R2_only.xlsx"

# --------------------- Utils ---------------------
PERCENT_LIKE = re.compile(r"%")
CURRENCY_LIKE = re.compile(r"[€$,]")  # remove currency & thousands
PAREN_LIKE = re.compile(r"^\((.*)\)$")  # (1.23) -> -1.23

def numify(x) -> Optional[float]:
    """Convert strings like '€ 1,388', '-2.8%', '(1,234)' into float."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s == "":
        return None
    # negative in parentheses
    m = PAREN_LIKE.match(s)
    if m:
        inner = m.group(1)
        s = "-" + inner
    # remove currency and thousand separators
    s = CURRENCY_LIKE.sub("", s)
    s = s.replace(" ", "")
    # percent
    if PERCENT_LIKE.search(s):
        s = s.replace("%", "")
        try:
            return float(s)
        except Exception:
            return None
    # plain float
    try:
        return float(s)
    except Exception:
        return None

def norm_round_token(x) -> Optional[str]:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip().upper()
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    return f"R{int(m.group(1))}"

def find_sheet_name(sheets: Dict[str, pd.DataFrame], keywords: List[str]) -> Optional[str]:
    keys = [k.lower() for k in keywords]
    for name in sheets:
        low = name.lower()
        if all(k in low for k in keys):
            return name
    return None

# --------------------- Load ---------------------
@st.cache_data(show_spinner=True)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    if not os.path.exists(path):
        st.error(f"Data file not found in repo: `{path}`")
        return {}
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")

@st.cache_data(show_spinner=False)
def list_rounds(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    found = set()
    for name in sheets:
        tok = norm_round_token(name)
        if tok:
            found.add(tok)
    for df in sheets.values():
        if isinstance(df, pd.DataFrame) and "Round" in df.columns:
            for v in df["Round"].unique():
                tok = norm_round_token(v)
                if tok:
                    found.add(tok)
    if not found:
        return ["R1", "R2"]
    return sorted(found, key=lambda t: int(re.search(r"\d+", t).group(0)))

# --------------------- KPI extraction ---------------------
def _kpi_from_two_cols(df: pd.DataFrame, round_token: str) -> Optional[pd.DataFrame]:
    """Expect columns like Metric / Value (or Amount/Score)."""
    candidates_metric = [c for c in df.columns if str(c).strip().lower() in ["metric", "kpi", "name"]]
    candidates_value  = [c for c in df.columns if str(c).strip().lower() in ["value", "amount", "score"]]
    if not (candidates_metric and candidates_value):
        return None
    mcol = candidates_metric[0]
    vcol = candidates_value[0]
    out = df[[mcol, vcol]].copy()
    out.columns = ["Metric", "Value"]
    out["Value"] = out["Value"].map(numify)
    out["Round"] = round_token
    out = out.dropna(subset=["Value"])
    return out

def _kpi_from_wide_row(df: pd.DataFrame, round_token: str) -> Optional[pd.DataFrame]:
    """
    Handle a sheet where metrics are columns in a single row.
    """
    if df.empty:
        return None
    row = df.iloc[0]
    items = []
    for c in df.columns:
        if str(c).strip().lower() == "round":  # skip
            continue
        val = numify(row[c])
        if val is not None:
            items.append({"Metric": str(c), "Value": val, "Round": round_token})
    if not items:
        return None
    return pd.DataFrame(items)

def extract_kpi_from_sheet(name: str, df: pd.DataFrame, round_token: str) -> Optional[pd.DataFrame]:
    # Try explicit metric/value
    t = _kpi_from_two_cols(df, round_token)
    if t is not None and not t.empty:
        return t
    # Try wide single-row
    t = _kpi_from_wide_row(df, round_token)
    if t is not None and not t.empty:
        return t
    # Try long with Round column: melt numeric columns
    if "Round" in df.columns:
        sub = df.copy()
        sub["__R__"] = sub["Round"].map(norm_round_token)
        sub = sub[sub["__R__"] == round_token]
        num_cols = [c for c in sub.columns if c not in ["Round", "__R__"]]
        melted = []
        for c in num_cols:
            vals = sub[c].apply(numify)
            vals = vals.dropna()
            if not vals.empty:
                # use mean if multiple rows
                melted.append({"Metric": str(c), "Value": float(vals.mean()), "Round": round_token})
        if melted:
            return pd.DataFrame(melted)
    return None

@st.cache_data(show_spinner=True)
def build_kpi_long(sheets: Dict[str, pd.DataFrame], rounds: List[str]) -> pd.DataFrame:
    """
    Assemble KPI long table with columns: Metric, Value, Round
    Looks for 'kpi' sheets (optionally with _R1/_R2) or a single KPI sheet with a Round column.
    """
    out = []
    # Prefer dedicated KPI sheet(s)
    kpi_sheet_any = find_sheet_name(sheets, ["kpi"])
    if kpi_sheet_any:
        df0 = sheets[kpi_sheet_any]
        if "Round" in df0.columns:
            # One KPI sheet with Round column
            for r in rounds:
                t = extract_kpi_from_sheet(kpi_sheet_any, df0, r)
                if t is not None:
                    out.append(t)
        else:
            # Try per-round KPI_x sheets (KPI_R1, KPI R2, etc.)
            for r in rounds:
                candidate = find_sheet_name(sheets, ["kpi", r])
                if candidate:
                    t = extract_kpi_from_sheet(candidate, sheets[candidate], r)
                    if t is not None:
                        out.append(t)
    else:
        # fallback: scan all sheets
        for name, df in sheets.items():
            for r in rounds:
                if r in name.upper():
                    t = extract_kpi_from_sheet(name, df, r)
                    if t is not None:
                        out.append(t)

    if not out:
        return pd.DataFrame(columns=["Metric", "Value", "Round"])
    kpi = pd.concat(out, ignore_index=True)
    # Clean metric names a bit
    kpi["Metric"] = kpi["Metric"].astype(str).str.strip()
    # Common aliases normalization (optional)
    aliases = {
        "ROI": ["ROI", "ROI (%)", "Return on investment"],
        "Gross margin (customer)": ["Gross margin (customer)", "Gross margin"],
        "Obsolete products (%)": ["Obsolete products (%)", "Obsoletes (%)", "Obsolete (%)"],
        "Service level outbound order lines": [
            "Service level outbound order lines",
            "Service level (order lines)",
            "Service level order lines",
        ],
    }
    rev = {}
    for std, al in aliases.items():
        for a in al:
            rev[a.lower()] = std
    kpi["Metric"] = kpi["Metric"].apply(lambda s: rev.get(s.lower(), s))
    return kpi

# --------------------- Generic views for other areas ---------------------
def strip_unnamed_cols(df: pd.DataFrame) -> pd.DataFrame:
    drop_cols = [c for c in df.columns if str(c).lower().startswith("unnamed")]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df

def get_area_df(sheets: Dict[str, pd.DataFrame], area_words: List[str], round_sel: str) -> Optional[pd.DataFrame]:
    # prefer a round-specific sheet
    name = find_sheet_name(sheets, area_words + [round_sel])
    if name:
        return strip_unnamed_cols(sheets[name].copy())
    # generic sheet, try filter by Round column
    name = find_sheet_name(sheets, area_words)
    if name:
        df = sheets[name].copy()
        if "Round" in df.columns:
            df = df[df["Round"].map(norm_round_token) == norm_round_token(round_sel)]
        return strip_unnamed_cols(df)
    return None

def quick_chart(df: pd.DataFrame, title: str):
    """A very light helper to plot first categorical vs first numeric."""
    if df is None or df.empty:
        st.info("No data available.")
        return
    # find a categorical
    cat_cols = [c for c in df.columns if df[c].dtype == "object"]
    # find a numeric (try to coerce strings using numify)
    num_cols = []
    for c in df.columns:
        if c in cat_cols:
            continue
        ser = pd.to_numeric(df[c], errors="coerce")
        if ser.notna().sum() == 0:
            ser = df[c].map(numify)
        if ser.notna().sum() > 0:
            num_cols.append(c)
    if cat_cols and num_cols:
        xcol, ycol = cat_cols[0], num_cols[0]
        tmp = df[[xcol, ycol]].copy()
        tmp[ycol] = pd.to_numeric(tmp[ycol], errors="coerce").fillna(tmp[ycol].map(numify))
        tmp = tmp.dropna(subset=[ycol])
        if not tmp.empty:
            fig = px.bar(tmp, x=xcol, y=ycol, title=title, text_auto=".2s")
            fig.update_layout(xaxis_title="", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

# --------------------- App UI ---------------------
with st.sidebar:
    st.header("Controls")
    sheets = load_workbook(DATA_FILE)
    rounds = list_rounds(sheets)
    round_sel = st.radio("Select Round", rounds, horizontal=True, index=0)
    st.caption(f"Data file: `{DATA_FILE}`")

# Build KPI long once (for overview across rounds)
kpi_long = build_kpi_long(sheets, rounds)

tabs = st.tabs([
    "KPI Overview",
    "Purchasing", "Operations", "Sales", "Supply Chain", "Finance",
    "Graph Builder"
])

# ---- KPI Overview ----
with tabs[0]:
    st.subheader("KPIs by Round (cleaned & normalized)")
    if kpi_long.empty:
        st.warning("No KPI data found. Check that your Excel has a KPI sheet (or KPI_R1 / KPI_R2).")
    else:
        # Show a tidy table (pivoted)
        pivot = kpi_long.pivot_table(index="Metric", columns="Round", values="Value", aggfunc="mean")
        st.dataframe(pivot, use_container_width=True)

        # Multi-metric line across rounds
        st.markdown("### Trend by Round (all KPIs)")
        fig = px.line(kpi_long, x="Round", y="Value", color="Metric", markers=True)
        fig.update_layout(yaxis_title="", xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

        # Small multiples (bar per metric)
        st.markdown("### KPI comparison (one tile per KPI)")
        metrics = sorted(kpi_long["Metric"].unique())
        pick = st.multiselect("Select KPIs to display", metrics, default=metrics)
        sm = kpi_long[kpi_long["Metric"].isin(pick)]
        if not sm.empty:
            fig2 = px.bar(
                sm, x="Round", y="Value", facet_col="Metric", facet_col_wrap=3,
                text_auto=".2s", color="Round"
            )
            fig2.update_layout(showlegend=False, yaxis_title="", xaxis_title="")
            st.plotly_chart(fig2, use_container_width=True)
        st.caption("Note: ROI and percentage KPIs are parsed from strings like '€ 1,388' or '-2.8%'. If something still looks off, check the raw Excel values on the KPI sheet(s).")

# ---- Purchasing ----
with tabs[1]:
    st.subheader("Purchasing")
    dfp = get_area_df(sheets, ["purchasing"], round_sel)
    if dfp is not None and not dfp.empty:
        st.dataframe(dfp, use_container_width=True)
        quick_chart(dfp, "Purchasing – first categorical vs first numeric")
    else:
        st.info("No Purchasing data for the selected round.")

# ---- Operations ----
with tabs[2]:
    st.subheader("Operations")
    dfo = get_area_df(sheets, ["operations"], round_sel)
    if dfo is not None and not dfo.empty:
        st.dataframe(dfo, use_container_width=True)
        quick_chart(dfo, "Operations – first categorical vs first numeric")
    else:
        st.info("No Operations data for the selected round.")

# ---- Sales ----
with tabs[3]:
    st.subheader("Sales")
    dfs = get_area_df(sheets, ["sales"], round_sel)
    if dfs is not None and not dfs.empty:
        st.dataframe(dfs, use_container_width=True)
        quick_chart(dfs, "Sales – first categorical vs first numeric")
    else:
        st.info("No Sales data for the selected round.")

# ---- Supply Chain ----
with tabs[4]:
    st.subheader("Supply Chain")
    dfsc = get_area_df(sheets, ["supply", "chain"], round_sel)
    if dfsc is not None and not dfsc.empty:
        st.dataframe(dfsc, use_container_width=True)
        quick_chart(dfsc, "Supply Chain – first categorical vs first numeric")
    else:
        st.info("No Supply Chain data for the selected round.")

# ---- Finance ----
with tabs[5]:
    st.subheader("Finance")
    dff = get_area_df(sheets, ["finance"], round_sel)
    if dff is not None and not dff.empty:
        st.dataframe(dff, use_container_width=True)
        quick_chart(dff, "Finance – first categorical vs first numeric")
    else:
        st.info("No Finance data for the selected round.")

# ---- Graph Builder ----
with tabs[6]:
    st.subheader("Build your own chart from any sheet")
    if not sheets:
        st.info("Workbook not loaded.")
    else:
        all_names = list(sheets.keys())
        sheet_pick = st.selectbox("Sheet", all_names)
        df_raw = sheets[sheet_pick].copy()
        df_raw = strip_unnamed_cols(df_raw)

        # Optional round filter if a Round column exists
        if "Round" in df_raw.columns:
            rd_opts = sorted(df_raw["Round"].dropna().astype(str).unique())
            rd_pick = st.multiselect("Filter Round (optional)", rd_opts, default=rd_opts)
            df_raw = df_raw[df_raw["Round"].astype(str).isin(rd_pick)]

        st.dataframe(df_raw, use_container_width=True)

        # Pick columns for a quick plot
        cat_cols = [c for c in df_raw.columns if df_raw[c].dtype == "object"]
        num_cols = []
        for c in df_raw.columns:
            ser = pd.to_numeric(df_raw[c], errors="coerce")
            if ser.notna().sum() == 0:
                ser = df_raw[c].map(numify)
            if ser.notna().sum() > 0:
                num_cols.append(c)

        c1, c2, c3 = st.columns(3)
        with c1:
            x_pick = st.selectbox("X (categorical)", cat_cols if cat_cols else [None])
        with c2:
            y_pick = st.selectbox("Y (numeric)", num_cols if num_cols else [None])
        with c3:
            chart_type = st.selectbox("Chart", ["bar", "line"])

        if x_pick and y_pick and x_pick in df_raw.columns and y_pick in df_raw.columns:
            plot_df = df_raw[[x_pick, y_pick]].copy()
            # coerce numerics safely
            plot_df[y_pick] = pd.to_numeric(plot_df[y_pick], errors="coerce").fillna(plot_df[y_pick].map(numify))
            plot_df = plot_df.dropna(subset=[y_pick])
            if not plot_df.empty:
                if chart_type == "bar":
                    fig = px.bar(plot_df, x=x_pick, y=y_pick, text_auto=".2s")
                else:
                    fig = px.line(plot_df, x=x_pick, y=y_pick, markers=True)
                fig.update_layout(xaxis_title="", yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No numeric data to plot after cleaning.")
        else:
            st.caption("Pick one categorical column for X and one numeric column for Y to draw a chart.")


