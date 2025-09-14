# app.py
# TFC Most-Important-KPIs Dashboard (Rounds 1 & 2)

import os
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="TFC – Core KPIs (R1 & R2)", layout="wide", initial_sidebar_state="expanded")
st.title("The Fresh Connection – Core KPIs (Round 1 & 2)")

DATA_FILE = "Dashboard_Metrics_R1_R2_only.xlsx"

# ---------- helpers ----------
PERCENT = re.compile(r"%")
CURRENCY = re.compile(r"[€$,]")
PARENS = re.compile(r"^\((.*)\)$")

def numify(x) -> Optional[float]:
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    s = str(x).strip()
    if s == "": return None
    m = PARENS.match(s)
    if m: s = "-" + m.group(1)
    s = CURRENCY.sub("", s).replace(" ", "")
    if PERCENT.search(s):
        s = s.replace("%", "")
        try: return float(s)
        except: return None
    try: return float(s)
    except: return None

def norm_round(x) -> Optional[str]:
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    m = re.search(r"(\d+)", str(x))
    return f"R{int(m.group(1))}" if m else None

@st.cache_data
def load_wb(path: str) -> Dict[str, pd.DataFrame]:
    if not os.path.exists(path):
        st.error(f"Data file not found: `{path}`")
        return {}
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")

def find_sheet(sheets: Dict[str, pd.DataFrame], keywords: List[str]) -> Optional[str]:
    keys = [k.lower() for k in keywords]
    for name in sheets:
        low = name.lower()
        if all(k in low for k in keys):
            return name
    return None

@st.cache_data
def list_rounds(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    found = set()
    for n in sheets:
        r = norm_round(n)
        if r: found.add(r)
    for df in sheets.values():
        if isinstance(df, pd.DataFrame) and "Round" in df.columns:
            for v in df["Round"].unique():
                r = norm_round(v)
                if r: found.add(r)
    return sorted(found, key=lambda t: int(re.search(r"\d+", t).group(0))) or ["R1","R2"]

# ---------- KPI mining ----------
# We will look for metrics in three places:
# 1) KPI sheet(s) (single sheet with Round column OR per-round sheets)
# 2) Sales > customer report (Revenue, Distribution costs, OSA, shelf life)
# 3) Supply Chain > product (MAPE, plan adherence, stock value, economic inv weeks, obsoletes value)
# 4) Purchasing / Component tables (delivery reliability, component availability, rejection)

KPI_ALIASES = {
    "ROI (%)": ["ROI", "ROI (%)", "Return on investment"],
    "Service level (order lines, %)": ["Service level outbound order lines", "Service level (order lines)"],
    "Obsolete products (%)": ["Obsolete products (%)", "Obsoletes (%)"],
    "Gross margin (customer) (€ / week)": ["Gross margin (customer)","Gross margin per week","Gross margin"],
}

SALES_COLUMNS = {  # columns we try to find on Sales/customer sheet
    "Revenue": ["Revenue"],
    "Distribution costs": ["Distribution costs"],
    "OSA (%)": ["OSA"],
    "Attained shelf life (%)": ["Attained shelf life (%)","Shelf life attained (%)","Shelf life (%)"],
    "Attained contract index": ["Attained contract index","Contract index"],
}

SCM_PRODUCT_COLUMNS = {
    "Forecast error (MAPE, %)": ["Forecast error (MAPE)","MAPE"],
    "Production plan adherence (%)": ["Production plan adherence (%)","Plan adherence (%)"],
    "Stock value (€)": ["Stock value"],
    "Economic inventory (weeks)": ["Economic inventory of products (weeks)","Economic inventory (weeks)"],
    "Obsoletes value (€ / week)": ["Obsoletes value per week"],
    "Rejects (€ / week)": ["Rejects (value)","Rejects value per week"],
    "Start-up productivity loss (€ / batch)": ["Start up productivity loss per batch (value)","Startup productivity loss"],
    "Demand per week": ["Demand per week (pieces)","Demand per week","Demand"],
}

PURCH_COLUMNS = {
    "Delivery reliability (%)": ["Delivery reliability (%)"],
    "Component availability (%)": ["Component availability (%)"],
    "Rejection rate (%)": ["Rejection (%)","Reject rate (%)"],
}

def get_col(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    for a in aliases:
        for c in df.columns:
            if a.lower() == str(c).lower().strip():
                return c
    return None

def sheet_table_with(df: pd.DataFrame, alias_dict: Dict[str, List[str]]) -> Dict[str, str]:
    """Return a dict of canonical_name -> actual_column_name for matches found in df."""
    found = {}
    for k, al in alias_dict.items():
        c = get_col(df, al)
        if c: found[k] = c
    return found

def clean_numeric_series(s: pd.Series) -> pd.Series:
    s2 = pd.to_numeric(s, errors="coerce")
    if s2.notna().any():
        return s2
    return s.map(numify)

def sum_metric(df: pd.DataFrame, col: str) -> float:
    return clean_numeric_series(df[col]).fillna(0).sum()

def avg_metric(df: pd.DataFrame, col: str) -> float:
    ser = clean_numeric_series(df[col])
    return float(ser.dropna().mean()) if ser.notna().any() else float("nan")

def wavg_metric(df: pd.DataFrame, val_col: str, w_col: str) -> float:
    v = clean_numeric_series(df[val_col])
    w = clean_numeric_series(df[w_col])
    mask = v.notna() & w.notna()
    if not mask.any(): return float("nan")
    return float((v[mask] * w[mask]).sum() / w[mask].sum()) if w[mask].sum() != 0 else float("nan")

# ---- Load workbook
sheets = load_wb(DATA_FILE)
rounds = list_rounds(sheets)

with st.sidebar:
    st.header("Controls")
    round_sel = st.radio("Select round", rounds, horizontal=True, index=0)
    st.caption(f"File: `{DATA_FILE}`")

# ---- Resolve sheet candidates
def pick_sheet(base_words: List[str], round_token: Optional[str] = None) -> Optional[str]:
    if round_token:
        n = find_sheet(sheets, base_words + [round_token])
        if n: return n
    return find_sheet(sheets, base_words)

# KPI sheet(s)
kpi_sheet = pick_sheet(["kpi"])  # single KPI sheet with Round col, or generic
# Sales – customer report
sales_sheet = pick_sheet(["sales","customer"])
# Supply Chain – product
scm_prod_sheet = pick_sheet(["supply","chain","product"])
# Purchasing – component
purch_comp_sheet = pick_sheet(["purchasing","component"])

# ---------- Extract per-round KPI dict ----------
def kpi_value_from_kpi_sheet(metric_label: str, r_tok: str) -> Optional[float]:
    """Try to read KPI values directly from KPI sheet(s)."""
    if not kpi_sheet: return None
    df = sheets[kpi_sheet].copy()
    # single KPI sheet with Round column?
    if "Round" in df.columns:
        df["__R__"] = df["Round"].map(norm_round)
        sub = df[df["__R__"] == r_tok]
        # try Metric/Value
        metric_col = get_col(sub, ["Metric","KPI","Name"])
        value_col = get_col(sub, ["Value","Amount","Score"])
        if metric_col and value_col:
            # locate row by aliases
            for alias in KPI_ALIASES.get(metric_label, [metric_label]):
                row = sub[sub[metric_col].astype(str).str.strip().str.lower() == alias.lower()]
                if not row.empty:
                    return numify(row.iloc[0][value_col])
        # or wide format (metrics as columns)
        for alias in KPI_ALIASES.get(metric_label, [metric_label]):
            if alias in sub.columns:
                return numify(sub.iloc[0][alias])
        return None
    # else: maybe we have a per-round KPI sheet (e.g., KPI_R1)
    round_sheet = pick_sheet(["kpi", r_tok])
    if round_sheet:
        df = sheets[round_sheet].copy()
        metric_col = get_col(df, ["Metric","KPI","Name"])
        value_col = get_col(df, ["Value","Amount","Score"])
        if metric_col and value_col:
            for alias in KPI_ALIASES.get(metric_label, [metric_label]):
                row = df[df[metric_col].astype(str).str.strip().str.lower() == alias.lower()]
                if not row.empty:
                    return numify(row.iloc[0][value_col])
        # wide fallback
        for alias in KPI_ALIASES.get(metric_label, [metric_label]):
            if alias in df.columns:
                return numify(df.iloc[0][alias])
    return None

def collect_kpis_for_round(r_tok: str) -> Dict[str, float]:
    out: Dict[str, float] = {}

    # 1) KPI sheet values
    for label in ["ROI (%)", "Service level (order lines, %)", "Obsolete products (%)", "Gross margin (customer) (€ / week)"]:
        v = kpi_value_from_kpi_sheet(label, r_tok)
        if v is not None:
            out[label] = float(v)

    # 2) Sales customer sheet aggregates
    if sales_sheet:
        df = sheets[sales_sheet].copy()
        if "Round" in df.columns:
            df = df[df["Round"].map(norm_round) == r_tok]
        cols = sheet_table_with(df, SALES_COLUMNS)
        if "Revenue" in cols:
            out["Revenue (€ / week)"] = sum_metric(df, cols["Revenue"])
        if "Distribution costs" in cols:
            out["Distribution costs (€ / week)"] = sum_metric(df, cols["Distribution costs"])
        if "OSA (%)" in cols:
            out["OSA (%)"] = avg_metric(df, cols["OSA (%)"])
        if "Attained shelf life (%)" in cols:
            out["Attained shelf life (%)"] = avg_metric(df, cols["Attained shelf life (%)"])
        if "Attained contract index" in cols:
            out["Attained contract index"] = avg_metric(df, cols["Attained contract index"])

    # 3) SCM product sheet aggregates
    if scm_prod_sheet:
        dfp = sheets[scm_prod_sheet].copy()
        if "Round" in dfp.columns:
            dfp = dfp[dfp["Round"].map(norm_round) == r_tok]
        cols = sheet_table_with(dfp, SCM_PRODUCT_COLUMNS)
        # demand-weighted MAPE if demand exists
        if "Forecast error (MAPE, %)" in cols:
            if "Demand per week" in cols:
                out["Forecast error (MAPE, %)"] = wavg_metric(dfp, cols["Forecast error (MAPE, %)"], cols["Demand per week"])
            else:
                out["Forecast error (MAPE, %)"] = avg_metric(dfp, cols["Forecast error (MAPE, %)"])
        if "Production plan adherence (%)" in cols:
            out["Production plan adherence (%)"] = avg_metric(dfp, cols["Production plan adherence (%)"])
        if "Stock value (€)" in cols:
            out["Stock value (€)"] = sum_metric(dfp, cols["Stock value (€)"])
        if "Economic inventory (weeks)" in cols:
            # stock-value weighted weeks
            if "Stock value (€)" in cols:
                out["Economic inventory (weeks)"] = wavg_metric(dfp, cols["Economic inventory (weeks)"], cols["Stock value (€)"])
            else:
                out["Economic inventory (weeks)"] = avg_metric(dfp, cols["Economic inventory (weeks)"])
        if "Obsoletes value (€ / week)" in cols:
            out["Obsoletes value (€ / week)"] = sum_metric(dfp, cols["Obsoletes value (€ / week)"])
        if "Rejects (€ / week)" in cols:
            out["Rejects (€ / week)"] = sum_metric(dfp, cols["Rejects (€ / week)"])
        if "Start-up productivity loss (€ / batch)" in cols:
            out["Start-up productivity loss (€ / batch)"] = avg_metric(dfp, cols["Start-up productivity loss (€ / batch)"])

    # 4) Purchasing / Components
    if purch_comp_sheet:
        dpc = sheets[purch_comp_sheet].copy()
        if "Round" in dpc.columns:
            dpc = dpc[dpc["Round"].map(norm_round) == r_tok]
        cols = sheet_table_with(dpc, PURCH_COLUMNS)
        if "Delivery reliability (%)" in cols:
            out["Delivery reliability (%)"] = avg_metric(dpc, cols["Delivery reliability (%)"])
        if "Component availability (%)" in cols:
            out["Component availability (%)"] = avg_metric(dpc, cols["Component availability (%)"])
        if "Rejection rate (%)" in cols:
            out["Rejection rate (%)"] = avg_metric(dpc, cols["Rejection rate (%)"])

    return out

# build R1/R2 dicts
data_by_round = {r: collect_kpis_for_round(r) for r in rounds}

# ------ UI: KPI cards + grouped charts ------
def kpi_row(title: str, keys: List[str]):
    st.markdown(f"### {title}")
    cols = st.columns(len(keys))
    for i, k in enumerate(keys):
        r1 = data_by_round.get("R1", {}).get(k, None)
        r2 = data_by_round.get("R2", {}).get(k, None)
        if r1 is None and r2 is None:
            with cols[i]:
                st.info(f"'{k}' not found")
            continue
        # formatting
        def fmt(v, key):
            if v is None: return "–"
            if "€" in key or "value" in key.lower() or "cost" in key.lower():
                return f"€ {v:,.0f}"
            return f"{v:.1f}%"
        delta = None
        if r1 is not None and r2 is not None:
            delta = (r2 - r1)
            if "€" in k or "value" in k.lower() or "cost" in k.lower():
                delta = f"€ {delta:,.0f}"
            else:
                delta = f"{delta:.1f} pp"
        with cols[i]:
            st.metric(k, fmt(r2 if round_sel == "R2" else r1, k), delta=delta)

    # Bar compare R1 vs R2 (only for those we actually have)
    present = [k for k in keys if (k in data_by_round.get("R1", {}) or k in data_by_round.get("R2", {}))]
    if present:
        rows = []
        for k in present:
            for r in ["R1","R2"]:
                v = data_by_round.get(r, {}).get(k, None)
                if v is not None:
                    rows.append({"KPI": k, "Round": r, "Value": v})
        if rows:
            dfp = pd.DataFrame(rows)
            fig = px.bar(dfp, x="KPI", y="Value", color="Round", barmode="group", text_auto=".2s")
            fig.update_layout(xaxis_title="", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

# Executive
kpi_row("Executive", [
    "ROI (%)",
    "Revenue (€ / week)",
    "Gross margin (customer) (€ / week)",
    "Distribution costs (€ / week)",
])

# Customer reliability
kpi_row("Customer reliability", [
    "Service level (order lines, %)",
    "OSA (%)",
    "Attained shelf life (%)",
    "Attained contract index",
])

# Inventory & waste
kpi_row("Inventory & waste", [
    "Obsolete products (%)",
    "Obsoletes value (€ / week)",
    "Stock value (€)",
    "Economic inventory (weeks)",
])

# Planning & operations
kpi_row("Planning & operations", [
    "Forecast error (MAPE, %)",
    "Production plan adherence (%)",
    "Rejects (€ / week)",
    "Start-up productivity loss (€ / batch)",
])

# Supply risk
kpi_row("Purchasing – supply risk", [
    "Delivery reliability (%)",
    "Component availability (%)",
    "Rejection rate (%)",
])

st.caption("Notes: Values are parsed from the workbook, cleaning € / % / comma formats automatically. "
           "Metrics that do not exist in your Excel are skipped; you can add them to the KPI sheet or the relevant area sheet and they’ll show up on the next deploy.")



