# app.py
# The Fresh Connection – Dashboard (R1 & R2)
# Works with "Dashboard_Metrics_R1_R2_only.xlsx"
# - Handles both: sheets split by round (e.g., KPI_R1/KPI_R2) OR
#   a single sheet with a Round column (values like 1, 1.0, R1, "Round 1", etc.)

import os
import re
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
import plotly.express as px

# ---------- Page setup ----------
st.set_page_config(
    page_title="The Fresh Connection – Dashboard (R1 & R2)",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("The Fresh Connection – Dashboard (R1 & R2)")

DATA_FILE = "Dashboard_Metrics_R1_R2_only.xlsx"

# ---------- Helpers: robust round token ----------
def _norm_round_token(x) -> Optional[str]:
    """
    Normalize any input (1, 1.0, 'R1', 'Round 1', etc.) to 'R#'.
    Returns None if no digits are found.
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip().upper()
    m = re.search(r"(\d+)", s)
    if not m:
        return None
    return f"R{int(m.group(1))}"

# ---------- Data loading ----------
@st.cache_data(show_spinner=True)
def load_sheets(path: str) -> Dict[str, pd.DataFrame]:
    if not os.path.exists(path):
        st.error(f"Data file not found: {path}")
        return {}
    # Keep original dtypes; we’ll normalize 'Round' with _norm_round_token
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")

@st.cache_data(show_spinner=False)
def list_rounds(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    """
    Collect available rounds from sheet names and/or a 'Round' column,
    and return a sorted list like ['R1', 'R2'].
    """
    rs = set()

    # 1) Try to read round info from sheet names
    for name in sheets:
        tok = _norm_round_token(name)
        if tok:
            rs.add(tok)

    # 2) Try to read round info from any 'Round' column
    for df in sheets.values():
        if isinstance(df, pd.DataFrame) and "Round" in df.columns:
            for v in df["Round"].unique():
                tok = _norm_round_token(v)
                if tok:
                    rs.add(tok)

    if not rs:
        # Safe fallback for the assignment context
        return ["R1", "R2"]

    # Sort numerically by the digits within 'R#'
    return sorted(rs, key=lambda t: int(re.search(r"\d+", t).group(0)))

def _find_first_sheet_with_keywords(sheets: Dict[str, pd.DataFrame], keywords: List[str]) -> Optional[str]:
    """
    Return the first sheet name that contains ALL keywords (case-insensitive).
    """
    kl = [k.lower() for k in keywords]
    for name in sheets:
        low = name.lower()
        if all(k in low for k in kl):
            return name
    return None

@st.cache_data(show_spinner=False)
def build_views(sheets: Dict[str, pd.DataFrame], round_sel: str) -> Dict[str, pd.DataFrame]:
    """
    Returns a dict of DataFrames filtered to the selected round for each area:
    KPI, Purchasing, Operations, Sales, Supply Chain, Finance.
    Handles:
      - Sheet per round (e.g., 'KPI_R1') OR
      - Single sheet with a 'Round' column
    """
    out: Dict[str, pd.DataFrame] = {}
    rtoken = _norm_round_token(round_sel)

    def get_df(area_name: str, area_keywords: List[str]) -> Optional[pd.DataFrame]:
        # Prefer a sheet that also explicitly contains the round token
        name_with_round = _find_first_sheet_with_keywords(sheets, area_keywords + [round_sel])
        if name_with_round:
            return sheets[name_with_round].copy()

        # Fallback to a generic area sheet
        generic_name = _find_first_sheet_with_keywords(sheets, area_keywords)
        if generic_name:
            df = sheets[generic_name].copy()
            # If it has a 'Round' column, filter it
            if "Round" in df.columns:
                norm_col = df["Round"].map(_norm_round_token)
                df = df[norm_col == rtoken]
            return df

        return None

    out["KPI"] = get_df("KPI", ["kpi"])
    out["Purchasing"] = get_df("Purchasing", ["purchasing"])
    out["Operations"] = get_df("Operations", ["operations"])
    out["Sales"] = get_df("Sales", ["sales"])
    out["Supply Chain"] = get_df("Supply Chain", ["supply", "chain"])
    out["Finance"] = get_df("Finance", ["finance"])

    # Strip unnamed columns that sometimes appear when exporting to Excel
    for k, df in list(out.items()):
        if df is None:
            continue
        drop_cols = [c for c in df.columns if str(c).lower().startswith("unnamed")]
        if drop_cols:
            df = df.drop(columns=drop_cols)
        out[k] = df

    return out

# ---------- Simple visualization helpers ----------
def _try_number(v):
    try:
        return float(v)
    except Exception:
        return None

def show_table_and_quick_charts(title: str, df: Optional[pd.DataFrame]):
    st.subheader(title)
    if df is None or df.empty:
        st.info("No data found for this section (for the selected round).")
        return

    st.dataframe(df, use_container_width=True)

    # Try very gentle visualizations if the shape looks reasonable
    with st.expander("Quick charts", expanded=False):
        # Case 1: Metric/Value wide table
        metric_cols = [c for c in df.columns if str(c).strip().lower() in ["metric", "kpi", "name"]]
        value_cols = [c for c in df.columns if str(c).strip().lower() in ["value", "amount", "score"]]
        if metric_cols and value_cols:
            mcol = metric_cols[0]
            vcol = value_cols[0]
            # Only keep numeric values
            tmp = df[[mcol, vcol]].copy()
            tmp[vcol] = tmp[vcol].map(_try_number)
            tmp = tmp.dropna(subset=[vcol])
            if not tmp.empty:
                fig = px.bar(tmp, x=mcol, y=vcol, title=f"{title} – {vcol}", text_auto=".2s")
                fig.update_layout(xaxis_title="", yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)
                return  # one good chart is enough

        # Case 2: If there is a 'Category'/'Subcategory' + numeric column
        cat_cols = [c for c in df.columns if df[c].dtype == "object"]
        num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        if cat_cols and num_cols:
            fig = px.bar(df, x=cat_cols[0], y=num_cols[0], title=f"{title} – {num_cols[0]}", text_auto=".2s")
            fig.update_layout(xaxis_title="", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

# ---------- Main UI ----------
with st.sidebar:
    st.header("Controls")
    sheets_dict = load_sheets(DATA_FILE)
    rounds = list_rounds(sheets_dict)
    round_sel = st.radio("Select Round", rounds, index=0, horizontal=True)
    st.caption(f"Data file: `{DATA_FILE}`")

views = build_views(sheets_dict, round_sel)

# Tabs for sections
tabs = st.tabs(["KPI", "Purchasing", "Operations", "Sales", "Supply Chain", "Finance"])
sections = ["KPI", "Purchasing", "Operations", "Sales", "Supply Chain", "Finance"]

for tab, name in zip(tabs, sections):
    with tab:
        show_table_and_quick_charts(name, views.get(name))

st.markdown("---")
st.caption("Tip: If your Excel layout changes, this app will still try to adapt. "
           "If a section shows 'No data', check the sheet names or ensure a 'Round' column exists.")


