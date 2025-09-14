import time
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

# ---------- Page setup ----------
st.set_page_config(page_title="TFC Dashboard", layout="wide")

DATA_PATH = Path(__file__).parent / "Dashboard_Metrics_R1_R2_only.xlsx"

# ---------- Utilities ----------
def small_timer(msg: str):
    """Simple context timer to spot slow spots in logs."""
    class _T:
        def __enter__(self): self.t0 = time.perf_counter()
        def __exit__(self, *exc):
            st.session_state.setdefault("_timings", []).append(
                f"{msg}: {(time.perf_counter() - self.t0)*1000:.1f} ms"
            )
    return _T()

def optimize_df(df: pd.DataFrame) -> pd.DataFrame:
    # Downcast numerics and use categorical for low-cardinality text
    for c in df.select_dtypes(include="number").columns:
        df[c] = pd.to_numeric(df[c], downcast="float")
    for c in df.select_dtypes(include="object").columns:
        if df[c].nunique() <= max(8, len(df)//20):
            df[c] = df[c].astype("category")
    return df

# ---------- Cached IO & transforms ----------
@st.cache_data(show_spinner=False, ttl=3600)
def load_all_sheets(xlsx_path: Path) -> dict[str, pd.DataFrame]:
    with small_timer("read_excel(all)"):
        # Load every sheet once; keep as dict
        sheets = pd.read_excel(xlsx_path, sheet_name=None, engine="openpyxl")
    # Light optimization
    return {k: optimize_df(v) for k, v in sheets.items()}

@st.cache_data(show_spinner=False)
def list_rounds(sheets: dict[str, pd.DataFrame]) -> list[str]:
    # Infer rounds from sheet names like "..._R1", "..._R2" or a "Round" column
    # Works with either layout. Falls back to ["R1", "R2"].
    rs = set()
    for name, df in sheets.items():
        for token in ("R1", "R2"):
            if token in name:
                rs.add(token)
        if "Round" in df.columns:
            rs.update(df["Round"].astype(str).unique())
    if rs:
        return sorted(rs, key=lambda r: int(r.strip("R")))
    return ["R1", "R2"]

@st.cache_data(show_spinner=False)
def build_views(sheets: dict[str, pd.DataFrame], round_sel: str) -> dict[str, pd.DataFrame]:
    """
    Return small, ready-to-plot tables for the chosen round only.
    This function isolates the minimal data you need per section.
    """
    out: dict[str, pd.DataFrame] = {}

    # Try to be resilient to sheet names – lookups by substring:
    def find_sheet(keywords: list[str]):
        for name in sheets:
            name_low = name.lower()
            if all(k in name_low for k in keywords):
                return sheets[name]
        return None

    # KPIs
    kpi_df = find_sheet(["kpi", round_sel.lower()]) or find_sheet(["kpi"])
    if kpi_df is not None:
        df = kpi_df.copy()
        if "Round" in df.columns:
            df = df[df["Round"].astype(str).eq(round_sel)]
        out["KPI"] = df

    # Purchasing
    pur_df = find_sheet(["purchasing", round_sel.lower()]) or find_sheet(["purchasing"])
    if pur_df is not None:
        df = pur_df.copy()
        if "Round" in df.columns:
            df = df[df["Round"].astype(str).eq(round_sel)]
        out["Purchasing"] = df

    # Operations
    ops_df = find_sheet(["operations", round_sel.lower()]) or find_sheet(["operations"])
    if ops_df is not None:
        df = ops_df.copy()
        if "Round" in df.columns:
            df = df[df["Round"].astype(str).eq(round_sel)]
        out["Operations"] = df

    # Sales
    sales_df = find_sheet(["sales", round_sel.lower()]) or find_sheet(["sales"])
    if sales_df is not None:
        df = sales_df.copy()
        if "Round" in df.columns:
            df = df[df["Round"].astype(str).eq(round_sel)]
        out["Sales"] = df

    # Supply Chain
    sc_df = find_sheet(["supply", "chain", round_sel.lower()]) or find_sheet(["supply", "chain"])
    if sc_df is not None:
        df = sc_df.copy()
        if "Round" in df.columns:
            df = df[df["Round"].astype(str).eq(round_sel)]
        out["Supply Chain"] = df

    # Finance
    fin_df = find_sheet(["finance", round_sel.lower()]) or find_sheet(["finance"])
    if fin_df is not None:
        df = fin_df.copy()
        if "Round" in df.columns:
            df = df[df["Round"].astype(str).eq(round_sel)]
        out["Finance"] = df

    return out

# ---------- UI ----------
st.title("The Fresh Connection – Dashboard (R1 & R2)")

# File check with friendly error
if not DATA_PATH.exists():
    st.error(f"Excel file not found: `{DATA_PATH.name}`. "
             "Confirm the filename in the repo matches exactly.")
    st.stop()

with st.spinner("Loading metrics... (cached)"):
    sheets_dict = load_all_sheets(DATA_PATH)
    rounds = list_rounds(sheets_dict)

colL, colR = st.columns([1, 3], gap="large")
with colL:
    round_sel = st.radio("Round", options=rounds, index=len(rounds)-1, horizontal=True)
    section = st.selectbox("Section", ["KPI", "Purchasing", "Operations", "Sales", "Supply Chain", "Finance"])

    # Debug timings (optional)
    if "_timings" in st.session_state and st.toggle("Show load timings", value=False):
        st.code("\n".join(st.session_state["_timings"]))

with colR:
    views = build_views(sheets_dict, round_sel)
    df = views.get(section)
    if df is None or df.empty:
        st.info(f"No data found for **{section}** in **{round_sel}**.")
    else:
        st.subheader(f"{section} — {round_sel}")

        # Try to pick smart defaults for quick visual
        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        text_cols = df.select_dtypes(exclude="number").columns.tolist()

        # 1) KPI cards if we see common KPI fields
        kpi_like = [c for c in numeric_cols if any(k in c.lower() for k in ["roi", "margin", "service", "obsolete", "penalty"])]
        if kpi_like:
            kpi_like = kpi_like[:3]
            cks = st.columns(len(kpi_like))
            for i, c in enumerate(kpi_like):
                with c:
                    st.metric(label=c, value=f"{df[c].iloc[0]:,.2f}" if len(df) else "—")

        # 2) Quick chart: pick first text as x, first numeric as y
        if text_cols and numeric_cols:
            x, y = text_cols[0], numeric_cols[0]
            fig = px.bar(df, x=x, y=y, title=f"{y} by {x}", height=380)
            fig.update_layout(margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig, use_container_width=True)

        # 3) Table
        st.dataframe(df, use_container_width=True, height=420)

