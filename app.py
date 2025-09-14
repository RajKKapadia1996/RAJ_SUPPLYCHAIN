# app.py
import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="TFC Dashboard (Rounds 1 & 2)", layout="wide")

# ---------- Paths (Option B: file in repo root) ----------
DATA_PATH = Path(__file__).parent / "Dashboard_Metrics_R1_R2_only.xlsx"

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def load_workbook(path: Path) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(path, engine="openpyxl")
    sheets = {name: xl.parse(name) for name in xl.sheet_names}
    # normalize column names
    for k, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
    return sheets

def find_sheet(sheets: dict, candidates: list[str]) -> str | None:
    lc_map = {k.lower(): k for k in sheets}
    for want in candidates:
        for k_lc, k_real in lc_map.items():
            if all(token in k_lc for token in want.split("|")):  # simple fuzzy contains
                return k_real
    return None

def try_find_col(df: pd.DataFrame, *contains) -> str | None:
    # return the first column that includes all tokens (case-insensitive)
    cols = list(df.columns)
    lc = {c: c.lower() for c in cols}
    for c in cols:
        if all(tok in lc[c] for tok in contains):
            return c
    return None

def nice_pct(v):
    try:
        return f"{float(v):.1f}%"
    except Exception:
        return "—"

def kpi_block(df: pd.DataFrame, round_col: str, round_val):
    # guess common KPI columns
    roi_col = try_find_col(df, "roi")
    sl_ol_col = (try_find_col(df, "service", "order")
                 or try_find_col(df, "outbound", "order"))
    obs_col = try_find_col(df, "obsolete") or try_find_col(df, "obsolet")
    gm_col = (try_find_col(df, "gross", "margin")
              or try_find_col(df, "gm"))

    row = df[df[round_col] == round_val].tail(1)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if roi_col in df:
            st.metric("ROI", nice_pct(row[roi_col].iloc[0]))
        else:
            st.metric("ROI", "—")
    with c2:
        if sl_ol_col in df:
            st.metric("Service level (order lines)", nice_pct(row[sl_ol_col].iloc[0]))
        else:
            st.metric("Service level (order lines)", "—")
    with c3:
        if obs_col in df:
            st.metric("Obsolete products", nice_pct(row[obs_col].iloc[0]))
        else:
            st.metric("Obsolete products", "—")
    with c4:
        if gm_col in df:
            val = row[gm_col].iloc[0]
            st.metric("Gross margin (weekly/total)", f"€{val:,.0f}" if pd.notna(val) else "—")
        else:
            st.metric("Gross margin (weekly/total)", "—")

    # time series (if columns exist)
    times = []
    if roi_col in df: times.append(("ROI (%)", roi_col))
    if sl_ol_col in df: times.append(("Service level (order lines) %", sl_ol_col))
    if obs_col in df: times.append(("Obsolete products %", obs_col))
    if gm_col in df: times.append(("Gross margin", gm_col))

    if times:
        ts = df[[round_col] + [c for _, c in times]].copy()
        ts = ts.melt(id_vars=round_col, var_name="KPI", value_name="Value")
        # Relabel human-friendly
        rename_map = {c: label for label, c in times}
        ts["KPI"] = ts["KPI"].map(rename_map).fillna(ts["KPI"])
        fig = px.line(ts, x=round_col, y="Value", color="KPI",
                      markers=True, title="KPI trend by round")
        st.plotly_chart(fig, use_container_width=True)

# ---------- Load ----------
if not DATA_PATH.exists():
    st.error(f"Could not find Excel file at: {DATA_PATH.name}\n"
             "Make sure it is committed to the repo root.")
    st.stop()

sheets = load_workbook(DATA_PATH)

# ---------- Guess key sheets ----------
kpi_sheet = find_sheet(sheets, ["kpi", "overview", "dashboard"])
purch_sheet = find_sheet(sheets, ["purch", "buyer"])
ops_sheet = find_sheet(sheets, ["oper", "mixing|bottl"])
sales_sheet = find_sheet(sheets, ["sales", "customer"])
scm_sheet = find_sheet(sheets, ["supply", "chain", "scm"])
fin_sheet = find_sheet(sheets, ["finan"])

# pick a dataframe to get the round values
round_source = None
for s in [kpi_sheet, scm_sheet, sales_sheet, purch_sheet, ops_sheet, fin_sheet]:
    if s:
        round_source = sheets[s]
        break

round_col = try_find_col(round_source, "round") if round_source is not None else None
if round_col is None and round_source is not None:
    # create a synthetic round if needed (1..n)
    round_source = round_source.copy()
    round_source["Round"] = range(1, len(round_source) + 1)
    round_col = "Round"

round_values = sorted(round_source[round_col].unique().tolist()) if round_source is not None else [1, 2]
default_round = 2 if 2 in round_values else round_values[-1]

# ---------- Sidebar ----------
st.sidebar.header("Filters")
sel_round = st.sidebar.selectbox("Round", round_values, index=round_values.index(default_round))
st.sidebar.write("Excel file:", f"`{DATA_PATH.name}`")

# ---------- Main ----------
st.title("The Fresh Connection — Rounds 1 & 2 Dashboard")

tab_labels = ["Overview KPIs", "Purchasing", "Operations", "Sales", "Supply Chain", "Finance", "All Sheets"]
tabs = st.tabs(tab_labels)

# ---- Overview KPIs ----
with tabs[0]:
    if kpi_sheet:
        df = sheets[kpi_sheet].copy()
        if round_col not in df.columns:
            # If KPI sheet doesn't have round, try to inject from global round list (assumes 2 rows)
            df.insert(0, round_col, round_values[:len(df)])
        st.subheader(f"Overview KPIs · Sheet: `{kpi_sheet}`")
        kpi_block(df, round_col, sel_round)
        with st.expander("Show KPI data"):
            st.dataframe(df, use_container_width=True)
    else:
        st.info("No KPI/Overview sheet detected. Showing first sheet instead.")
        first_name = list(sheets.keys())[0]
        st.dataframe(sheets[first_name], use_container_width=True)

# ---- Purchasing ----
with tabs[1]:
    if purch_sheet:
        df = sheets[purch_sheet].copy()
        st.subheader(f"Purchasing · Sheet: `{purch_sheet}`")

        comp_col = try_find_col(df, "component") or try_find_col(df, "material") or try_find_col(df, "sku")
        dr_col = try_find_col(df, "delivery", "reliab")
        ca_col = try_find_col(df, "component", "avail")
        price_col = try_find_col(df, "purchase", "price")

        if comp_col and round_col in df:
            sub = df[df[round_col] == sel_round]
            c1, c2 = st.columns([2, 1])
            with c1:
                if dr_col:
                    fig = px.bar(sub, x=comp_col, y=dr_col, title="Delivery reliability by component")
                    st.plotly_chart(fig, use_container_width=True)
            with c2:
                if price_col:
                    st.metric("Avg purchase price", f"€{sub[price_col].mean():,.4f}")
            st.dataframe(sub, use_container_width=True)
        else:
            st.dataframe(df, use_container_width=True)
    else:
        st.info("No Purchasing sheet detected.")

# ---- Operations ----
with tabs[2]:
    if ops_sheet:
        df = sheets[ops_sheet].copy()
        st.subheader(f"Operations · Sheet: `{ops_sheet}`")
        name_col = try_find_col(df, "line") or try_find_col(df, "mixer") or try_find_col(df, "work")
        rt_col = try_find_col(df, "run", "time")
        ot_col = try_find_col(df, "overtime")
        ad_col = try_find_col(df, "adherence") or try_find_col(df, "plan", "adher")
        if round_col in df.columns:
            sub = df[df[round_col] == sel_round]
        else:
            sub = df
        c1, c2, c3 = st.columns(3)
        with c1:
            if rt_col:
                st.metric("Run time (h, avg)", f"{sub[rt_col].mean():.1f} h")
        with c2:
            if ot_col:
                st.metric("Overtime (h, avg)", f"{sub[ot_col].mean():.1f} h")
        with c3:
            if ad_col:
                st.metric("Plan adherence (avg)", nice_pct(sub[ad_col].mean()))
        if name_col and rt_col:
            fig = px.bar(sub, x=name_col, y=rt_col, title="Run time by line")
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(sub, use_container_width=True)
    else:
        st.info("No Operations sheet detected.")

# ---- Sales ----
with tabs[3]:
    if sales_sheet:
        df = sheets[sales_sheet].copy()
        st.subheader(f"Sales · Sheet: `{sales_sheet}`")
        prod_col = try_find_col(df, "product")
        cust_col = try_find_col(df, "customer")
        sl_pieces = try_find_col(df, "service", "pieces") or try_find_col(df, "service level (pieces)")
        gm_piece = try_find_col(df, "gross", "piece")
        if round_col in df.columns:
            sub = df[df[round_col] == sel_round]
        else:
            sub = df
        c1, c2 = st.columns(2)
        with c1:
            if prod_col and sl_pieces:
                fig = px.bar(sub, x=prod_col, y=sl_pieces, color=cust_col if cust_col in sub else None,
                             title="Service level (pieces) by product")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            if gm_piece and prod_col:
                fig = px.bar(sub, x=prod_col, y=gm_piece, color=cust_col if cust_col in sub else None,
                             title="Gross margin per piece by product")
                st.plotly_chart(fig, use_container_width=True)
        st.dataframe(sub, use_container_width=True)
    else:
        st.info("No Sales sheet detected.")

# ---- Supply Chain ----
with tabs[4]:
    if scm_sheet:
        df = sheets[scm_sheet].copy()
        st.subheader(f"Supply Chain · Sheet: `{scm_sheet}`")
        prod_col = try_find_col(df, "product")
        sl_ol_col = (try_find_col(df, "service", "order") or try_find_col(df, "order", "lines"))
        obs_col = try_find_col(df, "obsolete")
        mape_col = try_find_col(df, "mape") or try_find_col(df, "forecast", "error")
        if round_col in df.columns:
            sub = df[df[round_col] == sel_round]
        else:
            sub = df
        c1, c2, c3 = st.columns(3)
        with c1:
            if sl_ol_col:
                st.metric("Service level (order lines)", nice_pct(sub[sl_ol_col].mean()))
        with c2:
            if obs_col:
                st.metric("Obsoletes (%)", nice_pct(sub[obs_col].mean()))
        with c3:
            if mape_col:
                st.metric("Forecast error (MAPE)", nice_pct(sub[mape_col].mean()))
        if prod_col and obs_col:
            fig = px.bar(sub, x=prod_col, y=obs_col, title="Obsoletes by product")
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(sub, use_container_width=True)
    else:
        st.info("No Supply Chain sheet detected.")

# ---- Finance ----
with tabs[5]:
    if fin_sheet:
        df = sheets[fin_sheet].copy()
        st.subheader(f"Finance · Sheet: `{fin_sheet}`")
        if round_col in df.columns:
            sub = df[df[round_col] == sel_round]
        else:
            sub = df
        st.dataframe(sub, use_container_width=True)
        # Try a simple bar chart for a few money columns
        money_cols = [c for c in sub.columns if any(k in c.lower() for k in ["€", "revenue", "margin", "cost", "cash"])]
        if money_cols:
            long = sub[[round_col] + money_cols].melt(id_vars=round_col, var_name="Metric", value_name="Value")
            fig = px.bar(long, x="Metric", y="Value", barmode="group", title="Finance metrics")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No Finance sheet detected.")

# ---- All sheets browser ----
with tabs[6]:
    st.subheader("Browse all sheets")
    name = st.selectbox("Sheet", list(sheets.keys()))
    st.dataframe(sheets[name], use_container_width=True)
