import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="The Fresh Connection – R1 vs R2", page_icon="📈", layout="wide")

DATA_PATH = Path("data/Dashboard_Metrics_R1_R2_only.xlsx")

SHEETS = {
    "dim_round": "dim_round",
    "kpi_core": "kpi_core_rounds_1_2",        # compact KPI table
    "roi_series": "kpi_roi_by_round",         # ROI by round
    "finance_r1": "finance_round1_tidy",
    "finance_r2": "finance_round2_tidy",
    "ops_bottling": "ops_bottling_r2",
    "ops_warehouse": "ops_warehousing_r2",
    "sc_components": "sc_components_r2",
    "sc_fg": "sc_fg_r2",
    "sales_customer": "sales_customer_r2",
    "sales_product": "sales_product_r2",
    "sales_decisions": "sales_decisions_r2",
    "purchasing": "purchasing_r2",
    "sc_decisions": "sc_decisions_r2",
}

@st.cache_data(show_spinner=False)
def load_data(path: Path) -> dict:
    xls = pd.ExcelFile(path)
    data = {}
    for key, sheet in SHEETS.items():
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            # Normalize column whitespace
            df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
            data[key] = df
    return data

if not DATA_PATH.exists():
    st.error("Data file not found. Add `data/Dashboard_Metrics_R1_R2_only.xlsx` to the repo.")
    st.stop()

D = load_data(DATA_PATH)
st.title("📈 The Fresh Connection — Dashboard (Round 1 vs Round 2)")

# ------------------ Filters ------------------
round_choices = [1, 2]
col1, col2, col3 = st.columns(3)
with col1:
    selected_rounds = st.multiselect("Round", round_choices, default=round_choices)
with col2:
    customers = sorted(D["sales_customer"]["Customer"].unique()) if "sales_customer" in D else []
    selected_customers = st.multiselect("Customer", customers, default=customers)
with col3:
    prods_col = "SKU" if "SKU" in (D.get("sales_product", pd.DataFrame()).columns) else "Product"
    products = sorted(D["sales_product"][prods_col].unique()) if "sales_product" in D else []
    selected_products = st.multiselect("Product", products, default=products)

def fr(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["Round"].isin(selected_rounds)] if "Round" in df.columns else df

# ------------------ KPI cards ------------------
c1, c2, c3, c4 = st.columns(4)

# ROI (from roi_series)
roi_df = D.get("roi_series", pd.DataFrame())
if not roi_df.empty:
    roi_f = fr(roi_df)
    latest = roi_f.sort_values("Round").tail(1)
    if not latest.empty:
        c1.metric("ROI (%)", f"{float(latest['Value'].iloc[0])*100:.2f}")
else:
    c1.metric("ROI (%)", "—")

# Gross margin (from kpi_core if present; else try finance)
kpi_core = D.get("kpi_core", pd.DataFrame())
gm_val = None
if not kpi_core.empty and "Gross margin (customer)" in kpi_core.columns:
    gm_val = kpi_core[kpi_core["Round"].isin(selected_rounds)].sort_values("Round").tail(1)["Gross margin (customer)"].squeeze()
else:
    # fallback: try finance_r2_tidy
    fin_all = pd.concat([D.get("finance_r1", pd.DataFrame()), D.get("finance_r2", pd.DataFrame())], ignore_index=True)
    if not fin_all.empty:
        gm = fin_all[(fin_all["Metric"] == "Gross margin") & (fin_all["Round"].isin(selected_rounds))].sort_values("Round").tail(1)
        if not gm.empty:
            gm_val = gm["Value"].squeeze()
c2.metric("Gross margin (€)", f"{gm_val:,.0f}" if gm_val is not None else "—")

# Outbound service (if provided in kpi_core — optional)
if "Service level outbound order lines (%)" in kpi_core.columns:
    sl = kpi_core[kpi_core["Round"].isin(selected_rounds)].sort_values("Round").tail(1)["Service level outbound order lines (%)"].squeeze()
    c3.metric("Outbound service (order lines) %", f"{sl:.1f}")
else:
    c3.metric("Outbound service (order lines) %", "—")

# Obsolescence % (if provided in kpi_core — optional)
if "Obsolete products (%)" in kpi_core.columns:
    ob = kpi_core[kpi_core["Round"].isin(selected_rounds)].sort_values("Round").tail(1)["Obsolete products (%)"].squeeze()
    c4.metric("Obsolete products (%)", f"{ob:.1f}")
else:
    c4.metric("Obsolete products (%)", "—")

st.markdown("---")

# ------------------ Charts row 1: ROI + Service ------------------
colA, colB = st.columns(2)
if not roi_df.empty:
    fig = px.line(fr(roi_df), x="Round", y="Value", markers=True, title="ROI by Round")
    fig.update_yaxes(ticksuffix="")  # already %
    colA.plotly_chart(fig, use_container_width=True)

scm_fg = D.get("sc_fg", pd.DataFrame())
if not scm_fg.empty:
    sp = scm_fg.copy()
    if selected_products:
        sp = sp[sp["SKU"].isin(selected_products)]
    if {"SKU", "Service level (order lines)(%)"}.issubset(sp.columns):
        fig2 = px.bar(sp, x="SKU", y="Service level (order lines)(%)", title="Service Level (Order Lines) by SKU")
        colB.plotly_chart(fig2, use_container_width=True)

# ------------------ Charts row 2: Sales by product ------------------
sales_prod = D.get("sales_product", pd.DataFrame())
if not sales_prod.empty:
    sp = sales_prod.copy()
    name_col = "SKU" if "SKU" in sp.columns else prods_col
    if selected_products:
        sp = sp[sp[name_col].isin(selected_products)]
    cC, cD = st.columns(2)
    if {name_col, "Demand value/week"}.issubset(sp.columns):
        fig3 = px.bar(sp, x=name_col, y="Demand value/week", color="Round", barmode="group",
                      title="Demand Value per Week by Product")
        cC.plotly_chart(fig3, use_container_width=True)
    if {name_col, "Obsoletes(%)"}.issubset(sp.columns):
        fig4 = px.bar(sp, x=name_col, y="Obsoletes(%)", color="Round", barmode="group",
                      title="Obsolescence (%) by Product")
        cD.plotly_chart(fig4, use_container_width=True)

# ------------------ Charts row 3: Component reliability & econ inv ------------------
sc_comp = D.get("sc_components", pd.DataFrame())
if not sc_comp.empty:
    cE, cF = st.columns(2)
    if {"Component", "Delivery reliability(%)"}.issubset(sc_comp.columns):
        fig5 = px.bar(sc_comp, x="Component", y="Delivery reliability(%)",
                      title="Supplier Delivery Reliability (R2 Components)")
        cE.plotly_chart(fig5, use_container_width=True)
    if {"Component", "Economic inventory (weeks)"}.issubset(sc_comp.columns):
        fig6 = px.bar(sc_comp, x="Component", y="Economic inventory (weeks)",
                      title="Economic Inventory (Weeks) by Component (R2)")
        cF.plotly_chart(fig6, use_container_width=True)

# ------------------ Charts row 4: Customer view ------------------
sales_cust = D.get("sales_customer", pd.DataFrame())
if not sales_cust.empty:
    sc = sales_cust.copy()
    if selected_customers:
        sc = sc[sc["Customer"].isin(selected_customers)]
    cG, cH = st.columns(2)
    if {"Customer", "Gross margin/week"}.issubset(sc.columns):
        fig7 = px.bar(sc, x="Customer", y="Gross margin/week", title="Gross Margin per Week by Customer (R2)")
        cG.plotly_chart(fig7, use_container_width=True)
    if {"Customer", "Service level (order lines) %"}.issubset(sc.columns):
        fig8 = px.bar(sc, x="Customer", y="Service level (order lines) %", title="Service Level (Order Lines) by Customer (R2)")
        cH.plotly_chart(fig8, use_container_width=True)

# ------------------ Charts row 5: Ops snapshots ------------------
ops_w = D.get("ops_warehouse", pd.DataFrame())
if not ops_w.empty:
    cI, cJ = st.columns(2)
    if {"Location", "Cube utilization %"}.issubset(ops_w.columns):
        fig9 = px.bar(ops_w, x="Location", y="Cube utilization %", title="Warehouse Cube Utilization (R2)")
        cI.plotly_chart(fig9, use_container_width=True)
    if {"Location", "Overflow %"}.issubset(ops_w.columns):
        fig10 = px.bar(ops_w, x="Location", y="Overflow %", title="Warehouse Overflow (R2)")
        cJ.plotly_chart(fig10, use_container_width=True)

ops_line = D.get("ops_bottling", pd.DataFrame())
if not ops_line.empty and {"Run time (h)", "Changeover time (h)", "Breakdown time (h)", "Overtime (h)"}.issubset(ops_line.columns):
    st.subheader("Mixing & Bottling — Time Breakdown (R2)")
    melted = ops_line.melt(value_vars=["Run time (h)", "Changeover time (h)", "Breakdown time (h)", "Overtime (h)"],
                           var_name="Category", value_name="Hours")
    fig11 = px.bar(melted, x="Category", y="Hours", title="Line Time Allocation (Hours)")
    st.plotly_chart(fig11, use_container_width=True)

st.caption("Data source: Dashboard_Metrics_R1_R2_only.xlsx")
