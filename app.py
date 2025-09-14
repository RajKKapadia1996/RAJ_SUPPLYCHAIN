import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Fresh Connection Dashboard", layout="wide")

# ----------------------------
# HELPERS WITH CACHING
# ----------------------------
@st.cache_data
def load_finance():
    return pd.read_excel("FinanceReport.xlsx", sheet_name="Output")

@st.cache_data
def load_sheet(file, sheet):
    return pd.read_excel(file, sheet_name=sheet)

def extract_rounds(df):
    """Extract only Round 1 and 2"""
    out = df[[df.columns[0], 1, 2]].copy()
    out.columns = ["Metric", "Round 1", "Round 2"]
    return out

def plot_bar(df, x, y_cols, title, ylabel):
    """Generic bar chart with Plotly"""
    df_long = df.melt(id_vars=[x], value_vars=y_cols, var_name="Round", value_name=ylabel)
    fig = px.bar(df_long, x=x, y=ylabel, color="Round", barmode="group", title=title)
    st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# LOAD DATA
# ----------------------------
finance = load_finance()
finance_r = extract_rounds(finance)

# ----------------------------
# NAVIGATION
# ----------------------------
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "Go to:",
    ["Q1: Strategy", "Q2: Functional Decisions", "Q3: KPI Outcomes", "Q4: Next Round Plan"]
)

# ----------------------------
# Q1: STRATEGY
# ----------------------------
if page == "Q1: Strategy":
    st.header("Q1. Supply Chain Strategy – Cost Leadership with Service Floor")

    roi = finance_r[finance_r["Metric"]=="ROI"]
    penalties_total = finance_r[finance_r["Metric"].str.contains("Bonus or penalties$")]
    realized_rev = finance_r[finance_r["Metric"]=="Realized revenue"]

    st.subheader("ROI")
    st.dataframe(roi)
    plot_bar(roi, "Metric", ["Round 1","Round 2"], "ROI (R1 vs R2)", "ROI")

    st.subheader("Total Penalties")
    st.dataframe(penalties_total)
    plot_bar(penalties_total, "Metric", ["Round 1","Round 2"], "Total Penalties", "€")

    st.subheader("Net Realized Revenue")
    st.dataframe(realized_rev)
    plot_bar(realized_rev, "Metric", ["Round 1","Round 2"], "Realized Revenue", "€")

# ----------------------------
# Q2: FUNCTIONAL DECISIONS
# ----------------------------
elif page == "Q2: Functional Decisions":
    st.header("Q2. Functional Alignment – Sales, SCM, Operations, Purchasing")

    # SALES: Penalties by Customer
    pen_cust = finance_r[finance_r["Metric"].str.contains("Bonus or penalties - Contracted sales revenue -")]
    if not pen_cust.empty:
        pen_cust = pen_cust.copy()
        pen_cust["Customer"] = pen_cust["Metric"].str.split("-").str[-1].str.strip()
        st.subheader("Sales – Penalties by Customer")
        st.dataframe(pen_cust[["Customer","Round 1","Round 2"]])
        plot_bar(pen_cust, "Customer", ["Round 1","Round 2"], "Penalties by Customer", "€")

    # SALES: Revenue by Customer
    rev_cust = finance_r[finance_r["Metric"].str.contains("Contracted sales revenue -")]
    if not rev_cust.empty:
        rev_cust = rev_cust.copy()
        rev_cust["Customer"] = rev_cust["Metric"].str.split("-").str[-1].str.strip()
        st.subheader("Sales – Revenue by Customer")
        st.dataframe(rev_cust[["Customer","Round 1","Round 2"]])
        plot_bar(rev_cust, "Customer", ["Round 1","Round 2"], "Revenue by Customer", "€")

    # PURCHASING
    supc = load_sheet("Purchase.xlsx", "Supplier - Component")
    if "Delivery reliability (%)" in supc.columns:
        rel_summary = supc.groupby("Round")["Delivery reliability (%)"].mean().reset_index()
        st.subheader("Purchasing – Average Supplier Reliability")
        st.dataframe(rel_summary)
        fig = px.bar(rel_summary, x="Round", y="Delivery reliability (%)", title="Supplier Reliability (%)")
        st.plotly_chart(fig, use_container_width=True)

    if "Order size" in supc.columns:
        order_summary = supc.groupby("Round")["Order size"].mean().reset_index()
        st.subheader("Purchasing – Average Order Size")
        st.dataframe(order_summary)
        fig = px.bar(order_summary, x="Round", y="Order size", title="Order Size")
        st.plotly_chart(fig, use_container_width=True)

    # OPERATIONS
    bott = load_sheet("Operations.xlsx", "Bottling line")
    bott_summary = bott.groupby("Round")[["Production plan adherence (%)","Overtime per week (hours)"]].mean().reset_index()
    st.subheader("Operations – Bottling KPIs")
    st.dataframe(bott_summary)
    fig = px.bar(bott_summary.melt(id_vars="Round"), x="Round", y="value", color="variable", barmode="group",
                 title="Bottling Line KPIs")
    st.plotly_chart(fig, use_container_width=True)

    # SUPPLY CHAIN
    comp = load_sheet("Supply chain.xlsx", "Component")
    if "Obsoletes (%)" in comp.columns:
        obs_summary = comp.groupby("Round")["Obsoletes (%)"].mean().reset_index()
        st.subheader("Supply Chain – Obsolete Products (%)")
        st.dataframe(obs_summary)
        fig = px.bar(obs_summary, x="Round", y="Obsoletes (%)", title="Obsolete Products (%)")
        st.plotly_chart(fig, use_container_width=True)

    if "Stock value" in comp.columns:
        stock_summary = comp.groupby("Round")["Stock value"].sum().reset_index()
        st.subheader("Supply Chain – Stock Value (€)")
        st.dataframe(stock_summary)
        fig = px.bar(stock_summary, x="Round", y="Stock value", title="Stock Value (€)")
        st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# Q3: KPI OUTCOMES
# ----------------------------
elif page == "Q3: KPI Outcomes":
    st.header("Q3. KPI Outcomes – Achieved vs Not Achieved")

    st.subheader("Finance KPIs")
    st.dataframe(finance_r.head(15))
    plot_bar(finance_r.head(15), "Metric", ["Round 1","Round 2"], "Finance KPIs (R1 vs R2)", "Value")

    # OPERATIONS
    bott = load_sheet("Operations.xlsx", "Bottling line")
    ops_out = bott.groupby("Round")[["Run time (%)","Overtime (%)","Production plan adherence (%)"]].mean().reset_index()
    st.subheader("Operations – Outcomes")
    st.dataframe(ops_out)
    fig = px.bar(ops_out.melt(id_vars="Round"), x="Round", y="value", color="variable", barmode="group",
                 title="Operations Outcomes")
    st.plotly_chart(fig, use_container_width=True)

    # SUPPLY CHAIN
    comp = load_sheet("Supply chain.xlsx", "Component")
    if "Delivery reliability (%)" in comp.columns and "Component availability (%)" in comp.columns:
        scm_out = comp.groupby("Round")[["Delivery reliability (%)","Component availability (%)"]].mean().reset_index()
        st.subheader("Supply Chain – Reliability & Availability")
        st.dataframe(scm_out)
        fig = px.bar(scm_out.melt(id_vars="Round"), x="Round", y="value", color="variable", barmode="group",
                     title="SCM Reliability & Availability")
        st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# Q4: NEXT ROUND PLAN
# ----------------------------
elif page == "Q4: Next Round Plan":
    st.header("Q4. Strategy for Next Round")
    st.markdown("""
    **Key Next Steps:**
    - Keep 17:00 cut-off (trial 14:00 for LAND/Dominick’s).
    - Trim PET FG SS to 1.8–2.0 weeks if service ≥95% and scrap low.
    - Reduce 1-L FG SS from 3.0 → 2.5 weeks if service stable.
    - Maintain 2 bottling shifts; if outbound flex >200h, add 1 FTE.
    - Increase Vit-C SS to 3.0 weeks if availability <97%.
    - Use promotions selectively only if capacity utilization <70%.
    """)


