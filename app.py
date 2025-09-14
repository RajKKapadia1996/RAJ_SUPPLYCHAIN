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
# LOAD DATA (CACHED)
# ----------------------------
finance = load_finance()
finance_r = extract_rounds(finance)

# Other files will be loaded when needed
# ----------------------------
# MAIN DASHBOARD
# ----------------------------
st.title("📊 Fresh Connection Supply Chain Dashboard")
st.markdown("Compare **Round 1 vs Round 2** KPIs across all functions.")

tabs = st.tabs(["Finance", "Sales", "Supply Chain", "Operations", "Purchasing"])

# ----------------------------
# FINANCE TAB
# ----------------------------
with tabs[0]:
    st.header("💰 Finance KPIs")

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
# SALES TAB
# ----------------------------
with tabs[1]:
    st.header("🛒 Sales KPIs")

    # Penalties by Customer
    pen_cust = finance_r[finance_r["Metric"].str.contains("Bonus or penalties - Contracted sales revenue -")]
    if not pen_cust.empty:
        pen_cust = pen_cust.copy()
        pen_cust["Customer"] = pen_cust["Metric"].str.split("-").str[-1].str.strip()
        st.subheader("Penalties by Customer")
        st.dataframe(pen_cust[["Customer","Round 1","Round 2"]])
        plot_bar(pen_cust, "Customer", ["Round 1","Round 2"], "Penalties by Customer", "€")

    # Revenue by Customer
    rev_cust = finance_r[finance_r["Metric"].str.contains("Contracted sales revenue -")]
    if not rev_cust.empty:
        rev_cust = rev_cust.copy()
        rev_cust["Customer"] = rev_cust["Metric"].str.split("-").str[-1].str.strip()
        st.subheader("Revenue by Customer")
        st.dataframe(rev_cust[["Customer","Round 1","Round 2"]])
        plot_bar(rev_cust, "Customer", ["Round 1","Round 2"], "Revenue by Customer", "€")

# ----------------------------
# SUPPLY CHAIN TAB
# ----------------------------
with tabs[2]:
    st.header("📦 Supply Chain KPIs")

    comp = load_sheet("Supply chain.xlsx", "Component")

    if "Obsoletes (%)" in comp.columns:
        obs_summary = comp.groupby("Round")["Obsoletes (%)"].mean().reset_index()
        st.subheader("Obsolete Products (%)")
        st.dataframe(obs_summary)
        fig = px.bar(obs_summary, x="Round", y="Obsoletes (%)", title="Obsolete Products (%)")
        st.plotly_chart(fig, use_container_width=True)

    if "Stock value" in comp.columns:
        stock_summary = comp.groupby("Round")["Stock value"].sum().reset_index()
        st.subheader("Stock Value (€)")
        st.dataframe(stock_summary)
        fig = px.bar(stock_summary, x="Round", y="Stock value", title="Stock Value (€)")
        st.plotly_chart(fig, use_container_width=True)

    if "Component availability (%)" in comp.columns:
        avail_summary = comp.groupby("Round")["Component availability (%)"].mean().reset_index()
        st.subheader("Component Availability (%)")
        st.dataframe(avail_summary)
        fig = px.bar(avail_summary, x="Round", y="Component availability (%)", title="Component Availability (%)")
        st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# OPERATIONS TAB
# ----------------------------
with tabs[3]:
    st.header("🏭 Operations KPIs")

    bott = load_sheet("Operations.xlsx", "Bottling line")
    bott = bott[bott["Round"].isin([1,2])]
    bott_summary = bott.groupby("Round")[["Production plan adherence (%)","Overtime per week (hours)"]].mean().reset_index()

    st.subheader("Bottling KPIs")
    st.dataframe(bott_summary)
    fig = px.bar(bott_summary.melt(id_vars="Round"), x="Round", y="value", color="variable",
                 barmode="group", title="Bottling KPIs")
    st.plotly_chart(fig, use_container_width=True)

    ops_out = bott.groupby("Round")[["Run time (%)","Overtime (%)"]].mean().reset_index()
    st.subheader("Operational Outcomes")
    st.dataframe(ops_out)
    fig = px.bar(ops_out.melt(id_vars="Round"), x="Round", y="value", color="variable",
                 barmode="group", title="Operations Outcomes")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# PURCHASING TAB
# ----------------------------
with tabs[4]:
    st.header("📑 Purchasing KPIs")

    supc = load_sheet("Purchase.xlsx", "Supplier - Component")
    supc = supc[supc["Round"].isin([1,2])]

    if "Delivery reliability (%)" in supc.columns:
        rel_summary = supc.groupby("Round")["Delivery reliability (%)"].mean().reset_index()
        st.subheader("Supplier Reliability (%)")
        st.dataframe(rel_summary)
        fig = px.bar(rel_summary, x="Round", y="Delivery reliability (%)", title="Supplier Reliability (%)")
        st.plotly_chart(fig, use_container_width=True)

    if "Order size" in supc.columns:
        order_summary = supc.groupby("Round")["Order size"].mean().reset_index()
        st.subheader("Average Order Size")
        st.dataframe(order_summary)
        fig = px.bar(order_summary, x="Round", y="Order size", title="Order Size")
        st.plotly_chart(fig, use_container_width=True)

