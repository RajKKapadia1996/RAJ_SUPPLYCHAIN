import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os

st.set_page_config(page_title="Fresh Connection Dashboard", layout="wide")

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------
def load_finance():
    return pd.read_excel("FinanceReport.xlsx", sheet_name="Output")

def load_file(name, sheet=None):
    path = os.path.join(".", name)
    if sheet:
        return pd.read_excel(path, sheet_name=sheet)
    return pd.ExcelFile(path)

def extract_rounds(df):
    """Extract only Round 1 and Round 2 values"""
    out = df[[df.columns[0], 1, 2]].copy()
    out.columns = ["Metric", "Round 1", "Round 2"]
    return out

def plot_bar(df, index_col, title, ylabel):
    """Draws bar chart for Round 1 vs Round 2"""
    fig, ax = plt.subplots(figsize=(6,4))
    df.set_index(index_col)[["Round 1","Round 2"]].plot(kind="bar", ax=ax)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    st.pyplot(fig)

# ----------------------------
# LOAD DATA
# ----------------------------
finance = load_finance()
finance_r = extract_rounds(finance)

ops = load_file("Operations.xlsx")
purch = load_file("Purchase.xlsx")
sales = load_file("Sales.xlsx")
scm = load_file("Supply chain.xlsx")

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
    plot_bar(roi, "Metric", "ROI (R1 vs R2)", "ROI")

    st.subheader("Total Penalties")
    st.dataframe(penalties_total)
    plot_bar(penalties_total, "Metric", "Total Penalties", "€")

    st.subheader("Net Realized Revenue")
    st.dataframe(realized_rev)
    plot_bar(realized_rev, "Metric", "Realized Revenue", "€")

# ----------------------------
# Q2: FUNCTIONAL DECISIONS
# ----------------------------
elif page == "Q2: Functional Decisions":
    st.header("Q2. Functional Alignment – Sales, SCM, Operations, Purchasing")

    # SALES: Penalties + Revenue by Customer
    pen_cust = finance_r[finance_r["Metric"].str.contains("Bonus or penalties - Contracted sales revenue -")]
    if not pen_cust.empty:
        pen_cust = pen_cust.copy()
        pen_cust["Customer"] = pen_cust["Metric"].str.split("-").str[-1].str.strip()
        st.subheader("Sales – Penalties by Customer")
        st.dataframe(pen_cust[["Customer","Round 1","Round 2"]])
        plot_bar(pen_cust, "Customer", "Penalties by Customer", "€")

    rev_cust = finance_r[finance_r["Metric"].str.contains("Contracted sales revenue -")]
    if not rev_cust.empty:
        rev_cust = rev_cust.copy()
        rev_cust["Customer"] = rev_cust["Metric"].str.split("-").str[-1].str.strip()
        st.subheader("Sales – Revenue by Customer")
        st.dataframe(rev_cust[["Customer","Round 1","Round 2"]])
        plot_bar(rev_cust, "Customer", "Revenue by Customer", "€")

    # PURCHASING: Reliability + Order Size
    supc = pd.read_excel("Purchase.xlsx", sheet_name="Supplier - Component")
    if "Delivery reliability (%)" in supc.columns:
        rel_summary = supc.groupby("Round")["Delivery reliability (%)"].mean().reset_index()
        rel_summary = rel_summary.set_index("Round").T
        st.subheader("Purchasing – Average Supplier Reliability")
        st.dataframe(rel_summary)
        plot_bar(rel_summary.T.reset_index(), "Round", "Supplier Reliability", "%")

    if "Order size" in supc.columns:
        order_summary = supc.groupby("Round")["Order size"].mean().reset_index()
        order_summary = order_summary.set_index("Round").T
        st.subheader("Purchasing – Average Order Size")
        st.dataframe(order_summary)
        plot_bar(order_summary.T.reset_index(), "Round", "Order Size", "Units")

    # OPERATIONS: Bottling line KPIs
    bott = pd.read_excel("Operations.xlsx", sheet_name="Bottling line")
    bott_summary = bott.groupby("Round")[["Production plan adherence (%)","Overtime per week (hours)"]].mean().reset_index()
    st.subheader("Operations – Bottling Line KPIs")
    st.dataframe(bott_summary)
    plot_bar(bott_summary.set_index("Round"), bott_summary.set_index("Round").columns.name or "Round", "Bottling KPIs", "Value")

    # SUPPLY CHAIN: Obsolescence + Stock Value
    comp = pd.read_excel("Supply chain.xlsx", sheet_name="Component")
    if "Obsoletes (%)" in comp.columns:
        obs_summary = comp.groupby("Round")["Obsoletes (%)"].mean().reset_index()
        st.subheader("Supply Chain – Obsolete Products (%)")
        st.dataframe(obs_summary)
        plot_bar(obs_summary.set_index("Round"), obs_summary.set_index("Round").columns.name or "Round", "Obsolete %", "%")

    if "Stock value" in comp.columns:
        stock_summary = comp.groupby("Round")["Stock value"].sum().reset_index()
        st.subheader("Supply Chain – Stock Value (€)")
        st.dataframe(stock_summary)
        plot_bar(stock_summary.set_index("Round"), stock_summary.set_index("Round").columns.name or "Round", "Stock Value", "€")

# ----------------------------
# Q3: KPI OUTCOMES
# ----------------------------
elif page == "Q3: KPI Outcomes":
    st.header("Q3. KPI Outcomes – Achieved vs Not Achieved")

    st.subheader("Finance KPIs")
    st.dataframe(finance_r.head(15))
    plot_bar(finance_r.head(15), "Metric", "Finance KPIs (R1 vs R2)", "Value")

    # Ops outcomes
    bott = pd.read_excel("Operations.xlsx", sheet_name="Bottling line")
    ops_out = bott.groupby("Round")[["Run time (%)","Overtime (%)","Production plan adherence (%)"]].mean().reset_index()
    st.subheader("Operations – Outcomes")
    st.dataframe(ops_out)
    plot_bar(ops_out.set_index("Round"), "Round", "Ops Outcomes", "%")

    # Supply chain outcomes
    comp = pd.read_excel("Supply chain.xlsx", sheet_name="Component")
    scm_out = comp.groupby("Round")[["Delivery reliability (%)","Component availability (%)"]].mean().reset_index()
    st.subheader("Supply Chain – Reliability & Availability")
    st.dataframe(scm_out)
    plot_bar(scm_out.set_index("Round"), "Round", "SCM Outcomes", "%")

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

