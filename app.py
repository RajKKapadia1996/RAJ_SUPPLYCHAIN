
---

## **3. app.py**
```python
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

def extract_rounds(df):
    out = df[[df.columns[0], 1, 2]].copy()
    out.columns = ["Metric", "Round 1", "Round 2"]
    return out

def plot_bar(df, index_col, title, ylabel):
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

# ----------------------------
# NAVIGATION
# ----------------------------
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to:", ["Q1: Strategy", "Q2: Functional Decisions", "Q3: KPI Outcomes", "Q4: Next Round Plan"])

# ----------------------------
# Q1: Strategy
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
# Q2: Functional Decisions
# ----------------------------
elif page == "Q2: Functional Decisions":
    st.header("Q2. Functional Alignment – Sales, SCM, Operations, Purchasing")

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

# ----------------------------
# Q3: KPI Outcomes
# ----------------------------
elif page == "Q3: KPI Outcomes":
    st.header("Q3. KPI Outcomes – Achieved vs Not Achieved")

    st.subheader("Finance KPIs")
    st.dataframe(finance_r.head(15))
    plot_bar(finance_r.head(15), "Metric", "Finance KPIs (R1 vs R2)", "Value")

    obs = finance_r[finance_r["Metric"].str.contains("Obsolescence", case=False)]
    if not obs.empty:
        st.subheader("Obsolescence KPIs")
        st.dataframe(obs)
        plot_bar(obs, "Metric", "Obsolescence (R1 vs R2)", "€ / %")

# ----------------------------
# Q4: Next Round Plan
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
