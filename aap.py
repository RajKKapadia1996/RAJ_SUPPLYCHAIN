import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="Fresh Connection — R1–R3 Metrics Dashboard", layout="wide")
st.title("Fresh Connection — R1–R3 Metrics Dashboard")
st.caption("Tables and visualizations for the exact policy metrics from the screenshot: Sales, Supply Chain, Operations, Purchasing (Rounds 1–3).")

# -------------------------------
# Data loading
# -------------------------------
@st.cache_data
def load_excel_from_bytes(file_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}

@st.cache_data
def load_excel_from_path(path: str):
    xls = pd.ExcelFile(path)
    return {name: pd.read_excel(xls, name) for name in xls.sheet_names}

def get_data():
    default_path = "metrics.xlsx"
    try:
        sheets = load_excel_from_path(default_path)
        st.success("Loaded metrics.xlsx from the app folder.")
        return sheets
    except Exception:
        st.warning("metrics.xlsx not found. Upload the Excel file below.")
        upl = st.file_uploader("Upload metrics.xlsx", type=["xlsx"])
        if upl is not None:
            return load_excel_from_bytes(upl.read())
        else:
            st.stop()

sheets = get_data()

# -------------------------------
# Sidebar filters
# -------------------------------
st.sidebar.header("Filters")
ALL_ROUNDS = ["Round1", "Round2", "Round3"]
round_select = st.sidebar.multiselect("Rounds to include in charts", ALL_ROUNDS, default=ALL_ROUNDS)
if not round_select:
    st.sidebar.error("Select at least one round.")
    st.stop()

# -------------------------------
# Helper: simple line plot across rounds
# -------------------------------
def plot_round_line(series_dict, title, ylabel):
    """
    series_dict: dict(label -> dict(round -> value))
    """
    fig, ax = plt.subplots()
    for label, rd in series_dict.items():
        # keep only selected rounds
        xs = [r for r in ALL_ROUNDS if r in rd and r in round_select]
        ys = [rd[r] for r in xs]
        ax.plot(xs, ys, marker="o", label=str(label))
    ax.set_title(title)
    ax.set_xlabel("Round")
    ax.set_ylabel(ylabel)
    ax.legend(loc="best", bbox_to_anchor=(1.0, 1.0))
    st.pyplot(fig)

# -------------------------------
# SALES
# -------------------------------
st.markdown("## Sales")
df_sales = sheets["Sales"].copy()
st.dataframe(df_sales, use_container_width=True)

# Visualize numeric Sales metrics per customer
for cust in df_sales["Customer"].unique():
    sub = df_sales[df_sales["Customer"] == cust]
    # Service level target
    s_row = sub[sub["Metric"] == "Service level target"]
    if not s_row.empty:
        row = s_row.iloc[0]
        data = {cust: {r: row[r] for r in ALL_ROUNDS}}
        plot_round_line(data, f"Service Level Target — {cust}", "%")

    # Shelf-life requirement
    sh_row = sub[sub["Metric"] == "Shelf-life requirement"]
    if not sh_row.empty:
        row = sh_row.iloc[0]
        data = {cust: {r: row[r] for r in ALL_ROUNDS}}
        plot_round_line(data, f"Shelf-life Requirement — {cust}", "%")

st.markdown("---")

# -------------------------------
# SUPPLY CHAIN (Planning & Inventory)
# -------------------------------
st.markdown("## Supply Chain (Planning & Inventory)")
df_sc = sheets["SupplyChain"].copy()
st.dataframe(df_sc, use_container_width=True)

# RM Safety Stock by component (weeks)
ss = df_sc[(df_sc["Category"] == "Raw Material") & (df_sc["Metric"] == "Safety stock")]
ss_series = {}
for comp in ss["Item"].unique():
    row = ss[ss["Item"] == comp][["Round1","Round2","Round3"]].iloc[0].to_dict()
    ss_series[comp] = row
plot_round_line(ss_series, "Raw Material Safety Stock (weeks) — by Component", "weeks")

# RM Lot sizes by component (weeks)
ls = df_sc[(df_sc["Category"] == "Raw Material") & (df_sc["Metric"] == "Lot size")]
ls_series = {}
for comp in ls["Item"].unique():
    row = ls[ls["Item"] == comp][["Round1","Round2","Round3"]].iloc[0].to_dict()
    ls_series[comp] = row
plot_round_line(ls_series, "Raw Material Lot Size (weeks) — by Component", "weeks")

# Global planning knobs
glob = df_sc[df_sc["Category"] == "Planning"]
for metric in glob["Metric"].unique():
    row = glob[glob["Metric"] == metric][["Round1","Round2","Round3"]].iloc[0].to_dict()
    plot_round_line({metric: row}, f"{metric}", glob[glob['Metric']==metric]['Unit'].iloc[0])

# FG SS by family
fg = df_sc[(df_sc["Category"] == "Finished Goods") & (df_sc["Metric"] == "Safety stock")]
fg_series = {}
for fam in fg["Item"].unique():
    row = fg[fg["Item"] == fam][["Round1","Round2","Round3"]].iloc[0].to_dict()
    fg_series[fam] = row
plot_round_line(fg_series, "Finished Goods Safety Stock (weeks) — by Family", "weeks")

st.markdown("---")

# -------------------------------
# OPERATIONS
# -------------------------------
st.markdown("## Operations")
df_ops = sheets["Operations"].copy()
st.dataframe(df_ops, use_container_width=True)

# Visualize numeric Ops metrics
# Breakdown trailing (weekly, hours)
try:
    row = df_ops[df_ops["Metric"]=="Breakdown trailing (weekly)"][["Round1","Round2","Round3"]].iloc[0].to_dict()
    plot_round_line({"Breakdown trailing (weekly)": row}, "Breakdown Trailing (weekly)", "hours")
except Exception:
    pass

# Capacities (Inbound/Outbound WH, shifts)
for metric, ylabel in [
    ("Inbound WH capacity", "pallet locations"),
    ("Outbound WH capacity", "pallet locations"),
    ("Number of shifts (bottling)", "count"),
]:
    try:
        row = df_ops[df_ops["Metric"]==metric][["Round1","Round2","Round3"]].iloc[0].to_dict()
        plot_round_line({metric: row}, metric, ylabel)
    except Exception:
        pass

st.markdown("---")

# -------------------------------
# PURCHASING
# -------------------------------
st.markdown("## Purchasing")
df_purch = sheets["Purchasing"].copy()
st.dataframe(df_purch, use_container_width=True)

# Lead time (days) by component
lt = df_purch[df_purch["Metric"] == "Lead time"]
lt_series = {}
for comp in lt["Component"].unique():
    row = lt[(lt["Component"] == comp)][["Round1","Round2","Round3"]].iloc[0].to_dict()
    lt_series[comp] = row
plot_round_line(lt_series, "Lead Time (days) — by Component", "days")

# Delivery window (days) by component
dw = df_purch[df_purch["Metric"] == "Delivery window"]
dw_series = {}
for comp in dw["Component"].unique():
    row = dw[(dw["Component"] == comp)][["Round1","Round2","Round3"]].iloc[0].to_dict()
    dw_series[comp] = row
plot_round_line(dw_series, "Delivery Window (days) — by Component", "days")

st.success("Ready. Use the sidebar to include/exclude rounds. Tables are sortable; charts show numeric metrics per round.")
