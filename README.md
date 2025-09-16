# Fresh Connection — R1–R3 Metrics Dashboard (Streamlit)

This app shows **only the policy metrics from your screenshot** for **Sales, Supply Chain, Operations, and Purchasing**, across **Round 1, Round 2, Round 3**. It includes sortable tables and simple visualizations for numeric metrics.

## Files
- `aap.py` — Streamlit app
- `metrics.xlsx` — Data file with R1–R3 metrics (put this next to `aap.py`)
- `requirements.txt` — Python dependencies

## Quickstart
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
pip install -r requirements.txt
streamlit run aap.py
