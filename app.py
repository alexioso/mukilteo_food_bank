import streamlit as st

st.set_page_config(page_title="MFB Distribution Dashboard", layout="wide")

visuals = st.Page("pages/visuals.py", title="Visuals", default=True)
month_end_summary = st.Page("pages/month_end_summary.py", title="Month End Summary")

pg = st.navigation([visuals, month_end_summary])
pg.run()
