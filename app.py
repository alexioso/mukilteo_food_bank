import streamlit as st

st.set_page_config(page_title="MFB Distribution Dashboard", layout="wide")

month_end_summary = st.Page("pages/month_end_summary.py", title="Month End Summary", default=True)
visuals = st.Page("pages/visuals.py", title="Visuals")
forecast = st.Page("pages/forecast.py", title="Forecast")

pg = st.navigation([month_end_summary, visuals, forecast])
pg.run()
