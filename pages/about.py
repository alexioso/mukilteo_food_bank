import streamlit as st

MUTED = "#898781"

st.title("About")

st.markdown(
    "This dashboard is Metro Food Bank's distribution reporting pipeline — from the "
    "raw logs staff enter in FoodBank Manager on each service day, to the charts, "
    "monthly summary, and forecasts elsewhere in this app. Every stage of it runs "
    "on a free tier, with no server to maintain and no database to pay for."
)

st.subheader("The pipeline, end to end")

steps = ["FoodBank Manager", "Browserbase", "GitHub Actions", "GitHub repo", "Streamlit Cloud"]
cols = st.columns(len(steps) * 2 - 1)
for i, step in enumerate(steps):
    cols[i * 2].markdown(f"<div style='text-align:center; font-weight:600;'>{step}</div>", unsafe_allow_html=True)
    if i < len(steps) - 1:
        cols[i * 2 + 1].markdown(
            f"<div style='text-align:center; color:{MUTED}; font-size:1.4rem;'>&rarr;</div>", unsafe_allow_html=True
        )

st.write("")

with st.container(border=True):
    st.markdown("**1. FoodBank Manager — the source of truth**")
    st.markdown(
        "Staff log every distribution day directly into FoodBank Manager "
        "(`mfbfp.soxbox.co`): household visits, individuals served by age group, "
        "pounds distributed, and volunteer time entries. This dashboard only ever "
        "*reads* from it — nothing here writes back."
    )

with st.container(border=True):
    st.markdown("**2. Browserbase — automated data extraction**")
    st.markdown(
        "FoodBank Manager doesn't expose a data API, so `src/mandatory_report_refresh.py` "
        "drives a real, cloud-hosted browser session through "
        "[Browserbase](https://www.browserbase.com/) (via Selenium's remote WebDriver "
        "protocol) to log in, load each report (Total, Duplicated, Unduplicated, Time "
        "Entry), export it, and pull the resulting file down through Browserbase's "
        "Downloads API. Because the browser itself runs in Browserbase's cloud, the "
        "scrape doesn't depend on any laptop being turned on or having Chrome installed."
    )

with st.container(border=True):
    st.markdown("**3. GitHub Actions — scheduling & compute**")
    st.markdown(
        "`.github/workflows/main.yml` runs the scrape and the forecast refresh on a "
        "schedule, entirely on GitHub's free-tier runners. This replaced the original "
        "setup (`main_refresh.sh`), which required running the pipeline manually from "
        "a local machine with Chrome and a conda environment installed."
    )

with st.container(border=True):
    st.markdown("**4. Data processing & forecasting**")
    st.markdown(
        "The raw exports are merged and cleaned into `data/processed/df_monthly.csv` "
        "and `df_weekly.csv`. `src/generate_forecast.py` also refits the ARIMAX "
        "household-visits forecast (see the Forecast page) and writes the result to "
        "`data/processed/forecast_cache.json`, so the live app never has to run the "
        "model itself — it just reads a precomputed file."
    )

with st.container(border=True):
    st.markdown("**5. The GitHub repo — the database**")
    st.markdown(
        "Processed CSVs and the forecast cache get committed straight back into this "
        "repo by the GitHub Actions job. There's no database server: git itself is "
        "the storage layer, and every refresh is a commit — so the full history of "
        "the data is just `git log`."
    )

with st.container(border=True):
    st.markdown("**6. Streamlit Community Cloud — this app**")
    st.markdown(
        "This dashboard is hosted on [Streamlit Community Cloud](https://share.streamlit.io/)'s "
        "free tier, reading those files directly out of the repo. When new data lands "
        "on the default branch, the app picks it up on its next reboot/redeploy."
    )

st.subheader("Why it's all free")
st.info(
    "GitHub Actions' scheduled runners, Browserbase's cloud browser sessions, and "
    "Streamlit Community Cloud's app hosting are each used within that service's "
    "free tier. A biweekly-ish distribution schedule means a light, infrequent "
    "workload — well within what each free tier allows — so the whole pipeline "
    "runs at no ongoing cost.",
    icon="💡",
)

with st.expander("Maintainer notes — key files"):
    st.markdown(
        """
| File | Purpose |
|---|---|
| `.github/workflows/main.yml` | Scheduled cloud run: scrape → forecast → commit → push |
| `main_refresh.sh` | The original local/manual equivalent (conda env, Selenium on local Chrome) |
| `src/mandatory_report_refresh.py` | Browserbase-driven scrape of FoodBank Manager + data merge into `data/processed/` |
| `src/generate_forecast.py` | Fits the ARIMAX forecast + walk-forward backtest, writes `forecast_cache.json` |
| `data/raw/` | Exported reports as pulled from FoodBank Manager |
| `data/processed/` | Cleaned, merged data the app actually reads |
| `data/raw/next_distribution_override.txt` | Manual override for the next confirmed distribution date, when it differs from the model's cadence guess |
"""
    )
