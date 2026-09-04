import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

DATA_DIR = Path(__file__).parent.parent / "data"
CACHE_PATH = DATA_DIR / "processed" / "forecast_cache.json"

PALETTE = {"blue": "#2a78d6", "orange": "#eb6834"}
SURFACE = "#fcfcfb"
INK = "#0b0b0b"
MUTED = "#898781"
AXIS = "#c3c2b7"
GRID = "#e1e0d9"
FONT = "system-ui, -apple-system, 'Segoe UI', sans-serif"


@st.cache_data(ttl=3600)
def load_forecast_cache() -> dict:
    with open(CACHE_PATH) as f:
        return json.load(f)


def style_chart(fig: go.Figure, y_title: str = "") -> go.Figure:
    fig.update_layout(
        plot_bgcolor=SURFACE,
        paper_bgcolor=SURFACE,
        font=dict(color=INK, family=FONT),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
            font=dict(color=INK, family=FONT), bgcolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=40, r=20, t=40, b=40),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, linecolor=AXIS, tickfont=dict(color=MUTED))
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False, title=y_title, tickfont=dict(color=MUTED))
    return fig


st.title("Forecast")
st.caption(
    "ARIMAX models (auto-fit order, exogenous day-of-week / month / US-holiday-proximity "
    "regressors) fit on Monday & Tuesday service days since 2024-01-01, backtested via "
    "walk-forward validation against a same-weekday-last-time naive baseline."
)

if not CACHE_PATH.exists():
    st.error(
        "No forecast cache found. Run `python src/generate_forecast.py` to generate "
        "data/processed/forecast_cache.json, then refresh this page."
    )
    st.stop()

cache = load_forecast_cache()
generated_at = datetime.fromisoformat(cache["generated_at"])
st.caption(f"Forecast last generated {generated_at.strftime('%B %d, %Y %H:%M UTC')} — precomputed in the backend, not on page load.")

next_event = cache.get("next_distribution_event")
if next_event and "total_hh_visits" in next_event["metrics"]:
    monday = datetime.fromisoformat(next_event["monday"])
    tuesday = datetime.fromisoformat(next_event["tuesday"])
    hh = next_event["metrics"]["total_hh_visits"]

    with st.container(border=True):
        st.subheader(f"📅 Next Distribution Day: {monday.strftime('%A, %B %d')} & {tuesday.strftime('%A, %B %d, %Y')}")
        st.caption(
            f"Best guess based on the typical {cache.get('typical_gap_days', 14)}-day gap between past "
            "distribution events — usually every two weeks, with occasional exceptions, so treat this "
            "as a planning estimate rather than a confirmed date."
        )

        big_cols = st.columns(2)
        for c, day_key, day_dt in zip(big_cols, ["monday", "tuesday"], [monday, tuesday]):
            d = hh[day_key]
            with c:
                st.markdown(
                    f"""
                    <div style="text-align:center; padding: 8px 0 16px;">
                        <div style="font-size:1.1rem; font-weight:600; color:{MUTED};">{day_dt.strftime('%A, %b %d')}</div>
                        <div style="font-size:3.5rem; font-weight:700; line-height:1.1;">{d['forecast']:,}</div>
                        <div style="font-size:0.95rem; color:{MUTED};">household visits &middot; 95% CI {d['low_95']:,}&ndash;{d['high_95']:,}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=[monday.strftime("%a %b %d"), tuesday.strftime("%a %b %d")],
            y=[hh["monday"]["forecast"], hh["tuesday"]["forecast"]],
            marker_color=[PALETTE["blue"], PALETTE["orange"]],
            error_y=dict(
                type="data", symmetric=False,
                array=[hh["monday"]["high_95"] - hh["monday"]["forecast"], hh["tuesday"]["high_95"] - hh["tuesday"]["forecast"]],
                arrayminus=[hh["monday"]["forecast"] - hh["monday"]["low_95"], hh["tuesday"]["forecast"] - hh["tuesday"]["low_95"]],
                color=MUTED, thickness=2, width=10,
            ),
        ))
        chart = style_chart(fig, "Household Visits (95% CI)")
        chart.update_layout(showlegend=False)
        st.plotly_chart(chart, use_container_width=True, theme=None)
    st.divider()

for colname, metric in cache["metrics"].items():
    st.subheader(metric["label"])

    forecast_df = pd.DataFrame(metric["forecast"])
    forecast_df["date"] = pd.to_datetime(forecast_df["date"])

    cols = st.columns(len(forecast_df))
    for c, (_, row) in zip(cols, forecast_df.iterrows()):
        c.metric(f"{row['day']} {row['date'].strftime('%b %d')}", f"{row['forecast']:,}")
        c.caption(f"95% CI: {row['low_95']:,}–{row['high_95']:,}")

    history_df = pd.DataFrame(metric["history"])
    history_df["date"] = pd.to_datetime(history_df["date"])

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=history_df["date"], y=history_df["value"], mode="lines+markers", name="Actual",
        line=dict(color=PALETTE["blue"], width=2),
    ))
    fig.add_trace(go.Scatter(
        x=forecast_df["date"], y=forecast_df["forecast"], mode="lines+markers", name="Forecast",
        line=dict(color=PALETTE["orange"], width=2, dash="dash"),
    ))
    fig.add_trace(go.Scatter(
        x=pd.concat([forecast_df["date"], forecast_df["date"][::-1]]),
        y=pd.concat([forecast_df["high_95"], forecast_df["low_95"][::-1]]),
        fill="toself", fillcolor="rgba(235,104,52,0.15)", line=dict(color="rgba(0,0,0,0)"),
        name="95% interval",
    ))
    st.plotly_chart(style_chart(fig, metric["label"]), use_container_width=True, theme=None)

    if metric["backtest"]:
        st.caption(f"Backtest accuracy — walk-forward over last {metric['backtest_n']} service days")
        backtest_df = pd.DataFrame(metric["backtest"]).rename(
            columns={"method": "Method", "mae": "MAE", "rmse": "RMSE", "mape": "MAPE %"}
        )
        st.dataframe(backtest_df, hide_index=True, use_container_width=True)
    else:
        st.info("Not enough history yet for a reliable backtest.")

    st.divider()
