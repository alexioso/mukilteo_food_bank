import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"

PALETTE = {
    "blue": "#2a78d6",
    "orange": "#eb6834",
    "aqua": "#1baf7a",
    "yellow": "#eda100",
    "magenta": "#e87ba4",
}
SURFACE = "#fcfcfb"
INK = "#0b0b0b"
MUTED = "#898781"
AXIS = "#c3c2b7"
GRID = "#e1e0d9"
FONT = "system-ui, -apple-system, 'Segoe UI', sans-serif"


@st.cache_data(ttl=3600)
def load_monthly() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "processed" / "df_monthly.csv")
    df["year_month"] = pd.to_datetime(df["year_month"])
    return df.sort_values("year_month")


@st.cache_data(ttl=3600)
def load_weekly() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "processed" / "df_weekly.csv")
    df["week_start"] = pd.to_datetime(df["week_start"])
    return df.sort_values("week_start")


def style_chart(fig: go.Figure, y_title: str = "", show_legend: bool = True) -> go.Figure:
    fig.update_layout(
        plot_bgcolor=SURFACE,
        paper_bgcolor=SURFACE,
        font=dict(color=INK, family=FONT),
        showlegend=show_legend,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
            font=dict(color=INK, family=FONT), bgcolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=40, r=20, t=40, b=40),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, linecolor=AXIS, tickfont=dict(color=MUTED))
    fig.update_yaxes(
        showgrid=True, gridcolor=GRID, zeroline=False,
        title=y_title, tickfont=dict(color=MUTED),
    )
    return fig


df_monthly = load_monthly()
df_weekly = load_weekly()

st.title("Metro Food Bank — Distribution Dashboard")

latest = df_monthly.iloc[-1]
prior = df_monthly.iloc[-2] if len(df_monthly) > 1 else None

def delta_str(col: str) -> str | None:
    if prior is None:
        return None
    return f"{latest[col] - prior[col]:+,.0f}"


c1, c2, c3, c4 = st.columns(4)
c1.metric("HH Visits (latest month)", f"{latest['total_hh_visits']:,.0f}", delta=delta_str("total_hh_visits"))
c2.metric("Individuals Served", f"{latest['total_indivdiduals']:,.0f}", delta=delta_str("total_indivdiduals"))
c3.metric("Weight Distributed (lbs)", f"{latest['total_weight']:,.0f}", delta=delta_str("total_weight"))
c4.metric("Volunteer Hours (month)", f"{latest['volunteer_hours']:,.0f}", delta=delta_str("volunteer_hours"))

st.caption(f"Latest month: {latest['year_month'].strftime('%B %Y')}")

tab1, tab2, tab3, tab4 = st.tabs(
    ["Monthly Trends", "Age Breakdown", "Volunteers", "Weekly Visits"]
)

with tab1:
    st.subheader("Household Visits by Month")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_monthly["year_month"], y=df_monthly["total_hh_visits"],
        mode="lines", name="Total", line=dict(color=PALETTE["blue"], width=2),
    ))
    fig.add_trace(go.Scatter(
        x=df_monthly["year_month"], y=df_monthly["undup_hh_visits"],
        mode="lines", name="Unduplicated", line=dict(color=PALETTE["aqua"], width=2),
    ))
    st.plotly_chart(style_chart(fig, "HH Visits"), use_container_width=True, theme=None)

    st.subheader("Weight Distributed by Month")
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(
        x=df_monthly["year_month"], y=df_monthly["total_weight"],
        mode="lines", fill="tozeroy", name="Weight (lbs)",
        line=dict(color=PALETTE["blue"], width=2),
    ))
    st.plotly_chart(style_chart(fig2, "Pounds", show_legend=False), use_container_width=True, theme=None)

with tab2:
    st.subheader("Individuals Served by Age Group")
    age_cols = [
        ("total_age_0_to_2", "0–2", PALETTE["blue"]),
        ("total_age_3_to_18", "3–18", PALETTE["orange"]),
        ("total_age_19_to_54", "19–54", PALETTE["aqua"]),
        ("total_age_55_plus", "55+", PALETTE["yellow"]),
        ("total_age_anonymous", "Not provided", PALETTE["magenta"]),
    ]
    fig = go.Figure()
    for col, label, color in age_cols:
        fig.add_trace(go.Bar(x=df_monthly["year_month"], y=df_monthly[col], name=label, marker_color=color))
    fig.update_layout(barmode="stack")
    st.plotly_chart(style_chart(fig, "Individuals"), use_container_width=True, theme=None)

with tab3:
    st.subheader("Volunteer Hours by Month")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_monthly["year_month"], y=df_monthly["volunteer_hours"],
        name="Monthly Hours", marker_color=PALETTE["blue"],
    ))
    st.plotly_chart(style_chart(fig, "Hours", show_legend=False), use_container_width=True, theme=None)

    st.subheader("Volunteer Count by Month")
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(
        x=df_monthly["year_month"], y=df_monthly["volunteer_count"],
        mode="lines+markers", name="Volunteers",
        line=dict(color=PALETTE["aqua"], width=2), marker=dict(size=8),
    ))
    st.plotly_chart(style_chart(fig2, "Volunteers", show_legend=False), use_container_width=True, theme=None)

with tab4:
    st.subheader("Weekly Household Counts")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_weekly["week_start"], y=df_weekly["Monday"],
        mode="lines", name="Monday", line=dict(color=PALETTE["blue"], width=2),
    ))
    fig.add_trace(go.Scatter(
        x=df_weekly["week_start"], y=df_weekly["Tuesday"],
        mode="lines", name="Tuesday", line=dict(color=PALETTE["orange"], width=2),
    ))
    fig.add_trace(go.Scatter(
        x=df_weekly["week_start"], y=df_weekly["Total"],
        mode="lines", name="Total", line=dict(color=PALETTE["aqua"], width=2),
    ))
    st.plotly_chart(style_chart(fig, "Counts"), use_container_width=True, theme=None)

with st.expander("Raw monthly data"):
    st.dataframe(df_monthly, use_container_width=True)
