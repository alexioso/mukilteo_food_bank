import streamlit as st
import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"

TABLE_CSS = """
<style>
.mes-table { border-collapse: collapse; width: 100%; margin-bottom: 8px; }
.mes-table td, .mes-table th {
    border: 1px solid rgba(130,130,130,0.55);
    padding: 6px 14px;
    text-align: center;
}
.mes-table th { font-weight: 700; }
.mes-table td.label { text-align: left; font-weight: 600; }
.mes-table td.indent { text-align: left; padding-left: 28px; }
.mes-table td.section {
    text-align: left;
    font-weight: 700;
    border: 2px solid rgba(130,130,130,0.85);
}
.mes-table tr.total td { font-weight: 700; border-top: 2px solid rgba(130,130,130,0.85); }
.mes-spacer { height: 18px; }
</style>
"""


@st.cache_data(ttl=3600)
def load_monthly() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "processed" / "df_monthly.csv")
    df["year_month"] = pd.to_datetime(df["year_month"])
    return df.sort_values("year_month").fillna(0)


@st.cache_data(ttl=3600)
def load_daily() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "raw" / "total_report_daily.csv")
    df["date"] = pd.to_datetime(df["date"])
    return df.sort_values("date").fillna(0)


def fmt(n) -> str:
    return f"{n:,.0f}"


def render_table(html: str) -> None:
    # Collapse to a single line so Streamlit's markdown parser treats the
    # whole thing as one HTML block instead of splitting on embedded
    # newlines/indentation (which renders as half-table, half raw text).
    st.markdown(" ".join(html.split()), unsafe_allow_html=True)


st.markdown(TABLE_CSS, unsafe_allow_html=True)
st.title("Month End Summary")

df_monthly = load_monthly()
df_daily = load_daily()

month_options = df_monthly["year_month"].dt.strftime("%Y-%m").tolist()
selected_label = st.selectbox("Select Month:", month_options, index=len(month_options) - 1)
selected_month = pd.Timestamp(selected_label + "-01")

month_row = df_monthly.loc[df_monthly["year_month"] == selected_month].iloc[0]
ytd_df = df_monthly[
    (df_monthly["year_month"].dt.year == selected_month.year)
    & (df_monthly["year_month"] <= selected_month)
]

daily_month = df_daily[
    (df_daily["date"].dt.year == selected_month.year) & (df_daily["date"].dt.month == selected_month.month)
]
daily_ytd = df_daily[
    (df_daily["date"].dt.year == selected_month.year) & (df_daily["date"] <= selected_month + pd.offsets.MonthEnd(0))
]

# ---- Individuals by age group ----
age_groups = [
    ("Infants (0-2)", "dup_age_0_to_2", "undup_age_0_to_2"),
    ("Children (3-18)", "dup_age_3_to_18", "undup_age_3_to_18"),
    ("Adults (19-54)", "dup_age_19_to_54", "undup_age_19_to_54"),
    ("Seniors (55 Plus)", "dup_age_55_plus", "undup_age_55_plus"),
    ("Anonymous", "dup_age_anonymous", "undup_age_anonymous"),
]

rows_html = ""
for label, dup_col, undup_col in age_groups:
    m_dup, m_undup = month_row[dup_col], month_row[undup_col]
    y_dup, y_undup = ytd_df[dup_col].sum(), ytd_df[undup_col].sum()
    rows_html += f"""
    <tr>
        <td class="indent">{label}</td>
        <td>{fmt(m_dup)}</td><td>{fmt(m_undup)}</td><td>{fmt(m_dup + m_undup)}</td>
        <td>{fmt(y_dup)}</td><td>{fmt(y_undup)}</td><td>{fmt(y_dup + y_undup)}</td>
    </tr>"""

m_ind_dup, m_ind_undup = month_row["dup_indivdiduals"], month_row["undup_indivdiduals"]
y_ind_dup, y_ind_undup = ytd_df["dup_indivdiduals"].sum(), ytd_df["undup_indivdiduals"].sum()

individuals_table = f"""
<table class="mes-table">
<tr>
    <th class="section">INDIVIDUALS</th>
    <th colspan="3">Month</th>
    <th colspan="3">Year to Date</th>
</tr>
<tr>
    <th class="label">By Age Groups</th>
    <th>Duplicate</th><th>Unduplicated</th><th>Total</th>
    <th>Duplicate</th><th>Unduplicated</th><th>Total</th>
</tr>
{rows_html}
<tr class="total">
    <td class="indent">Total Individuals</td>
    <td>{fmt(m_ind_dup)}</td><td>{fmt(m_ind_undup)}</td><td>{fmt(m_ind_dup + m_ind_undup)}</td>
    <td>{fmt(y_ind_dup)}</td><td>{fmt(y_ind_undup)}</td><td>{fmt(y_ind_dup + y_ind_undup)}</td>
</tr>
</table>
"""

# ---- Households ----
m_hh_dup, m_hh_undup = month_row["dup_hh_visits"], month_row["undup_hh_visits"]
y_hh_dup, y_hh_undup = ytd_df["dup_hh_visits"].sum(), ytd_df["undup_hh_visits"].sum()

households_table = f"""
<table class="mes-table">
<tr>
    <th class="section">HOUSEHOLDS</th>
    <th colspan="3">Month</th>
    <th colspan="3">Year to Date</th>
</tr>
<tr>
    <th class="label"></th>
    <th>Duplicate</th><th>Unduplicated</th><th>Total</th>
    <th>Duplicate</th><th>Unduplicated</th><th>Total</th>
</tr>
<tr>
    <td class="indent">Total Served</td>
    <td>{fmt(m_hh_dup)}</td><td>{fmt(m_hh_undup)}</td><td>{fmt(m_hh_dup + m_hh_undup)}</td>
    <td>{fmt(y_hh_dup)}</td><td>{fmt(y_hh_undup)}</td><td>{fmt(y_hh_dup + y_hh_undup)}</td>
</tr>
</table>
"""

# ---- Pounds served ----
month_lbs = daily_month["total_weight"].sum()
ytd_lbs = daily_ytd["total_weight"].sum()
month_days = int((daily_month["total_hh_visits"] > 0).sum())
ytd_days = int((daily_ytd["total_hh_visits"] > 0).sum())

pounds_table = f"""
<table class="mes-table">
<tr><th class="section">POUNDS SERVED</th><th>Month</th><th>Year to Date</th></tr>
<tr><td class="indent">Total LBS Food Served</td><td>{fmt(month_lbs)}</td><td>{fmt(ytd_lbs)}</td></tr>
<tr><td class="indent">Number of Days Food Served</td><td>{fmt(month_days)}</td><td>{fmt(ytd_days)}</td></tr>
</table>
"""

# ---- Volunteers ----
v_count, v_hours = month_row["volunteer_count"], month_row["volunteer_hours"]
v_count_ytd, v_hours_ytd = month_row["volunteer_count_ytd"], month_row["volunteer_hours_ytd"]

volunteers_table = f"""
<table class="mes-table">
<tr><th class="section">VOLUNTEERS</th><th>Month</th><th>Year to Date</th></tr>
<tr><td class="indent">Number of Volunteers</td><td>{fmt(v_count)}</td><td>{fmt(v_count_ytd)}</td></tr>
<tr><td class="indent">Number of Volunteer Hours</td><td>{fmt(v_hours)}</td><td>{fmt(v_hours_ytd)}</td></tr>
</table>
"""

render_table(individuals_table)
st.markdown('<div class="mes-spacer"></div>', unsafe_allow_html=True)
render_table(households_table)
st.markdown('<div class="mes-spacer"></div>', unsafe_allow_html=True)
render_table(pounds_table)
st.markdown('<div class="mes-spacer"></div>', unsafe_allow_html=True)
render_table(volunteers_table)
