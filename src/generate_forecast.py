"""
Backend job: fits the ARIMAX forecast models and walk-forward backtest, then
writes the results to data/processed/forecast_cache.json for the Streamlit
app to read as static data. This is deliberately NOT run on every app page
load (auto_arima + a 16-fold walk-forward backtest per metric is too slow
for a free-tier web request) — run it as part of the data refresh pipeline
(see main_refresh.sh) instead.
"""
import json
import warnings
from datetime import datetime, timezone
from pathlib import Path

warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
from pandas.tseries.holiday import USFederalHolidayCalendar
from pmdarima import auto_arima
from statsmodels.tsa.statespace.sarimax import SARIMAX

DATA_DIR = Path(__file__).parent.parent / "data"
OUTPUT_PATH = DATA_DIR / "processed" / "forecast_cache.json"

EXOG_COLS = ["is_tuesday", "month_sin", "month_cos", "near_holiday"]
HOLIDAY_WINDOW_DAYS = 7
BACKTEST_HOLDOUT = 16  # last N service days (~8 weeks) held out for walk-forward validation
FORECAST_EVENTS = 2  # next N distribution events (each = a Monday + Tuesday) to forecast
HISTORY_TAIL = 16  # service days of recent actuals to include for charting

METRICS = {
    "total_hh_visits": "Household Visits",
}


def load_service_days() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "raw" / "total_report_daily.csv")
    df["date"] = pd.to_datetime(df["date"])
    df = df[(df["date"] >= "2024-01-01") & (df["day_of_week"].isin(["Monday", "Tuesday"]))]
    return df.sort_values("date").reset_index(drop=True)


def build_exog(dates: pd.DatetimeIndex, holidays: pd.DatetimeIndex) -> pd.DataFrame:
    dow = dates.day_name()
    month_num = dates.month
    days_to_holiday = np.array([np.min(np.abs((holidays - d).days)) for d in dates])
    return pd.DataFrame(
        {
            "is_tuesday": (dow == "Tuesday").astype(int),
            "month_sin": np.sin(2 * np.pi * month_num / 12),
            "month_cos": np.cos(2 * np.pi * month_num / 12),
            "near_holiday": (days_to_holiday <= HOLIDAY_WINDOW_DAYS).astype(int),
        },
        index=dates,
    )


def infer_distribution_cadence(service_days: pd.DataFrame) -> int:
    """
    Distribution isn't weekly — it's historically ~every two weeks with
    occasional exceptions (a skipped or doubled-up week). Use the most
    common historical gap between distribution Mondays (mode, falling back
    to the median) rather than assuming a fixed weekly cadence.
    """
    mondays = pd.DatetimeIndex(sorted(service_days.loc[service_days["day_of_week"] == "Monday", "date"].unique()))
    if len(mondays) < 2:
        return 14
    gaps = pd.Series((mondays[1:] - mondays[:-1]).days)
    mode = gaps.mode()
    return int(mode.iloc[0]) if len(mode) else int(gaps.median())


def next_distribution_events(service_days: pd.DataFrame, typical_gap_days: int, n_events: int) -> list:
    """Returns a list of (monday, tuesday) Timestamp pairs for the next N projected distribution events."""
    last_monday = service_days.loc[service_days["day_of_week"] == "Monday", "date"].max()
    events = []
    anchor = last_monday
    for _ in range(n_events):
        anchor = anchor + pd.Timedelta(days=typical_gap_days)
        events.append((anchor, anchor + pd.Timedelta(days=1)))
    return events


def walk_forward_backtest(dates: np.ndarray, y: np.ndarray, exog: np.ndarray, order: tuple, holdout: int) -> pd.DataFrame:
    n = len(y)
    holdout = max(min(holdout, n - 20), 0)
    records = []
    for i in range(holdout):
        cut = n - holdout + i
        y_train, x_train = y[:cut], exog[:cut]
        x_next = exog[cut:cut + 1]
        try:
            fit = SARIMAX(
                y_train, exog=x_train, order=order,
                enforce_stationarity=False, enforce_invertibility=False,
            ).fit(disp=False, maxiter=200)
            pred = float(fit.forecast(steps=1, exog=x_next)[0])
        except Exception:
            pred = float(y_train[-1])

        is_tue = x_next[0][0] == 1
        same_dow = y_train[x_train[:, 0] == (1 if is_tue else 0)]
        naive_pred = float(same_dow[-1]) if len(same_dow) else float(y_train[-1])

        records.append({"date": dates[cut], "actual": y[cut], "arimax": pred, "naive_same_weekday": naive_pred})
    return pd.DataFrame(records)


def accuracy(df: pd.DataFrame, col: str) -> tuple:
    err = df["actual"] - df[col]
    mae = err.abs().mean()
    rmse = (err ** 2).mean() ** 0.5
    mape = (err.abs() / df["actual"].replace(0, np.nan)).mean() * 100
    return mae, rmse, mape


def main() -> None:
    service_days = load_service_days()
    holidays = USFederalHolidayCalendar().holidays(start="2023-01-01", end="2030-12-31")
    exog_full = build_exog(pd.DatetimeIndex(service_days["date"]), holidays)[EXOG_COLS]

    typical_gap_days = infer_distribution_cadence(service_days)
    events = next_distribution_events(service_days, typical_gap_days, FORECAST_EVENTS)
    future_dates = pd.DatetimeIndex([d for pair in events for d in pair])
    future_exog = build_exog(future_dates, holidays)[EXOG_COLS]

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "typical_gap_days": typical_gap_days,
        "next_distribution_event": {
            "monday": events[0][0].strftime("%Y-%m-%d"),
            "tuesday": events[0][1].strftime("%Y-%m-%d"),
            "metrics": {},
        },
        "metrics": {},
    }

    for colname, label in METRICS.items():
        y = service_days[colname].astype(float).values
        exog_values = exog_full.values

        model = auto_arima(
            y, X=exog_values, seasonal=False,
            max_p=4, max_q=4, max_d=2,
            stepwise=True, information_criterion="aicc",
            suppress_warnings=True, error_action="ignore",
        )
        order = model.order

        fc, ci = model.predict(n_periods=len(future_dates), X=future_exog.values, return_conf_int=True)
        forecast_rows = [
            {
                "date": d.strftime("%Y-%m-%d"),
                "day": d.day_name(),
                "forecast": int(round(f)),
                "low_95": int(round(max(lo, 0))),
                "high_95": int(round(hi)),
            }
            for d, f, lo, hi in zip(future_dates, fc, ci[:, 0], ci[:, 1])
        ]

        # first two forecast rows correspond to the very next distribution event (Monday, Tuesday)
        output["next_distribution_event"]["metrics"][colname] = {
            "label": label,
            "monday": forecast_rows[0],
            "tuesday": forecast_rows[1],
        }

        backtest_df = walk_forward_backtest(service_days["date"].values, y, exog_values, order, BACKTEST_HOLDOUT)
        backtest_rows = []
        if len(backtest_df):
            for m_label, col in [(f"ARIMAX{order}", "arimax"), ("Naive (same weekday last time)", "naive_same_weekday")]:
                mae, rmse, mape = accuracy(backtest_df, col)
                backtest_rows.append({"method": m_label, "mae": round(mae, 1), "rmse": round(rmse, 1), "mape": round(mape, 1)})

        history = service_days[["date", colname]].tail(HISTORY_TAIL)
        history_rows = [
            {"date": d.strftime("%Y-%m-%d"), "value": float(v)}
            for d, v in zip(history["date"], history[colname])
        ]

        output["metrics"][colname] = {
            "label": label,
            "order": list(order),
            "forecast": forecast_rows,
            "backtest": backtest_rows,
            "backtest_n": len(backtest_df),
            "history": history_rows,
        }

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Wrote forecast cache to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
