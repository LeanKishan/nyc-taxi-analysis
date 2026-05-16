"""Streamlit dashboard for the NYC Yellow Taxi 2025 analysis.

Five interactive tabs:
1. Overview         — KPIs + daily volume/revenue trend
2. Time patterns    — hour-of-day and day-of-week demand heatmap
3. Geography        — top pickup zones
4. Payments & Tips  — breakdown by payment type and tip behaviour
5. ML Predictor     — interactive trip-duration prediction with the trained model

Run with:
    streamlit run dashboard/app.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import duckdb
import lightgbm as lgb
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.bootstrap import ensure_data_ready  # noqa: E402

DATASET = ROOT / "data" / "processed" / "yellow_2025_clean.parquet"
MODEL_PATH = ROOT / "models" / "duration_model.txt"
METRICS_PATH = ROOT / "models" / "duration_metrics.json"

st.set_page_config(
    page_title="NYC Yellow Taxi 2025 — Analytics",
    page_icon="🚕",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---- Data layer (cached) ----------------------------------------------------

@st.cache_resource(show_spinner="First launch — downloading + cleaning ~60 MB of taxi data (≈2 minutes)…")
def get_connection() -> duckdb.DuckDBPyConnection:
    ensure_data_ready()
    con = duckdb.connect(":memory:")
    con.execute(f"CREATE VIEW trips AS SELECT * FROM read_parquet('{DATASET.as_posix()}')")
    return con


@st.cache_data
def load_overview() -> pd.Series:
    con = get_connection()
    return con.execute("""
        SELECT
            COUNT(*) AS total_trips,
            ROUND(SUM(total_amount), 0) AS total_revenue,
            ROUND(AVG(fare_amount), 2) AS avg_fare,
            ROUND(AVG(trip_distance), 2) AS avg_distance,
            ROUND(AVG(trip_duration_min), 1) AS avg_duration,
            ROUND(AVG(tip_pct), 1) AS avg_tip_pct
        FROM trips
    """).df().iloc[0]


@st.cache_data
def load_daily() -> pd.DataFrame:
    con = get_connection()
    return con.execute("""
        SELECT pickup_date AS date,
               COUNT(*) AS trips,
               ROUND(SUM(total_amount), 0) AS revenue
        FROM trips
        GROUP BY pickup_date
        ORDER BY pickup_date
    """).df()


@st.cache_data
def load_hour_dow_heatmap() -> pd.DataFrame:
    con = get_connection()
    return con.execute("""
        SELECT pickup_dayofweek AS dow,
               pickup_day_name  AS day_name,
               pickup_hour      AS hour,
               COUNT(*)         AS trips
        FROM trips
        GROUP BY pickup_dayofweek, pickup_day_name, pickup_hour
        ORDER BY pickup_dayofweek, pickup_hour
    """).df()


@st.cache_data
def load_top_zones(n: int = 20) -> pd.DataFrame:
    con = get_connection()
    return con.execute(f"""
        SELECT pu_location_id AS zone,
               COUNT(*) AS trips,
               ROUND(AVG(fare_amount), 2) AS avg_fare,
               ROUND(SUM(total_amount), 0) AS revenue
        FROM trips
        GROUP BY pu_location_id
        ORDER BY trips DESC
        LIMIT {n}
    """).df()


@st.cache_data
def load_payment_breakdown() -> pd.DataFrame:
    con = get_connection()
    return con.execute("""
        SELECT payment_type_name AS payment_type,
               COUNT(*) AS trips,
               ROUND(AVG(fare_amount), 2) AS avg_fare,
               ROUND(AVG(tip_amount), 2) AS avg_tip,
               ROUND(AVG(tip_pct), 2) AS avg_tip_pct
        FROM trips
        GROUP BY payment_type_name
        ORDER BY trips DESC
    """).df()


@st.cache_data
def load_tip_distribution() -> pd.DataFrame:
    con = get_connection()
    return con.execute("""
        SELECT tip_pct
        FROM trips
        WHERE payment_type = 1
          AND tip_pct BETWEEN 0 AND 60
        USING SAMPLE 50000 ROWS
    """).df()


@st.cache_resource
def load_model() -> tuple[lgb.Booster | None, dict | None]:
    if not MODEL_PATH.exists():
        return None, None
    model = lgb.Booster(model_file=str(MODEL_PATH))
    metrics = json.loads(METRICS_PATH.read_text()) if METRICS_PATH.exists() else {}
    return model, metrics


# ---- UI ---------------------------------------------------------------------

st.title("NYC Yellow Taxi 2025 — Analytics Dashboard")
st.caption(
    "End-to-end exploration of NYC TLC Yellow Taxi trip records. "
    "Built with Python, DuckDB, LightGBM, and Streamlit."
)

ov = load_overview()
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Trips analysed", f"{int(ov['total_trips']):,}")
k2.metric("Total revenue", f"${ov['total_revenue']:,.0f}")
k3.metric("Avg fare", f"${ov['avg_fare']}")
k4.metric("Avg trip", f"{ov['avg_distance']} mi  ·  {ov['avg_duration']} min")
k5.metric("Avg tip (card)", f"{ov['avg_tip_pct']}%")

tab_overview, tab_time, tab_geo, tab_pay, tab_ml = st.tabs(
    ["📈 Overview", "🕐 Time patterns", "📍 Geography", "💳 Payments & tips", "🤖 ML predictor"]
)

with tab_overview:
    st.subheader("Daily trip volume and revenue")
    daily = load_daily()
    fig = px.line(
        daily, x="date", y=["trips", "revenue"],
        labels={"value": "Value", "date": "Date", "variable": "Metric"},
    )
    fig.update_layout(height=420, hovermode="x unified")
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(daily.tail(14), use_container_width=True, hide_index=True)

with tab_time:
    st.subheader("Demand heatmap — hour of day × day of week")
    hm = load_hour_dow_heatmap()
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    pivot = hm.pivot(index="day_name", columns="hour", values="trips").reindex(day_order)
    fig = px.imshow(
        pivot, aspect="auto", color_continuous_scale="YlOrRd",
        labels={"x": "Hour of day", "y": "Day of week", "color": "Trips"},
    )
    fig.update_layout(height=420)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Average trips by hour")
    hourly = hm.groupby("hour", as_index=False)["trips"].sum()
    fig2 = px.bar(hourly, x="hour", y="trips", labels={"hour": "Hour", "trips": "Total trips"})
    fig2.update_layout(height=320)
    st.plotly_chart(fig2, use_container_width=True)

with tab_geo:
    st.subheader("Top pickup zones")
    n_zones = st.slider("Number of zones to show", 5, 50, 20)
    zones = load_top_zones(n_zones)
    fig = px.bar(
        zones.sort_values("trips"), y="zone", x="trips", orientation="h",
        hover_data=["avg_fare", "revenue"],
        labels={"zone": "Pickup zone ID", "trips": "Total trips"},
    )
    fig.update_layout(height=600)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "Zone IDs map to NYC TLC Taxi Zones — see "
        "https://www1.nyc.gov/site/tlc/about/taxi-vehicle-information.page for the lookup."
    )

with tab_pay:
    st.subheader("Payment-type breakdown")
    pay = load_payment_breakdown()
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(pay, names="payment_type", values="trips", hole=0.45)
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(pay, use_container_width=True, hide_index=True)

    st.subheader("Tip-percentage distribution (credit card payments)")
    tips = load_tip_distribution()
    fig = px.histogram(tips, x="tip_pct", nbins=50, labels={"tip_pct": "Tip %"})
    fig.add_vline(x=tips["tip_pct"].median(), line_dash="dash", line_color="red",
                  annotation_text=f"Median: {tips['tip_pct'].median():.1f}%")
    fig.update_layout(height=380)
    st.plotly_chart(fig, use_container_width=True)

with tab_ml:
    st.subheader("Predict trip duration")
    model, metrics = load_model()
    if model is None:
        st.warning("Model not found. Run `python -m src.train_model` to train it first.")
    else:
        st.caption(
            f"Trained on {metrics.get('n_train', 0):,} trips · "
            f"Validation MAE: {metrics.get('mae_min', 0):.2f} min · "
            f"R²: {metrics.get('r2', 0):.3f}"
        )
        col1, col2, col3 = st.columns(3)
        with col1:
            trip_distance = st.number_input("Trip distance (miles)", 0.1, 50.0, 2.5, 0.1)
            passenger_count = st.number_input("Passengers", 1, 6, 1)
            pu_zone = st.number_input("Pickup zone ID", 1, 265, 161)
        with col2:
            pickup_hour = st.slider("Pickup hour", 0, 23, 17)
            pickup_dow = st.selectbox(
                "Day of week", list(range(7)),
                format_func=lambda i: ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"][i],
                index=2,
            )
            do_zone = st.number_input("Dropoff zone ID", 1, 265, 230)
        with col3:
            pickup_month = st.slider("Month", 1, 12, 3)
            ratecode = st.selectbox(
                "Rate code", [1, 2, 3, 4, 5, 6],
                format_func=lambda i: {1:"Standard",2:"JFK",3:"Newark",4:"Nassau/Westchester",5:"Negotiated",6:"Group"}[i],
            )
            is_airport = ratecode in (2, 3)

        is_weekend = pickup_dow in (5, 6)
        is_rush = (pickup_hour in range(7, 10) or pickup_hour in range(16, 20)) and not is_weekend

        if st.button("Predict duration", type="primary"):
            X = pd.DataFrame([{
                "trip_distance": trip_distance,
                "pickup_hour": pickup_hour,
                "pickup_dayofweek": pickup_dow,
                "pickup_month": pickup_month,
                "passenger_count": passenger_count,
                "pu_location_id": pu_zone,
                "do_location_id": do_zone,
                "ratecode_id": ratecode,
                "is_weekend": int(is_weekend),
                "is_rush_hour": int(is_rush),
                "is_airport_trip": int(is_airport),
            }])
            for c in ("pu_location_id", "do_location_id", "ratecode_id"):
                X[c] = X[c].astype("category")
            pred = float(model.predict(X)[0])
            avg_speed = trip_distance / (pred / 60) if pred > 0 else 0
            st.success(
                f"Estimated trip duration: **{pred:.1f} minutes**  ·  "
                f"average speed ≈ **{avg_speed:.1f} mph**"
            )
            if is_rush:
                st.info("Rush hour — expect more variance than usual.")

st.divider()
st.caption("Data: NYC Taxi & Limousine Commission · Built by Kishan · 2025")
