# NYC Yellow Taxi 2025 — End-to-End Analytics

End-to-end data project on **8.5 million NYC Yellow Taxi trips** from January–March 2025. Built to demonstrate the full analytics workflow: ingestion → cleaning → SQL analytics → exploratory analysis → machine learning → an interactive dashboard.

---

## Headline numbers

| | |
|---|---|
| **Trips analysed** | 8,516,174 |
| **Period** | Jan 1 – Mar 31, 2025 |
| **Total revenue** | $236 million |
| **Raw data dropped during cleaning** | 23.9% (sanity filters) |
| **ML model** | LightGBM trip-duration regressor |
| **Validation MAE** | **2.77 minutes** |
| **Validation R²** | **0.864** |

---

## What's in the project

```
nyc-taxi-analysis/
├── src/
│   ├── download_data.py      # Streaming Parquet download from TLC CDN
│   ├── data_cleaning.py      # Filtering + feature engineering
│   ├── sql_analytics.py      # DuckDB analytical queries
│   └── train_model.py        # LightGBM training pipeline
├── notebooks/
│   └── 01_eda.ipynb          # Visual EDA (heatmaps, distributions, tips)
├── dashboard/
│   └── app.py                # Streamlit dashboard (5 interactive tabs)
├── tests/
│   └── test_cleaning.py      # Unit tests for cleaning logic
├── data/
│   ├── raw/                  # Downloaded Parquet (gitignored)
│   └── processed/            # Cleaned dataset (gitignored)
├── models/                   # Trained model artifacts + metrics
├── reports/figures/          # Saved plots (PNG)
├── requirements.txt
└── README.md
```

---

## Tech stack

- **Python 3.12**
- **Data engineering**: pandas, pyarrow, requests (with progress bar via tqdm)
- **SQL analytics**: DuckDB (querying Parquet files directly — no ETL step required)
- **Visualization**: matplotlib, seaborn, plotly
- **Machine learning**: scikit-learn, LightGBM (gradient-boosted trees)
- **Dashboard**: Streamlit
- **Testing**: pytest
- **Notebooks**: Jupyter

---

## Quick start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Download 3 months of 2025 yellow taxi data (~190 MB)
python -m src.download_data --year 2025 --months 1 2 3

# 3. Clean and feature-engineer the data (~290 MB output)
python -m src.data_cleaning

# 4. Run SQL analytics (prints 8 business reports)
python -m src.sql_analytics

# 5. Train the trip-duration model
python -m src.train_model

# 6. Run the unit tests
pytest tests/ -v

# 7. Launch the interactive dashboard
streamlit run dashboard/app.py
```

The pipeline is idempotent — already-downloaded files are skipped, and re-running cleaning regenerates the processed dataset from the raw files.

---

## Findings

The exploratory analysis surfaces several patterns that are visually obvious once plotted:

- **Demand peaks at 5–6 PM weekdays** and again late Friday/Saturday nights. The full hour × day-of-week heatmap is in [reports/figures/demand_heatmap.png](reports/figures/demand_heatmap.png).
- **Median trip is short** — about 1.5 miles and ~10 minutes — but the long right tail of airport runs (avg 17.9 mi, $71 fare, 47 min) drags the mean significantly.
- **Traffic is visible in the data**: average trip speed drops from ~14 mph late-night to **under 10 mph** during the 4–6 PM rush window.
- **86% of trips are paid by credit card**, with a median tip percentage of ~20% — clearly anchored on the in-cab UX defaults.
- **Airport trips tip less** in percentage terms (20.1% vs 26.2% for regular trips) but more in absolute dollars.
- A handful of pickup zone IDs (notably **132**, **138** — JFK and LaGuardia) drive disproportionately large revenue despite lower trip counts.

All eight SQL reports (overview, hour-of-day, day-of-week, top zones, payments, airport vs regular, rush-hour impact, top routes) are produced by `src/sql_analytics.py`.

---

## Machine learning — trip duration prediction

A LightGBM regressor predicts trip duration in minutes using only **features known at the moment the trip starts** (no leakage from dropoff fields):

- `trip_distance`, `passenger_count`
- `pickup_hour`, `pickup_dayofweek`, `pickup_month`
- `pu_location_id`, `do_location_id`, `ratecode_id` (treated as categoricals)
- Derived flags: `is_weekend`, `is_rush_hour`, `is_airport_trip`

**Trained on 400,000 trips**, validated on 100,000 held-out trips, early-stopping on validation RMSE.

| Metric | Value |
|---|---|
| Validation MAE | **2.77 min** |
| Validation RMSE | 4.49 min |
| Validation R² | **0.864** |

Top features by gain: `trip_distance` (dominant), `ratecode_id` (captures airport flat-rate trips), `pickup_hour` (traffic), then the location IDs. Diagnostic plots at [reports/figures/predicted_vs_actual.png](reports/figures/predicted_vs_actual.png) and [reports/figures/feature_importance.png](reports/figures/feature_importance.png).

The trained model is loaded by the Streamlit dashboard for live inference.

---

## Interactive dashboard

`streamlit run dashboard/app.py` launches a five-tab dashboard:

1. **Overview** — top-level KPIs and a daily volume/revenue trend chart
2. **Time patterns** — interactive hour × day-of-week demand heatmap
3. **Geography** — busiest pickup zones with revenue and avg fare
4. **Payments & tips** — payment-type breakdown and tip-percentage distribution
5. **ML predictor** — interactive form that calls the trained LightGBM model for live duration predictions

Data is cached with `@st.cache_data` and the DuckDB connection is cached with `@st.cache_resource`, so navigating between tabs is instant after the first load.

---

## Data quality work

The raw TLC files are not analysis-ready. The cleaning step (`src/data_cleaning.py`) drops rows that fail any of:

- Trip distance outside `[0.1, 100]` miles
- Fare outside `[0, 500]` dollars
- Duration outside `[1, 180]` minutes
- Passenger count outside `[1, 6]`
- Pickup or dropoff zone ID outside the valid `[1, 265]` range
- Pickup timestamp outside the actual filing year (TLC sometimes leaks trips with timestamps from 2008, 2014, etc.)

Roughly 24% of raw rows are dropped — a meaningful share, and a reminder that "8.5 million clean rows" is the right number to quote, not the raw 11.2 million.

---

## What I'd build next

- **Geo visualization**: join the TLC taxi-zone shapefile and render a choropleth of trips/revenue per zone with Folium.
- **Comparison across vehicle types**: download Green Taxi and FHV (Uber/Lyft) data and compare market share over the same window.
- **Model improvements**: stack the LightGBM model with a quantile regressor to get prediction intervals (P10–P90 ETA), and try a sequence model on the per-zone time series.
- **Deployment**: containerise the dashboard and host it on Streamlit Community Cloud or Fly.io so the project is one click away from a recruiter's browser.

---

## Data attribution

Trip records are published monthly by the [NYC Taxi & Limousine Commission](https://www.nyc.gov/site/tlc/about/tlc-trip-record-data.page) as public-domain Parquet files. This project uses the Yellow Taxi files for January, February, and March 2025.

---

*Built by Kishan, 2026.*
