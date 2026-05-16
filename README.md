# NYC Yellow Taxi 2025 Analysis

Data analysis project on NYC Yellow Taxi trip records from January to March 2025. I built this to get hands-on practice with a real end-to-end data workflow: pulling raw data, cleaning it, doing SQL and visual analysis, training a model on it, and putting the whole thing behind a dashboard.

Around 11 million raw trip records came down from the NYC TLC. After cleaning I'm working with about 8.5 million, which adds up to roughly $236 million in fare revenue.

## What's in here

```
nyc-taxi-analysis/
├── src/
│   ├── download_data.py     # downloads the monthly parquet files from TLC
│   ├── data_cleaning.py     # filters out bad rows, adds features
│   ├── sql_analytics.py     # DuckDB queries
│   ├── train_model.py       # trains the LightGBM model
│   └── bootstrap.py         # rebuilds the dataset on first launch (for deploy)
├── dashboard/
│   └── app.py               # Streamlit dashboard
├── notebooks/
│   └── 01_eda.ipynb         # exploratory analysis with charts
├── tests/
│   └── test_cleaning.py     # pytest tests for the cleaning logic
├── data/                    # gitignored (too big), gets created locally
├── models/                  # trained LightGBM model + metrics
├── reports/
│   └── NYC_Taxi_2025_Analytics_Report.docx
├── requirements.txt
└── README.md
```

## How to run it

```bash
pip install -r requirements.txt

# download Jan/Feb/Mar 2025 (about 190 MB)
python -m src.download_data --year 2025 --months 1 2 3

# clean and add features (writes a ~290 MB parquet)
python -m src.data_cleaning

# print 8 SQL reports to the terminal
python -m src.sql_analytics

# train the trip-duration model
python -m src.train_model

# run the tests
pytest tests/ -v

# launch the dashboard in your browser
streamlit run dashboard/app.py
```

You don't have to re-run everything every time. If the raw files are already on disk, the download step skips them.

## Tech stack

Python 3.12. Pandas + pyarrow for the in-memory stuff, DuckDB for the SQL queries (it can read parquet files directly which is really nice). LightGBM for the model. Streamlit for the dashboard. matplotlib + seaborn + plotly for charts. pytest for the small test suite.

I picked Yellow Taxi specifically (not Green or Uber/Lyft) because the schema has been the most stable over the years, so the code I wrote here can pretty easily be pointed at older or newer files without changes.

## Findings from the analysis

- Demand peaks at 5-6 PM on weekdays, with a second peak late Friday/Saturday nights. The heatmap in [reports/figures/demand_heatmap.png](reports/figures/demand_heatmap.png) shows this clearly.
- The median trip is short, only about 1.5 miles and 10 minutes. But there's a long tail of airport runs that average 17.9 miles, $71 fare, and 47 minutes.
- Traffic shows up in the data. Average trip speed drops from around 14 mph late at night to under 10 mph during the afternoon rush. That's a real congestion footprint.
- 86% of trips are paid by credit card. Median tip is 20%, very tightly clustered, almost certainly because of the preset buttons in the cab.
- Airport trips tip a lower percentage (20% vs 26% for regular trips) but bigger absolute dollars because the fares are higher.
- Two zone IDs - 132 (JFK) and 138 (LaGuardia) - drive way more revenue than their trip count would suggest.

## The ML model

I trained a LightGBM regressor to predict trip duration in minutes. The constraint I gave myself was that the model could only use features that would be known at the moment a trip starts. So no leakage from the dropoff time or final fare.

Features used:
- `trip_distance`, `passenger_count`
- `pickup_hour`, `pickup_dayofweek`, `pickup_month`
- `pu_location_id`, `do_location_id`, `ratecode_id` (these are categorical)
- Engineered flags: `is_weekend`, `is_rush_hour`, `is_airport_trip`

Trained on 400K sampled trips, validated on 100K held out. Early stopping on validation RMSE.

| | |
|---|---|
| Validation MAE | 2.77 min |
| Validation RMSE | 4.49 min |
| Validation R² | 0.864 |

So the typical prediction is within about 3 minutes of the truth, which I think is decent given the model only sees information available at trip start.

Distance was by far the most important feature (which isn't surprising). After that, ratecode_id was the most useful because it cleanly identifies the airport flat-rate trips, which behave differently from regular metered trips. Then pickup hour, which captures the traffic effect. The plots are at [reports/figures/predicted_vs_actual.png](reports/figures/predicted_vs_actual.png) and [reports/figures/feature_importance.png](reports/figures/feature_importance.png).

The Streamlit dashboard loads this trained model and does live inference when you submit a form.

## The dashboard

`streamlit run dashboard/app.py` opens a dashboard with 5 tabs:

1. Overview - high level KPIs and a daily trips/revenue line chart
2. Time patterns - the hour-of-day x day-of-week heatmap
3. Geography - bar chart of busiest pickup zones (slider to change how many)
4. Payments & tips - payment breakdown plus the tip percentage histogram
5. ML predictor - form where you enter trip details and it predicts how long the trip will take

I cached the DuckDB connection and the query results so tab switching is fast.

## Data cleaning

The raw TLC files have a fair amount of junk in them. The cleaning step drops anything that fails these checks:

- Trip distance outside 0.1 to 100 miles
- Fare outside $0 to $500
- Duration outside 1 to 180 minutes
- Passenger count outside 1 to 6
- Pickup or dropoff zone ID outside the valid 1 to 265 range
- Pickup timestamp not in the actual filing year

That last one was annoying to figure out. The TLC files occasionally leak in trips with timestamps from 2008, 2014, other random years, and without filtering them out my time series charts had weird outliers.

About 24% of the raw rows get dropped, which sounds like a lot but is consistent with what other people have found working with this dataset.

## Stuff I'd add if I keep working on this

- A choropleth map of trips/revenue per zone using the TLC shapefile. Right now zones are just integer IDs in the dashboard, which isn't great.
- Compare against the Uber/Lyft (HVFHV) data for the same period. Yellow Taxi is actually a minority of the for-hire vehicle market in NYC at this point.
- Replace the point estimate from the model with prediction intervals (P10-P90) using quantile regression.
- Pull in weather data and add it as a feature. Rainy and snowy days probably mess with the predictions.

## Data source

All trip data is published monthly by the [NYC Taxi & Limousine Commission](https://www.nyc.gov/site/tlc/about/tlc-trip-record-data.page) as public Parquet files. I'm using Yellow Taxi for January, February, and March 2025.
