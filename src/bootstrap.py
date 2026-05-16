"""First-launch bootstrap for the deployed Streamlit app.

On Streamlit Community Cloud the container starts with only the contents of
the GitHub repo — the 290 MB cleaned Parquet isn't there. This module checks
whether the processed dataset exists and, if not, downloads a single month
of raw TLC data and runs the cleaning pipeline.

Cleaning is done in DuckDB rather than pandas because the free Streamlit tier
caps memory at 1 GB — pandas would OOM loading 3.5M rows + engineered features.
DuckDB streams Parquet, applies SQL filters / feature expressions, and writes
the result without ever holding the full dataset in RAM.

It's safe to call on every Streamlit run: when the data already exists the
function returns in microseconds.
"""
from __future__ import annotations

from pathlib import Path

import duckdb

from .download_data import download_months

ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw"
PROCESSED = ROOT / "data" / "processed" / "yellow_2025_clean.parquet"

# Single month keeps the cold start manageable on Streamlit Cloud. The 500k
# sample is what survives cleaning + sampling and is what the dashboard sees.
BOOTSTRAP_YEAR = 2025
BOOTSTRAP_MONTHS = [1]
SAMPLE_ROWS = 500_000


def ensure_data_ready() -> Path:
    """Make sure the processed dataset exists. Build it via DuckDB if missing."""
    if PROCESSED.exists() and PROCESSED.stat().st_size > 0:
        return PROCESSED

    print(f"[bootstrap] {PROCESSED.name} missing — downloading + cleaning via DuckDB")
    download_months("yellow", BOOTSTRAP_YEAR, BOOTSTRAP_MONTHS)

    raw_files = sorted(RAW_DIR.glob(f"yellow_tripdata_{BOOTSTRAP_YEAR}-*.parquet"))
    if not raw_files:
        raise RuntimeError("Bootstrap failed — no raw files downloaded.")

    PROCESSED.parent.mkdir(parents=True, exist_ok=True)
    raw_glob = (RAW_DIR / f"yellow_tripdata_{BOOTSTRAP_YEAR}-*.parquet").as_posix()

    # ISODOW returns Mon=1..Sun=7. Subtracting 1 gives us pandas' Mon=0..Sun=6
    # convention so the rest of the dashboard (which sorts by pickup_dayofweek
    # and expects Saturday=5, Sunday=6 as weekend) works unchanged.
    con = duckdb.connect(":memory:")
    con.execute(f"""
        COPY (
            WITH cleaned AS (
                SELECT
                    tpep_pickup_datetime  AS pickup_dt,
                    tpep_dropoff_datetime AS dropoff_dt,
                    PULocationID          AS pu_location_id,
                    DOLocationID          AS do_location_id,
                    RatecodeID            AS ratecode_id,
                    VendorID              AS vendor_id,
                    trip_distance, fare_amount, total_amount,
                    passenger_count, payment_type, tip_amount,
                    DATE_DIFF('second', tpep_pickup_datetime, tpep_dropoff_datetime) / 60.0
                        AS trip_duration_min
                FROM read_parquet('{raw_glob}')
                WHERE trip_distance     BETWEEN 0.1 AND 100
                  AND fare_amount       BETWEEN 0   AND 500
                  AND passenger_count   BETWEEN 1   AND 6
                  AND PULocationID      BETWEEN 1   AND 265
                  AND DOLocationID      BETWEEN 1   AND 265
                  AND total_amount      BETWEEN 0   AND 1000
                  AND tpep_pickup_datetime >= TIMESTAMP '{BOOTSTRAP_YEAR}-01-01'
                  AND tpep_pickup_datetime <  TIMESTAMP '{BOOTSTRAP_YEAR + 1}-01-07'
            ),
            filtered AS (
                SELECT * FROM cleaned WHERE trip_duration_min BETWEEN 1 AND 180
            ),
            sampled AS (
                SELECT * FROM filtered USING SAMPLE {SAMPLE_ROWS} ROWS
            )
            SELECT
                pickup_dt, dropoff_dt, pu_location_id, do_location_id,
                ratecode_id, vendor_id, trip_distance, fare_amount,
                total_amount, passenger_count, payment_type, tip_amount,
                trip_duration_min,

                CASE payment_type
                    WHEN 1 THEN 'Credit card' WHEN 2 THEN 'Cash'
                    WHEN 3 THEN 'No charge'   WHEN 4 THEN 'Dispute'
                    WHEN 5 THEN 'Unknown'     WHEN 6 THEN 'Voided trip'
                    ELSE 'Unknown'
                END                                  AS payment_type_name,

                CASE ratecode_id
                    WHEN 1 THEN 'Standard' WHEN 2 THEN 'JFK'
                    WHEN 3 THEN 'Newark'   WHEN 4 THEN 'Nassau/Westchester'
                    WHEN 5 THEN 'Negotiated fare' WHEN 6 THEN 'Group ride'
                    WHEN 99 THEN 'Unknown'
                    ELSE 'Unknown'
                END                                  AS ratecode_name,

                CAST(EXTRACT(HOUR    FROM pickup_dt) AS INTEGER) AS pickup_hour,
                CAST(EXTRACT(ISODOW  FROM pickup_dt) - 1 AS INTEGER) AS pickup_dayofweek,
                STRFTIME(pickup_dt, '%A')                            AS pickup_day_name,
                CAST(EXTRACT(MONTH   FROM pickup_dt) AS INTEGER) AS pickup_month,
                CAST(pickup_dt AS DATE)                              AS pickup_date,

                (EXTRACT(ISODOW FROM pickup_dt) - 1) IN (5, 6)       AS is_weekend,

                CASE
                    WHEN EXTRACT(HOUR FROM pickup_dt) <= 5  THEN 'Late night'
                    WHEN EXTRACT(HOUR FROM pickup_dt) <= 11 THEN 'Morning'
                    WHEN EXTRACT(HOUR FROM pickup_dt) <= 16 THEN 'Afternoon'
                    WHEN EXTRACT(HOUR FROM pickup_dt) <= 20 THEN 'Evening'
                    ELSE 'Night'
                END                                                  AS time_of_day,

                CASE WHEN trip_duration_min > 0
                     THEN trip_distance / (trip_duration_min / 60.0)
                     ELSE NULL
                END                                                  AS avg_speed_mph,

                CASE WHEN trip_distance > 0
                     THEN fare_amount / trip_distance
                     ELSE NULL
                END                                                  AS fare_per_mile,

                CASE WHEN payment_type = 1 AND fare_amount > 0
                     THEN tip_amount / fare_amount * 100
                     ELSE NULL
                END                                                  AS tip_pct,

                ratecode_id IN (2, 3)                                AS is_airport_trip,

                ((EXTRACT(HOUR FROM pickup_dt) BETWEEN 7 AND 9)
                  OR (EXTRACT(HOUR FROM pickup_dt) BETWEEN 16 AND 19))
                  AND (EXTRACT(ISODOW FROM pickup_dt) - 1) NOT IN (5, 6)
                                                                     AS is_rush_hour

            FROM sampled
        )
        TO '{PROCESSED.as_posix()}' (FORMAT PARQUET, COMPRESSION SNAPPY)
    """)
    con.close()

    size_mb = PROCESSED.stat().st_size / 1e6
    print(f"[bootstrap] Wrote {size_mb:.1f} MB to {PROCESSED}")
    return PROCESSED


if __name__ == "__main__":
    path = ensure_data_ready()
    print(f"Dataset ready: {path}")
