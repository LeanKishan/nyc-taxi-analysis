"""SQL-based analytics on NYC TLC data using DuckDB.

DuckDB queries the cleaned Parquet file directly — no ETL step needed. This
module collects the business-question SQL used throughout the project and
exposes them as named functions returning pandas DataFrames.

Run as a script to print every report.
"""
from __future__ import annotations

from pathlib import Path

import duckdb
import pandas as pd

PROCESSED_DIR = Path(__file__).resolve().parent.parent / "data" / "processed"
DEFAULT_DATASET = PROCESSED_DIR / "yellow_2025_clean.parquet"


def get_connection(dataset: Path = DEFAULT_DATASET) -> duckdb.DuckDBPyConnection:
    """Create a DuckDB connection with the cleaned parquet registered as `trips`."""
    if not dataset.exists():
        raise FileNotFoundError(
            f"{dataset} not found. Run: python -m src.data_cleaning"
        )
    con = duckdb.connect(":memory:")
    con.execute(f"CREATE VIEW trips AS SELECT * FROM read_parquet('{dataset.as_posix()}')")
    return con


def overview(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """High-level totals: row count, date range, revenue, average trip stats."""
    return con.execute("""
        SELECT
            COUNT(*)                         AS total_trips,
            MIN(pickup_dt)                   AS earliest_pickup,
            MAX(pickup_dt)                   AS latest_pickup,
            ROUND(SUM(total_amount), 2)      AS total_revenue,
            ROUND(AVG(fare_amount), 2)       AS avg_fare,
            ROUND(AVG(trip_distance), 2)     AS avg_distance_mi,
            ROUND(AVG(trip_duration_min), 2) AS avg_duration_min,
            ROUND(AVG(tip_pct), 2)           AS avg_tip_pct
        FROM trips
    """).df()


def trips_by_hour(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Demand curve across the day."""
    return con.execute("""
        SELECT
            pickup_hour,
            COUNT(*)                         AS trips,
            ROUND(AVG(fare_amount), 2)       AS avg_fare,
            ROUND(AVG(trip_duration_min), 2) AS avg_duration,
            ROUND(AVG(avg_speed_mph), 2)     AS avg_speed_mph
        FROM trips
        GROUP BY pickup_hour
        ORDER BY pickup_hour
    """).df()


def trips_by_day_of_week(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Weekly pattern."""
    return con.execute("""
        SELECT
            pickup_day_name,
            COUNT(*)                         AS trips,
            ROUND(AVG(fare_amount), 2)       AS avg_fare,
            ROUND(AVG(tip_pct), 2)           AS avg_tip_pct
        FROM trips
        GROUP BY pickup_day_name, pickup_dayofweek
        ORDER BY pickup_dayofweek
    """).df()


def top_pickup_zones(con: duckdb.DuckDBPyConnection, n: int = 20) -> pd.DataFrame:
    """Busiest pickup zones by trip count."""
    return con.execute(f"""
        SELECT
            pu_location_id,
            COUNT(*)                         AS trips,
            ROUND(AVG(fare_amount), 2)       AS avg_fare,
            ROUND(SUM(total_amount), 2)      AS revenue
        FROM trips
        GROUP BY pu_location_id
        ORDER BY trips DESC
        LIMIT {n}
    """).df()


def payment_breakdown(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Share and economics of each payment type."""
    return con.execute("""
        SELECT
            payment_type_name,
            COUNT(*)                                                  AS trips,
            ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 2)        AS pct_of_trips,
            ROUND(AVG(fare_amount), 2)                                AS avg_fare,
            ROUND(AVG(tip_amount), 2)                                 AS avg_tip,
            ROUND(AVG(tip_pct), 2)                                    AS avg_tip_pct
        FROM trips
        GROUP BY payment_type_name
        ORDER BY trips DESC
    """).df()


def airport_vs_regular(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Compare airport runs vs ordinary city trips."""
    return con.execute("""
        SELECT
            CASE WHEN is_airport_trip THEN 'Airport' ELSE 'Regular' END AS trip_type,
            COUNT(*)                         AS trips,
            ROUND(AVG(fare_amount), 2)       AS avg_fare,
            ROUND(AVG(trip_distance), 2)     AS avg_distance_mi,
            ROUND(AVG(trip_duration_min), 2) AS avg_duration_min,
            ROUND(AVG(tip_pct), 2)           AS avg_tip_pct
        FROM trips
        GROUP BY is_airport_trip
        ORDER BY trips DESC
    """).df()


def rush_hour_impact(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """How rush hour affects speed and fares."""
    return con.execute("""
        SELECT
            CASE WHEN is_rush_hour THEN 'Rush hour' ELSE 'Off-peak' END AS period,
            COUNT(*)                          AS trips,
            ROUND(AVG(avg_speed_mph), 2)      AS avg_speed_mph,
            ROUND(AVG(trip_duration_min), 2)  AS avg_duration_min,
            ROUND(AVG(fare_per_mile), 2)      AS fare_per_mile
        FROM trips
        WHERE avg_speed_mph IS NOT NULL
        GROUP BY is_rush_hour
        ORDER BY trips DESC
    """).df()


def daily_revenue(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Daily trip volume and revenue — useful for time-series plots."""
    return con.execute("""
        SELECT
            pickup_date                      AS date,
            COUNT(*)                         AS trips,
            ROUND(SUM(total_amount), 2)      AS revenue,
            ROUND(AVG(fare_amount), 2)       AS avg_fare
        FROM trips
        GROUP BY pickup_date
        ORDER BY pickup_date
    """).df()


def top_routes(con: duckdb.DuckDBPyConnection, n: int = 15) -> pd.DataFrame:
    """Most popular origin→destination zone pairs."""
    return con.execute(f"""
        SELECT
            pu_location_id                   AS pickup_zone,
            do_location_id                   AS dropoff_zone,
            COUNT(*)                         AS trips,
            ROUND(AVG(fare_amount), 2)       AS avg_fare,
            ROUND(AVG(trip_distance), 2)     AS avg_distance_mi
        FROM trips
        GROUP BY pu_location_id, do_location_id
        ORDER BY trips DESC
        LIMIT {n}
    """).df()


REPORTS = {
    "Overview": overview,
    "Trips by hour of day": trips_by_hour,
    "Trips by day of week": trips_by_day_of_week,
    "Top 20 pickup zones": top_pickup_zones,
    "Payment-type breakdown": payment_breakdown,
    "Airport vs regular trips": airport_vs_regular,
    "Rush hour vs off-peak": rush_hour_impact,
    "Top 15 routes": top_routes,
}


def main() -> None:
    con = get_connection()
    for title, fn in REPORTS.items():
        print(f"\n=== {title} ===")
        print(fn(con).to_string(index=False))
    con.close()


if __name__ == "__main__":
    main()
