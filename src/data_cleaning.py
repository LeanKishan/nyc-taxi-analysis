"""Cleaning and feature engineering for NYC TLC Yellow Taxi data.

The raw Parquet files are messy — they contain trips with negative fares,
unrealistic durations, trips outside NYC, etc. This module normalizes the
schema, drops obviously bad rows, and engineers features used by both the
analysis notebooks and the ML model.

Reference: TLC Yellow Taxi data dictionary
https://www.nyc.gov/assets/tlc/downloads/pdf/data_dictionary_trip_records_yellow.pdf
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

RAW_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"
PROCESSED_DIR = Path(__file__).resolve().parent.parent / "data" / "processed"

# Sane bounds for filtering. Values outside these are almost certainly bad rows.
MIN_TRIP_DISTANCE = 0.1      # miles
MAX_TRIP_DISTANCE = 100.0
MIN_FARE = 0.0
MAX_FARE = 500.0
MIN_DURATION_MIN = 1.0
MAX_DURATION_MIN = 180.0     # 3 hours
MIN_PASSENGERS = 1
MAX_PASSENGERS = 6

PAYMENT_TYPE_MAP = {
    1: "Credit card",
    2: "Cash",
    3: "No charge",
    4: "Dispute",
    5: "Unknown",
    6: "Voided trip",
}

RATECODE_MAP = {
    1: "Standard",
    2: "JFK",
    3: "Newark",
    4: "Nassau/Westchester",
    5: "Negotiated fare",
    6: "Group ride",
    99: "Unknown",
}


def load_raw_month(year: int, month: int) -> pd.DataFrame:
    """Load a single month of raw Yellow Taxi parquet data."""
    path = RAW_DIR / f"yellow_tripdata_{year}-{month:02d}.parquet"
    if not path.exists():
        raise FileNotFoundError(
            f"{path.name} not found. Run: python -m src.download_data --year {year} --months {month}"
        )
    return pd.read_parquet(path)


def load_raw_all(year: int) -> pd.DataFrame:
    """Load every downloaded monthly parquet for the given year."""
    files = sorted(RAW_DIR.glob(f"yellow_tripdata_{year}-*.parquet"))
    if not files:
        raise FileNotFoundError(f"No raw files found for year {year} in {RAW_DIR}")
    frames = [pd.read_parquet(f) for f in files]
    return pd.concat(frames, ignore_index=True)


def clean(df: pd.DataFrame) -> pd.DataFrame:
    """Apply sanity filters and normalize the schema."""
    df = df.copy()

    # Normalize column names — schema has shifted slightly across years.
    rename = {
        "tpep_pickup_datetime": "pickup_dt",
        "tpep_dropoff_datetime": "dropoff_dt",
        "PULocationID": "pu_location_id",
        "DOLocationID": "do_location_id",
        "RatecodeID": "ratecode_id",
        "VendorID": "vendor_id",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    df["pickup_dt"] = pd.to_datetime(df["pickup_dt"])
    df["dropoff_dt"] = pd.to_datetime(df["dropoff_dt"])
    df["trip_duration_min"] = (
        (df["dropoff_dt"] - df["pickup_dt"]).dt.total_seconds() / 60.0
    )

    initial = len(df)

    # TLC data sometimes contains trips with wildly wrong timestamps (year 2001,
    # 2008, etc.). Constrain to the actual filing year +/- a small buffer so the
    # time-series plots aren't polluted with outliers.
    year = df["pickup_dt"].dt.year.mode().iloc[0]
    year_lo = pd.Timestamp(year, 1, 1)
    year_hi = pd.Timestamp(year + 1, 1, 7)  # allow a few overflow days

    mask = (
        df["trip_distance"].between(MIN_TRIP_DISTANCE, MAX_TRIP_DISTANCE)
        & df["fare_amount"].between(MIN_FARE, MAX_FARE)
        & df["trip_duration_min"].between(MIN_DURATION_MIN, MAX_DURATION_MIN)
        & df["passenger_count"].between(MIN_PASSENGERS, MAX_PASSENGERS)
        & df["pu_location_id"].between(1, 265)
        & df["do_location_id"].between(1, 265)
        & df["total_amount"].between(0, 1000)
        & df["pickup_dt"].between(year_lo, year_hi)
    )
    df = df.loc[mask].copy()

    df["payment_type_name"] = (
        df["payment_type"].map(PAYMENT_TYPE_MAP).fillna("Unknown")
    )
    df["ratecode_name"] = df["ratecode_id"].map(RATECODE_MAP).fillna("Unknown")

    dropped = initial - len(df)
    print(f"  Cleaned: kept {len(df):,} of {initial:,} rows ({dropped:,} dropped, {dropped/initial:.1%})")
    return df.reset_index(drop=True)


def add_features(df: pd.DataFrame) -> pd.DataFrame:
    """Engineer features used by EDA and the ML model."""
    df = df.copy()

    df["pickup_hour"] = df["pickup_dt"].dt.hour
    df["pickup_dayofweek"] = df["pickup_dt"].dt.dayofweek
    df["pickup_day_name"] = df["pickup_dt"].dt.day_name()
    df["pickup_month"] = df["pickup_dt"].dt.month
    df["pickup_date"] = df["pickup_dt"].dt.date
    df["is_weekend"] = df["pickup_dayofweek"].isin([5, 6])

    bins = [-1, 5, 11, 16, 20, 24]
    labels = ["Late night", "Morning", "Afternoon", "Evening", "Night"]
    df["time_of_day"] = pd.cut(
        df["pickup_hour"], bins=bins, labels=labels, include_lowest=True
    )

    df["avg_speed_mph"] = df["trip_distance"] / (df["trip_duration_min"] / 60.0)
    df["avg_speed_mph"] = df["avg_speed_mph"].replace([np.inf, -np.inf], np.nan)

    df["fare_per_mile"] = df["fare_amount"] / df["trip_distance"]
    df["fare_per_mile"] = df["fare_per_mile"].replace([np.inf, -np.inf], np.nan)

    # Tip percentage — only meaningful for card payments (cash tips aren't recorded).
    df["tip_pct"] = np.where(
        (df["payment_type"] == 1) & (df["fare_amount"] > 0),
        df["tip_amount"] / df["fare_amount"] * 100,
        np.nan,
    )

    df["is_airport_trip"] = df["ratecode_id"].isin([2, 3])
    df["is_rush_hour"] = (
        (df["pickup_hour"].between(7, 9)) | (df["pickup_hour"].between(16, 19))
    ) & (~df["is_weekend"])

    return df


def save_processed(df: pd.DataFrame, name: str = "yellow_2025_clean.parquet") -> Path:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    out = PROCESSED_DIR / name
    df.to_parquet(out, index=False, compression="snappy")
    print(f"  Saved {len(df):,} rows to {out} ({out.stat().st_size / 1e6:.1f} MB)")
    return out


def main(year: int = 2025) -> None:
    print(f"Loading raw data for {year}...")
    df = load_raw_all(year)
    print(f"  Loaded {len(df):,} raw rows from {len(list(RAW_DIR.glob(f'yellow_tripdata_{year}-*.parquet')))} files")

    print("\nCleaning...")
    df = clean(df)

    print("\nAdding features...")
    df = add_features(df)

    print("\nSaving processed dataset...")
    save_processed(df)
    print("\nDone.")


if __name__ == "__main__":
    main()
