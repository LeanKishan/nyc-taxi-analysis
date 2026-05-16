"""Basic sanity tests for the cleaning + feature-engineering module.

These don't require the real dataset — they construct small synthetic frames
and verify that filters drop the right rows and engineered features take the
expected values. Run with: pytest tests/
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.data_cleaning import add_features, clean


def _toy_frame() -> pd.DataFrame:
    # 2025-03-03 is a Monday, 2025-03-08 is a Saturday.
    return pd.DataFrame({
        "tpep_pickup_datetime":  pd.to_datetime([
            "2025-03-03 08:30:00",  # valid weekday morning (rush hour)
            "2025-03-03 17:30:00",  # valid weekday evening (rush hour)
            "2025-03-08 02:00:00",  # valid weekend late night
            "2025-03-03 12:00:00",  # bad: negative fare
            "2025-03-03 12:00:00",  # bad: zero duration
            "2025-03-03 12:00:00",  # bad: invalid zone
        ]),
        "tpep_dropoff_datetime": pd.to_datetime([
            "2025-03-03 08:45:00",
            "2025-03-03 17:50:00",
            "2025-03-08 02:20:00",
            "2025-03-03 12:15:00",
            "2025-03-03 12:00:00",
            "2025-03-03 12:10:00",
        ]),
        "trip_distance":     [2.0, 3.0, 4.0, 1.5, 1.0, 2.0],
        "fare_amount":       [12.0, 18.0, 22.0, -5.0, 10.0, 12.0],
        "total_amount":      [15.0, 22.0, 28.0,  0.0, 12.0, 14.0],
        "passenger_count":   [1, 2, 1, 1, 1, 1],
        "PULocationID":      [161, 230, 132, 161, 230, 999],
        "DOLocationID":      [230, 161, 230, 161, 230, 1],
        "RatecodeID":        [1, 1, 2, 1, 1, 1],
        "payment_type":      [1, 1, 2, 1, 1, 1],
        "tip_amount":        [2.5, 3.0, 0.0, 0.0, 0.0, 0.0],
        "VendorID":          [1, 2, 1, 1, 2, 1],
    })


def test_clean_drops_bad_rows():
    out = clean(_toy_frame())
    assert len(out) == 3, "Expected exactly 3 valid rows after cleaning"
    assert (out["fare_amount"] > 0).all()
    assert out["pu_location_id"].between(1, 265).all()
    assert "trip_duration_min" in out.columns


def test_add_features_produces_expected_columns():
    cleaned = clean(_toy_frame())
    feat = add_features(cleaned)
    for col in [
        "pickup_hour", "pickup_dayofweek", "pickup_day_name",
        "is_weekend", "is_rush_hour", "is_airport_trip",
        "avg_speed_mph", "fare_per_mile", "tip_pct", "time_of_day",
    ]:
        assert col in feat.columns, f"Missing engineered column {col}"


def test_rush_hour_flag_logic():
    cleaned = clean(_toy_frame())
    feat = add_features(cleaned)
    rush_trip = feat[feat["pickup_hour"] == 17].iloc[0]
    assert rush_trip["is_rush_hour"]
    assert not rush_trip["is_weekend"]


def test_weekend_flag_logic():
    cleaned = clean(_toy_frame())
    feat = add_features(cleaned)
    weekend_trip = feat[feat["pickup_dt"].dt.dayofweek == 5].iloc[0]  # Saturday
    assert weekend_trip["is_weekend"]
    assert not weekend_trip["is_rush_hour"]


def test_airport_flag():
    cleaned = clean(_toy_frame())
    feat = add_features(cleaned)
    assert feat["is_airport_trip"].sum() == 1


if __name__ == "__main__":
    import subprocess
    sys.exit(subprocess.call([sys.executable, "-m", "pytest", str(Path(__file__).parent), "-v"]))
