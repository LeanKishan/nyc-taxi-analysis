"""First-launch bootstrap for the deployed Streamlit app.

On Streamlit Community Cloud the container starts with only the contents of
the GitHub repo — the 290 MB cleaned Parquet isn't there. This module checks
whether the processed dataset exists and, if not, downloads a single month
of raw TLC data and runs the cleaning pipeline.

It's safe to call on every Streamlit run: when the data already exists the
function returns in microseconds.
"""
from __future__ import annotations

from pathlib import Path

from .data_cleaning import add_features, clean, save_processed
from .download_data import download_months

ROOT = Path(__file__).resolve().parent.parent
PROCESSED = ROOT / "data" / "processed" / "yellow_2025_clean.parquet"

# Single month keeps the cold start manageable on Streamlit Cloud (~60 MB
# download + ~30s cleaning). Locally you can still pull more months with
# `python -m src.download_data --year 2025 --months 1 2 3 ...`.
BOOTSTRAP_YEAR = 2025
BOOTSTRAP_MONTHS = [1]


def ensure_data_ready() -> Path:
    """Make sure the processed dataset exists. Build it if missing."""
    if PROCESSED.exists() and PROCESSED.stat().st_size > 0:
        return PROCESSED

    print(f"[bootstrap] {PROCESSED.name} missing — downloading and cleaning a single month")
    download_months("yellow", BOOTSTRAP_YEAR, BOOTSTRAP_MONTHS)

    import pandas as pd
    raw_files = sorted(
        (ROOT / "data" / "raw").glob(f"yellow_tripdata_{BOOTSTRAP_YEAR}-*.parquet")
    )
    if not raw_files:
        raise RuntimeError("Bootstrap failed — no raw files downloaded.")
    df = pd.concat([pd.read_parquet(f) for f in raw_files], ignore_index=True)
    df = clean(df)
    df = add_features(df)
    save_processed(df, name=PROCESSED.name)
    return PROCESSED


if __name__ == "__main__":
    path = ensure_data_ready()
    print(f"Dataset ready: {path}")
