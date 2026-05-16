"""Download NYC TLC Yellow Taxi trip data.

The NYC Taxi & Limousine Commission publishes monthly trip records as Parquet
files on a CloudFront CDN. This script downloads the requested months into
`data/raw/`, with size reporting and resume-safe behavior (skip if exists).

Usage:
    python -m src.download_data --year 2025 --months 1 2 3
    python -m src.download_data --year 2025 --months 1 2 3 4 5 6 --taxi-type yellow
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import requests
from tqdm import tqdm

BASE_URL = "https://d37ci6vzurychx.cloudfront.net/trip-data"
DATA_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"


def build_url(taxi_type: str, year: int, month: int) -> str:
    return f"{BASE_URL}/{taxi_type}_tripdata_{year}-{month:02d}.parquet"


def download_file(url: str, dest: Path, chunk_size: int = 1 << 20) -> bool:
    """Download a file with a progress bar. Returns True if downloaded, False if skipped."""
    if dest.exists() and dest.stat().st_size > 0:
        print(f"  [skip] {dest.name} already exists ({dest.stat().st_size / 1e6:.1f} MB)")
        return False

    with requests.get(url, stream=True, timeout=60) as r:
        if r.status_code == 403:
            print(f"  [miss] {url} (not yet published)")
            return False
        r.raise_for_status()
        total = int(r.headers.get("Content-Length", 0))
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_suffix(dest.suffix + ".part")
        with open(tmp, "wb") as f, tqdm(
            total=total, unit="B", unit_scale=True, desc=dest.name, leave=False
        ) as bar:
            for chunk in r.iter_content(chunk_size=chunk_size):
                if chunk:
                    f.write(chunk)
                    bar.update(len(chunk))
        tmp.rename(dest)
    print(f"  [done] {dest.name} ({dest.stat().st_size / 1e6:.1f} MB)")
    return True


def download_months(taxi_type: str, year: int, months: list[int]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Downloading {taxi_type} taxi data for {year}, months {months}")
    print(f"Destination: {DATA_DIR}\n")

    downloaded = 0
    for month in months:
        url = build_url(taxi_type, year, month)
        dest = DATA_DIR / f"{taxi_type}_tripdata_{year}-{month:02d}.parquet"
        try:
            if download_file(url, dest):
                downloaded += 1
        except requests.HTTPError as e:
            print(f"  [error] {url}: {e}")

    print(f"\nDone. {downloaded} new file(s) downloaded.")
    total_size = sum(f.stat().st_size for f in DATA_DIR.glob("*.parquet"))
    print(f"Total raw data size: {total_size / 1e6:.1f} MB")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Download NYC TLC trip data")
    p.add_argument("--year", type=int, default=2025, help="Year (default 2025)")
    p.add_argument(
        "--months",
        type=int,
        nargs="+",
        default=[1, 2, 3],
        help="Months to download (default Jan-Mar)",
    )
    p.add_argument(
        "--taxi-type",
        choices=["yellow", "green", "fhv", "fhvhv"],
        default="yellow",
        help="Taxi type (default: yellow)",
    )
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    try:
        download_months(args.taxi_type, args.year, args.months)
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        sys.exit(130)
