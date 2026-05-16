"""Train a LightGBM model to predict NYC taxi trip duration.

Predicting trip duration is a classic ride-sharing problem: dispatchers need
realistic ETAs, and surge pricing depends on it. The model uses pickup-time
features, location IDs, and trip distance — everything known at the moment
the trip starts (no leakage from dropoff fields).

Outputs:
- models/duration_model.txt       LightGBM model artifact
- models/duration_metrics.json    Train/val metrics
- models/feature_importance.csv   Sorted feature importances
- reports/figures/predicted_vs_actual.png
- reports/figures/feature_importance.png
"""
from __future__ import annotations

import json
from pathlib import Path

import joblib
import lightgbm as lgb
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split

PROCESSED_DIR = Path(__file__).resolve().parent.parent / "data" / "processed"
MODELS_DIR = Path(__file__).resolve().parent.parent / "models"
FIGURES_DIR = Path(__file__).resolve().parent.parent / "reports" / "figures"

FEATURES = [
    "trip_distance",
    "pickup_hour",
    "pickup_dayofweek",
    "pickup_month",
    "passenger_count",
    "pu_location_id",
    "do_location_id",
    "ratecode_id",
    "is_weekend",
    "is_rush_hour",
    "is_airport_trip",
]
TARGET = "trip_duration_min"
CATEGORICAL = ["pu_location_id", "do_location_id", "ratecode_id"]


def load_dataset() -> pd.DataFrame:
    path = PROCESSED_DIR / "yellow_2025_clean.parquet"
    if not path.exists():
        raise FileNotFoundError(
            f"{path} not found. Run: python -m src.data_cleaning"
        )
    return pd.read_parquet(path)


def prepare(df: pd.DataFrame, sample_size: int | None = 500_000) -> pd.DataFrame:
    """Select feature columns; optionally subsample to keep training quick."""
    df = df[FEATURES + [TARGET]].dropna()
    df = df[df[TARGET].between(1, 120)]
    if sample_size and len(df) > sample_size:
        df = df.sample(n=sample_size, random_state=42)
        print(f"  Subsampled to {len(df):,} rows for tractable training")
    for col in CATEGORICAL:
        df[col] = df[col].astype("category")
    df["is_weekend"] = df["is_weekend"].astype(int)
    df["is_rush_hour"] = df["is_rush_hour"].astype(int)
    df["is_airport_trip"] = df["is_airport_trip"].astype(int)
    return df


def train(df: pd.DataFrame) -> tuple[lgb.Booster, dict, pd.DataFrame, pd.Series, np.ndarray]:
    X = df[FEATURES]
    y = df[TARGET]

    X_train, X_val, y_train, y_val = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    train_set = lgb.Dataset(X_train, y_train, categorical_feature=CATEGORICAL)
    val_set = lgb.Dataset(X_val, y_val, categorical_feature=CATEGORICAL, reference=train_set)

    params = {
        "objective": "regression",
        "metric": "rmse",
        "learning_rate": 0.05,
        "num_leaves": 63,
        "feature_fraction": 0.9,
        "bagging_fraction": 0.8,
        "bagging_freq": 5,
        "verbose": -1,
    }

    print("  Training LightGBM (early stopping on validation RMSE)...")
    model = lgb.train(
        params,
        train_set,
        num_boost_round=500,
        valid_sets=[train_set, val_set],
        valid_names=["train", "val"],
        callbacks=[lgb.early_stopping(stopping_rounds=30), lgb.log_evaluation(period=50)],
    )

    preds_val = model.predict(X_val, num_iteration=model.best_iteration)
    metrics = {
        "rmse_min": float(np.sqrt(mean_squared_error(y_val, preds_val))),
        "mae_min": float(mean_absolute_error(y_val, preds_val)),
        "r2": float(r2_score(y_val, preds_val)),
        "n_train": int(len(X_train)),
        "n_val": int(len(X_val)),
        "best_iteration": int(model.best_iteration),
    }
    print(f"\n  Validation metrics: {metrics}")
    return model, metrics, X_val, y_val, preds_val


def save_artifacts(
    model: lgb.Booster,
    metrics: dict,
    X_val: pd.DataFrame,
    y_val: pd.Series,
    preds: np.ndarray,
) -> None:
    MODELS_DIR.mkdir(parents=True, exist_ok=True)
    FIGURES_DIR.mkdir(parents=True, exist_ok=True)

    model.save_model(str(MODELS_DIR / "duration_model.txt"))
    joblib.dump(
        {"features": FEATURES, "categorical": CATEGORICAL, "target": TARGET},
        MODELS_DIR / "model_schema.pkl",
    )
    (MODELS_DIR / "duration_metrics.json").write_text(json.dumps(metrics, indent=2))

    importance = pd.DataFrame({
        "feature": model.feature_name(),
        "importance": model.feature_importance(importance_type="gain"),
    }).sort_values("importance", ascending=False)
    importance.to_csv(MODELS_DIR / "feature_importance.csv", index=False)
    print(f"  Top 5 features by gain:")
    print(importance.head().to_string(index=False))

    fig, ax = plt.subplots(figsize=(7, 7))
    sample = np.random.choice(len(y_val), size=min(5000, len(y_val)), replace=False)
    ax.scatter(y_val.iloc[sample], preds[sample], alpha=0.2, s=8)
    lim = max(y_val.iloc[sample].max(), preds[sample].max())
    ax.plot([0, lim], [0, lim], "r--", linewidth=1.5, label="Perfect prediction")
    ax.set_xlabel("Actual duration (min)")
    ax.set_ylabel("Predicted duration (min)")
    ax.set_title(f"Predicted vs Actual Trip Duration\nMAE: {metrics['mae_min']:.2f} min   R²: {metrics['r2']:.3f}")
    ax.legend()
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(FIGURES_DIR / "predicted_vs_actual.png", dpi=120)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8, 5))
    importance.head(15).iloc[::-1].plot.barh(x="feature", y="importance", ax=ax, legend=False, color="#1f77b4")
    ax.set_title("LightGBM feature importance (gain)")
    ax.set_xlabel("Gain")
    fig.tight_layout()
    fig.savefig(FIGURES_DIR / "feature_importance.png", dpi=120)
    plt.close(fig)

    print(f"\n  Saved model + figures to {MODELS_DIR} and {FIGURES_DIR}")


def main() -> None:
    print("Loading processed dataset...")
    df = load_dataset()
    print(f"  Loaded {len(df):,} rows")

    print("\nPreparing features...")
    df = prepare(df)

    print("\nTraining...")
    model, metrics, X_val, y_val, preds = train(df)

    print("\nSaving artifacts...")
    save_artifacts(model, metrics, X_val, y_val, preds)
    print("\nDone.")


if __name__ == "__main__":
    main()
