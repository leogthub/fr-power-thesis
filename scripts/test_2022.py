"""
Robustness check: out-of-sample validation on the 2022 energy crisis year.

Rationale
---------
The main pipeline tests on May 2024 – April 2025, a period of relatively
stable prices after the energy crisis.  A referee could argue that results
are overly optimistic because the test window avoids the most turbulent
period.  This script addresses that critique directly:

  Train : 2018-01-08 -> 2021-12-31  (4 years, pre-crisis)
  Test  : 2022-01-01 -> 2022-12-31  (full crisis year)

2022 is the most informative stress test for the French market:
  - EDF nuclear fleet at ~40 % availability (corrosion issues)
  - TTF gas prices 10× historical average (Aug 2022 peak: 339 EUR/MWh)
  - Day-ahead prices spiked above 1 000 EUR/MWh multiple times
  - Extreme price volatility unlike any post-2010 year

Models trained identically to the main pipeline (same hyperparameters,
same feature set).  Results saved alongside the main metrics so both
periods can be reported in the thesis Results section.

Outputs
-------
  outputs/metrics/validation_2022_A.json   (naive)
  outputs/metrics/validation_2022_B.json   (RF sans meteo)
  outputs/metrics/validation_2022_C.json   (RF avec meteo)
  outputs/metrics/validation_2022_D.json   (XGBoost)
  outputs/metrics/dm_tests_2022.json
  outputs/tables/model_comparison_2022.csv
"""
import json
import sys
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import PROCESSED_DIR, TABLES_DIR, METRICS_DIR
from src.features import build_feature_matrix, WEATHER_FEATURE_COLS
from src.models import train_random_forest, train_xgboost, naive_forecast
from src.evaluate import compute_all, dm_test

TABLES_DIR.mkdir(parents=True, exist_ok=True)
METRICS_DIR.mkdir(parents=True, exist_ok=True)

TRAIN_END = "2021-12-31 23:00:00"
TEST_START = "2022-01-01 00:00:00"
TEST_END   = "2022-12-31 23:00:00"


def main():
    print("=" * 60)
    print("FR Power Thesis — 2022 Crisis Robustness Validation")
    print("=" * 60)

    # ------------------------------------------------------------------
    # 1. Load data & build feature matrices
    # ------------------------------------------------------------------
    print("\n[1/4] Loading processed dataset...")
    df = pd.read_parquet(PROCESSED_DIR / "features.parquet")
    print(f"  Full dataset: {df.index.min().date()} -> {df.index.max().date()}")

    print("\n[2/4] Building feature matrices...")
    X,       y       = build_feature_matrix(df, include_weather=True)
    X_no_wx, y_no_wx = build_feature_matrix(df, include_weather=False)

    # ------------------------------------------------------------------
    # 2. 2022 train/test split
    # ------------------------------------------------------------------
    # Timezone-aware slicing — handle both tz-aware and tz-naive indices
    def tz_slice(series, start, end):
        idx = series.index
        if idx.tz is not None:
            start = pd.Timestamp(start).tz_localize(idx.tz)
            end   = pd.Timestamp(end).tz_localize(idx.tz)
            train_end_ts = pd.Timestamp(TRAIN_END).tz_localize(idx.tz)
        else:
            start = pd.Timestamp(start)
            end   = pd.Timestamp(end)
            train_end_ts = pd.Timestamp(TRAIN_END)
        train = series.loc[:train_end_ts]
        test  = series.loc[start:end]
        return train, test

    X_train,    X_test    = tz_slice(X,       TEST_START, TEST_END)
    y_train,    y_test    = tz_slice(y,        TEST_START, TEST_END)
    X_b_train,  X_b_test  = tz_slice(X_no_wx, TEST_START, TEST_END)
    y_b_train,  y_b_test  = tz_slice(y_no_wx, TEST_START, TEST_END)

    print(f"  Train: {len(X_train):,} hours  ({X_train.index.min().date()} -> {X_train.index.max().date()})")
    print(f"  Test : {len(X_test):,} hours   ({X_test.index.min().date()} -> {X_test.index.max().date()})")

    # ------------------------------------------------------------------
    # 3. Train models
    # ------------------------------------------------------------------
    print("\n[3/4] Training models on pre-crisis data (2018–2021)...")

    print("  Modèle A — Naive (lag-168h)...")
    # Naive uses the full y series to look back 168h from test start
    y_pred_naive = naive_forecast(y).loc[y_test.index]

    print("  Modèle B — Random Forest WITHOUT weather...")
    rf_b = train_random_forest(X_b_train, y_b_train)
    y_pred_rf_b = pd.Series(rf_b.predict(X_b_test), index=y_b_test.index)

    print("  Modèle C — Random Forest WITH weather...")
    rf_c = train_random_forest(X_train, y_train)
    y_pred_rf_c = pd.Series(rf_c.predict(X_test), index=y_test.index)

    print("  Modèle D — XGBoost WITH weather...")
    xgb = train_xgboost(X_train, y_train)
    y_pred_xgb = pd.Series(xgb.predict(X_test), index=y_test.index)

    # Align to common index
    common_idx = y_test.index
    for pred in [y_pred_naive, y_pred_rf_b, y_pred_rf_c, y_pred_xgb]:
        common_idx = common_idx.intersection(pred.dropna().index)

    y_true     = y_test.loc[common_idx]
    naive_pred = y_pred_naive.loc[common_idx]
    rf_b_pred  = y_pred_rf_b.loc[common_idx]
    rf_c_pred  = y_pred_rf_c.loc[common_idx]
    xgb_pred   = y_pred_xgb.loc[common_idx]

    # ------------------------------------------------------------------
    # 4. Evaluate & save
    # ------------------------------------------------------------------
    print("\n[4/4] Evaluating models on 2022 crisis test set...")

    results = {
        "A_naive":       compute_all(y_true, naive_pred),
        "B_rf_no_wx":    compute_all(y_true, rf_b_pred),
        "C_rf_wx":       compute_all(y_true, rf_c_pred),
        "D_xgboost":     compute_all(y_true, xgb_pred),
    }

    print(f"\n{'Model':<22} {'MAE':>8} {'RMSE':>8} {'sMAPE':>8} {'R²':>8} {'Hit%':>7}")
    print("-" * 60)
    for name, m in results.items():
        print(f"  {name:<20} {m['mae']:>7.2f} {m['rmse']:>8.2f} "
              f"{m['smape']:>7.1f}% {m['r2']:>7.3f} {m['hit_ratio']:>6.1f}%")

    # Save individual JSON metrics
    label_map = {
        "A_naive":    "validation_2022_A",
        "B_rf_no_wx": "validation_2022_B",
        "C_rf_wx":    "validation_2022_C",
        "D_xgboost":  "validation_2022_D",
    }
    for key, fname in label_map.items():
        path = METRICS_DIR / f"{fname}.json"
        with open(path, "w") as f:
            json.dump(results[key], f, indent=2)
        print(f"  Saved: {path}")

    # DM tests
    print("\n  DM tests (2022 crisis period):")
    dm_results = {}
    pairs = [
        ("C_vs_B_2022", rf_c_pred, rf_b_pred, "Weather vs no-weather (C vs B)"),
        ("C_vs_A_2022", rf_c_pred, naive_pred, "RF-weather vs naive (C vs A)"),
        ("B_vs_A_2022", rf_b_pred, naive_pred, "RF-no-weather vs naive (B vs A)"),
        ("D_vs_C_2022", xgb_pred,  rf_c_pred,  "XGBoost vs RF-weather (D vs C)"),
    ]
    print(f"  {'Test':<38} {'DM stat':>9} {'p-value':>9} {'Signif.':>9}")
    print("  " + "-" * 68)
    for key, p1, p2, label in pairs:
        stat, pval = dm_test(y_true, p1, p2)
        sig = "***" if pval < 0.01 else ("**" if pval < 0.05 else ("*" if pval < 0.10 else "n.s."))
        dm_results[key] = {
            "description": label,
            "dm_stat": round(stat, 4),
            "p_value": round(pval, 4),
            "significant_5pct": pval < 0.05,
        }
        print(f"  {label:<38} {stat:>9.3f} {pval:>9.4f} {sig:>9}")

    dm_path = METRICS_DIR / "dm_tests_2022.json"
    with open(dm_path, "w") as f:
        json.dump(dm_results, f, indent=2)
    print(f"\n  DM-test 2022 results -> {dm_path}")

    # Comparison table
    model_labels = {
        "A_naive":    "Naif (lag-168h)",
        "B_rf_no_wx": "Random Forest (sans meteo)",
        "C_rf_wx":    "Random Forest (avec meteo)",
        "D_xgboost":  "XGBoost (avec meteo)",
    }
    rows = []
    for key, label in model_labels.items():
        m = results[key]
        rows.append({
            "Model":         label,
            "MAE (EUR/MWh)":   round(m["mae"], 3),
            "RMSE (EUR/MWh)":  round(m["rmse"], 3),
            "sMAPE (%)":     round(m["smape"], 3),
            "R²":            round(m["r2"], 3),
            "Hit ratio (%)": round(m["hit_ratio"], 1),
        })
    comp_df = pd.DataFrame(rows)
    csv_path = TABLES_DIR / "model_comparison_2022.csv"
    comp_df.to_csv(csv_path, index=False)
    print(f"\nComparison table (2022) -> {csv_path}")
    print(comp_df.to_string(index=False))

    print("\nOK 2022 crisis robustness validation complete.")


if __name__ == "__main__":
    main()
