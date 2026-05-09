"""
End-to-end pipeline: load processed data -> train 4 models -> evaluate -> DM-test -> save outputs.

Models trained:
  A — Naive (benchmark) : price(h) = price(h-168), same hour previous week
  B — Random Forest WITHOUT weather features
  C — Random Forest WITH weather features   ← main model
  D — XGBoost WITH weather features

Statistical tests:
  DM-test (Diebold-Mariano, 1995) with Harvey-Leybourne-Newbold correction:
    C vs B : does adding weather significantly improve over RF baseline?
    C vs A : does RF-weather beat naive?
    B vs A : does RF (no weather) beat naive?
"""
import json
import pandas as pd
from pathlib import Path
from src.config import PROCESSED_DIR, TABLES_DIR, METRICS_DIR, TEST_MONTHS
from src.features import build_feature_matrix, WEATHER_FEATURE_COLS
from src.models import train_random_forest, train_xgboost, naive_forecast
from src.evaluate import compute_all, save_metrics, dm_test
from src.plots import plot_forecast_vs_actual, plot_feature_importance

TABLES_DIR.mkdir(parents=True, exist_ok=True)
METRICS_DIR.mkdir(parents=True, exist_ok=True)


def main():
    print("=" * 60)
    print("FR Power Thesis — Full Training Pipeline")
    print("=" * 60)

    # ------------------------------------------------------------------
    # 1. Load data & build feature matrices
    # ------------------------------------------------------------------
    print("\n[1/5] Loading processed dataset...")
    df = pd.read_parquet(PROCESSED_DIR / "features.parquet")
    print(f"  {len(df):,} rows, {df.shape[1]} columns")
    print(f"  Period: {df.index.min()} -> {df.index.max()}")

    print("\n[2/5] Building feature matrices...")

    # Modèle C / D — with full weather features
    X, y = build_feature_matrix(df, include_weather=True)
    print(f"  X_weather   : {X.shape[0]:,} rows × {X.shape[1]} features")

    # Modèle B — without weather (drop ERA5 and derived weather cols)
    X_no_wx, y_no_wx = build_feature_matrix(df, include_weather=False)
    print(f"  X_no_weather: {X_no_wx.shape[0]:,} rows × {X_no_wx.shape[1]} features")

    weather_features_used = [c for c in WEATHER_FEATURE_COLS if c in X.columns]
    print(f"  Weather cols dropped for Modèle B: {weather_features_used}")

    # ------------------------------------------------------------------
    # 2. Train/test split (last TEST_MONTHS months = out-of-sample)
    # ------------------------------------------------------------------
    cutoff = len(X) - TEST_MONTHS * 30 * 24
    X_train,    X_test    = X.iloc[:cutoff],       X.iloc[cutoff:]
    y_train,    y_test    = y.iloc[:cutoff],        y.iloc[cutoff:]

    cutoff_b = len(X_no_wx) - TEST_MONTHS * 30 * 24
    X_b_train, X_b_test  = X_no_wx.iloc[:cutoff_b], X_no_wx.iloc[cutoff_b:]
    y_b_train, y_b_test  = y_no_wx.iloc[:cutoff_b],  y_no_wx.iloc[cutoff_b:]

    print(f"\n  Train: {len(X_train):,} hours | Test: {len(X_test):,} hours")
    print(f"  Test period: {y_test.index.min()} -> {y_test.index.max()}")

    # ------------------------------------------------------------------
    # 3. Train models
    # ------------------------------------------------------------------
    print("\n[3/5] Training models...")

    print("  Modèle A — Naive (lag-168h)...")
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

    # Align all predictions to the same test index (intersection)
    common_idx = y_test.index
    for pred in [y_pred_naive, y_pred_rf_b, y_pred_rf_c, y_pred_xgb]:
        common_idx = common_idx.intersection(pred.dropna().index)

    y_true     = y_test.loc[common_idx]
    naive_pred = y_pred_naive.loc[common_idx]
    rf_b_pred  = y_pred_rf_b.loc[common_idx]
    rf_c_pred  = y_pred_rf_c.loc[common_idx]
    xgb_pred   = y_pred_xgb.loc[common_idx]

    # ------------------------------------------------------------------
    # 4. Evaluate models
    # ------------------------------------------------------------------
    print("\n[4/5] Evaluating models...")

    results = {
        "naive":          compute_all(y_true, naive_pred),
        "rf_no_weather":  compute_all(y_true, rf_b_pred),
        "rf_weather":     compute_all(y_true, rf_c_pred),
        "xgboost":        compute_all(y_true, xgb_pred),
    }

    print(f"\n{'Model':<20} {'MAE':>8} {'RMSE':>8} {'sMAPE':>8} {'R²':>8} {'Hit%':>7}")
    print("-" * 60)
    for name, m in results.items():
        print(
            f"  {name:<18} {m['mae']:>7.2f} {m['rmse']:>8.2f} "
            f"{m['smape']:>7.1f}% {m['r2']:>7.3f} {m['hit_ratio']:>6.1f}%"
        )

    for name, m in results.items():
        save_metrics(m, name)

    # ------------------------------------------------------------------
    # 5. Diebold-Mariano tests
    # ------------------------------------------------------------------
    print("\n[5/5] Diebold-Mariano tests...")

    dm_results = {}

    pairs = [
        ("rf_weather_vs_rf_no_weather", rf_c_pred, rf_b_pred,
         "Does weather add value? (C vs B)"),
        ("rf_weather_vs_naive",         rf_c_pred, naive_pred,
         "Does RF-weather beat naive? (C vs A)"),
        ("rf_no_weather_vs_naive",      rf_b_pred, naive_pred,
         "Does RF-no-weather beat naive? (B vs A)"),
        ("xgboost_vs_rf_weather",       xgb_pred,  rf_c_pred,
         "XGBoost vs RF-weather (D vs C)"),
    ]

    print(f"\n  {'Test':<40} {'DM stat':>9} {'p-value':>9} {'Signif.':>9}")
    print("  " + "-" * 70)
    for key, pred1, pred2, label in pairs:
        stat, pval = dm_test(y_true, pred1, pred2)
        sig = "***" if pval < 0.01 else ("**" if pval < 0.05 else ("*" if pval < 0.10 else "n.s."))
        dm_results[key] = {
            "description": label,
            "dm_stat": round(stat, 4),
            "p_value": round(pval, 4),
            "significant_5pct": pval < 0.05,
        }
        print(f"  {label:<40} {stat:>9.3f} {pval:>9.4f} {sig:>9}")

    dm_path = METRICS_DIR / "dm_tests.json"
    with open(dm_path, "w") as f:
        json.dump(dm_results, f, indent=2)
    print(f"\n  DM-test results saved -> {dm_path}")

    # ------------------------------------------------------------------
    # 6. Save comparison table
    # ------------------------------------------------------------------
    model_labels = {
        "naive":         "Naïf (lag-168h)",
        "rf_no_weather": "Random Forest (sans météo)",
        "rf_weather":    "Random Forest (avec météo)",
        "xgboost":       "XGBoost (avec météo)",
    }
    rows = []
    for key, label in model_labels.items():
        m = results[key]
        rows.append({
            "Model": label,
            "MAE (€/MWh)": round(m["mae"], 3),
            "RMSE (€/MWh)": round(m["rmse"], 3),
            "sMAPE (%)": round(m["smape"], 3),
            "R²": round(m["r2"], 3),
            "Hit ratio (%)": round(m["hit_ratio"], 1),
        })
    comp_df = pd.DataFrame(rows)
    comp_path = TABLES_DIR / "model_comparison.csv"
    comp_df.to_csv(comp_path, index=False)
    print(f"\nModel comparison table saved -> {comp_path}")
    print(comp_df.to_string(index=False))

    # ------------------------------------------------------------------
    # 7. Plots
    # ------------------------------------------------------------------
    print("\nGenerating plots...")
    try:
        plot_forecast_vs_actual(y_true, rf_c_pred, save_as="forecast_vs_actual.png")
        plot_feature_importance(rf_c, X_train.columns.tolist(), save_as="feature_importance_rf.png")
        print("  Plots saved to outputs/figures/")
    except Exception as e:
        print(f"  WARNING: plot generation failed: {e}")

    print("\n✓ Pipeline complete. All outputs saved to outputs/")


if __name__ == "__main__":
    main()
