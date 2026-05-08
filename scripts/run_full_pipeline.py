"""
End-to-end pipeline: fetch data → preprocess → features → train → evaluate → save outputs.
"""
import pandas as pd
from src.config import PROCESSED_DIR, TEST_MONTHS
from src.features import build_feature_matrix
from src.models import train_random_forest, train_xgboost, naive_forecast
from src.evaluate import compute_all, save_metrics
from src.backtest import walk_forward_backtest, summarise_backtest
from src.plots import plot_forecast_vs_actual, plot_feature_importance

PROCESSED_DIR.mkdir(parents=True, exist_ok=True)


def main():
    print("Loading processed dataset...")
    df = pd.read_parquet(PROCESSED_DIR / "features.parquet")

    print("Building feature matrix...")
    X, y = build_feature_matrix(df)

    cutoff = len(X) - TEST_MONTHS * 30 * 24
    X_train, X_test = X.iloc[:cutoff], X.iloc[cutoff:]
    y_train, y_test = y.iloc[:cutoff], y.iloc[cutoff:]

    print("Training Random Forest...")
    rf = train_random_forest(X_train, y_train)
    y_pred_rf = pd.Series(rf.predict(X_test), index=y_test.index)

    print("Training XGBoost...")
    xgb = train_xgboost(X_train, y_train)
    y_pred_xgb = pd.Series(xgb.predict(X_test), index=y_test.index)

    y_pred_naive = naive_forecast(y).loc[y_test.index]

    metrics = {
        "random_forest": compute_all(y_test, y_pred_rf),
        "xgboost": compute_all(y_test, y_pred_xgb),
        "naive": compute_all(y_test.dropna(), y_pred_naive.dropna()),
    }

    for name, m in metrics.items():
        save_metrics(m, name)
        print(f"{name}: MAE={m['mae']:.2f} RMSE={m['rmse']:.2f} MAPE={m['mape']:.2f}% R²={m['r2']:.3f}")

    plot_forecast_vs_actual(y_test, y_pred_rf, save_as="rf_forecast.png")
    plot_feature_importance(rf, X_train.columns.tolist(), save_as="rf_feature_importance.png")

    print("Done. Outputs saved to outputs/")


if __name__ == "__main__":
    main()
