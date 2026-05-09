import json
import numpy as np
import pandas as pd
from pathlib import Path
from scipy import stats
from src.config import METRICS_DIR


def mae(y_true: pd.Series, y_pred: pd.Series) -> float:
    return float(np.mean(np.abs(y_true - y_pred)))


def rmse(y_true: pd.Series, y_pred: pd.Series) -> float:
    return float(np.sqrt(np.mean((y_true - y_pred) ** 2)))


def smape(y_true: pd.Series, y_pred: pd.Series) -> float:
    """Symmetric MAPE — robust to near-zero and negative prices."""
    denom = (np.abs(y_true) + np.abs(y_pred)) / 2
    mask = denom > 0
    return float(np.mean(np.abs(y_true[mask] - y_pred[mask]) / denom[mask]) * 100)


def r2(y_true: pd.Series, y_pred: pd.Series) -> float:
    ss_res = np.sum((y_true - y_pred) ** 2)
    ss_tot = np.sum((y_true - y_true.mean()) ** 2)
    return float(1 - ss_res / ss_tot)


def hit_ratio(y_true: pd.Series, y_pred: pd.Series) -> float:
    """Directional accuracy: % of hours where price movement direction is correct."""
    prev = y_true.shift(1).dropna()
    actual_dir = np.sign(y_true.loc[prev.index] - prev)
    pred_dir = np.sign(y_pred.loc[prev.index] - prev)
    return float((actual_dir == pred_dir).mean() * 100)


def dm_test(
    y_true: pd.Series,
    y_pred1: pd.Series,
    y_pred2: pd.Series,
    h: int = 1,
) -> tuple[float, float]:
    """
    Diebold-Mariano test (1995) — H0: equal predictive accuracy (MAE-based).
    Returns (DM statistic, two-sided p-value).
    A negative DM stat means model 1 is MORE accurate than model 2.
    """
    e1 = y_true - y_pred1
    e2 = y_true - y_pred2
    d = np.abs(e1.values) - np.abs(e2.values)
    n = len(d)
    d_mean = d.mean()

    # Newey-West long-run variance with h-1 lags
    gamma0 = np.var(d, ddof=1)
    gamma_sum = sum(
        (1 - k / h) * np.cov(d[k:], d[:-k])[0, 1]
        for k in range(1, h)
    ) if h > 1 else 0
    var_d = (gamma0 + 2 * gamma_sum) / n

    if var_d <= 0:
        return float("nan"), float("nan")

    dm_stat = d_mean / np.sqrt(var_d)
    # Harvey, Leybourne & Newbold (1997) small-sample correction
    hln_factor = np.sqrt((n + 1 - 2 * h + h * (h - 1) / n) / n)
    dm_stat_adj = dm_stat * hln_factor
    p_value = float(2 * (1 - stats.norm.cdf(abs(dm_stat_adj))))
    return float(dm_stat_adj), p_value


def compute_all(y_true: pd.Series, y_pred: pd.Series) -> dict:
    return {
        "mae": mae(y_true, y_pred),
        "rmse": rmse(y_true, y_pred),
        "smape": smape(y_true, y_pred),
        "r2": r2(y_true, y_pred),
        "hit_ratio": hit_ratio(y_true, y_pred),
    }


def save_metrics(metrics: dict, name: str) -> Path:
    METRICS_DIR.mkdir(parents=True, exist_ok=True)
    path = METRICS_DIR / f"{name}.json"
    with open(path, "w") as f:
        json.dump(metrics, f, indent=2)
    return path
