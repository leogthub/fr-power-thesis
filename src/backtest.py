import pandas as pd
from typing import Callable
from src.evaluate import compute_all


def walk_forward_backtest(
    X: pd.DataFrame,
    y: pd.Series,
    train_fn: Callable,
    predict_fn: Callable,
    n_test_hours: int = 24 * 30,
    step_hours: int = 24 * 7,
) -> pd.DataFrame:
    """
    Walk-forward expanding-window backtest.
    Returns a DataFrame with columns [y_true, y_pred, fold].
    """
    results = []
    n = len(X)
    initial_train = n - n_test_hours

    fold = 0
    cursor = initial_train

    while cursor < n:
        test_end = min(cursor + step_hours, n)
        X_train, y_train = X.iloc[:cursor], y.iloc[:cursor]
        X_test, y_test = X.iloc[cursor:test_end], y.iloc[cursor:test_end]

        model = train_fn(X_train, y_train)
        y_pred = pd.Series(predict_fn(model, X_test), index=y_test.index)

        fold_df = pd.DataFrame({"y_true": y_test, "y_pred": y_pred, "fold": fold})
        results.append(fold_df)

        cursor = test_end
        fold += 1

    return pd.concat(results)


def summarise_backtest(results: pd.DataFrame) -> dict:
    return compute_all(results["y_true"], results["y_pred"])
