import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns
import pandas as pd
import numpy as np
from src.config import FIGURES_DIR

sns.set_theme(style="whitegrid", palette="muted")
FIGURES_DIR.mkdir(parents=True, exist_ok=True)


def plot_price_series(prices: pd.Series, title: str = "Day-Ahead Prices — France", save_as: str = None):
    fig, ax = plt.subplots(figsize=(14, 4))
    ax.plot(prices.index, prices.values, linewidth=0.8)
    ax.set_title(title)
    ax.set_ylabel("€/MWh")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    fig.autofmt_xdate()
    _save(fig, save_as)
    return fig


def plot_forecast_vs_actual(y_true: pd.Series, y_pred: pd.Series, save_as: str = None):
    fig, ax = plt.subplots(figsize=(14, 4))
    ax.plot(y_true.index, y_true.values, label="Actual", linewidth=0.9)
    ax.plot(y_pred.index, y_pred.values, label="Forecast", linewidth=0.9, linestyle="--")
    ax.set_title("Forecast vs Actual — Day-Ahead Price")
    ax.set_ylabel("€/MWh")
    ax.legend()
    _save(fig, save_as)
    return fig


def plot_feature_importance(model, feature_names: list, top_n: int = 20, save_as: str = None):
    importances = pd.Series(model.feature_importances_, index=feature_names)
    importances = importances.nlargest(top_n).sort_values()
    fig, ax = plt.subplots(figsize=(8, top_n * 0.35))
    importances.plot(kind="barh", ax=ax)
    ax.set_title(f"Top {top_n} Feature Importances")
    ax.set_xlabel("Importance")
    _save(fig, save_as)
    return fig


def plot_error_distribution(y_true: pd.Series, y_pred: pd.Series, save_as: str = None):
    errors = y_pred - y_true
    fig, axes = plt.subplots(1, 2, figsize=(12, 4))
    axes[0].hist(errors, bins=60, edgecolor="white")
    axes[0].set_title("Forecast Error Distribution")
    axes[0].set_xlabel("Error (€/MWh)")
    axes[1].scatter(y_true, y_pred, alpha=0.3, s=5)
    lims = [min(y_true.min(), y_pred.min()), max(y_true.max(), y_pred.max())]
    axes[1].plot(lims, lims, "r--", linewidth=1)
    axes[1].set_title("Actual vs Predicted")
    axes[1].set_xlabel("Actual (€/MWh)")
    axes[1].set_ylabel("Predicted (€/MWh)")
    _save(fig, save_as)
    return fig


def _save(fig: plt.Figure, name: str = None):
    if name:
        fig.tight_layout()
        fig.savefig(FIGURES_DIR / name, dpi=150, bbox_inches="tight")
