"""
Phase 3 — Trading Backtest (French Day-Ahead Electricity Market)

Strategy (replication of German thesis approach, adapted to FR):
  At each hour h, compare the model's predicted price to the last observed price.
  - If pred(h) > actual(h-1): expect a price RISE  -> go LONG  (+1 MW)
  - If pred(h) < actual(h-1): expect a price FALL  -> go SHORT (-1 MW)
  - P&L(h) = signal(h) * (actual(h) - actual(h-1))   [EUR/MWh per MW]

This is a pure directional strategy: no transaction costs, 1 MW position size.
Position is held for exactly 1 hour, then rebalanced.

Models compared:
  A — Naive lag-168h (benchmark)
  B — Random Forest without weather
  C — Random Forest with weather  (main model)
  D — XGBoost with weather
  E — Long-only (always long, upper-bound reference)

Metrics computed per model:
  - Total P&L (EUR/MW over test period)
  - Annualised Sharpe Ratio  = mean(pnl_h) / std(pnl_h) * sqrt(8760)
  - Maximum Drawdown (%)
  - Calmar Ratio             = Annualised P&L / |Max Drawdown|
  - Win Rate                 = % hours with positive P&L

Outputs:
  outputs/metrics/backtest_results.json
  outputs/tables/backtest_comparison.csv
  outputs/figures/equity_curves.png
  outputs/figures/monthly_pnl.png
  outputs/figures/pnl_distribution.png
"""
import json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR   = Path(__file__).resolve().parents[1]
PREDS_PATH = BASE_DIR / "outputs" / "predictions" / "test_predictions.parquet"
METRICS_DIR = BASE_DIR / "outputs" / "metrics"
TABLES_DIR  = BASE_DIR / "outputs" / "tables"
FIGS_DIR    = BASE_DIR / "outputs" / "figures"

for d in [METRICS_DIR, TABLES_DIR, FIGS_DIR]:
    d.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Core trading functions
# ---------------------------------------------------------------------------

def compute_signals(actual: pd.Series, predicted: pd.Series) -> pd.Series:
    """
    Signal = sign(pred(h) - actual(h-1)).
    +1 = LONG (expect price rise), -1 = SHORT (expect price fall).
    First hour is dropped (no lagged actual available).
    """
    actual_lag = actual.shift(1)
    signal = np.sign(predicted - actual_lag)
    signal = signal.replace(0, 1)   # ties -> long
    return signal.dropna()


def compute_pnl(actual: pd.Series, signal: pd.Series) -> pd.Series:
    """
    P&L per hour = signal(h) * (actual(h) - actual(h-1)).
    Units: EUR per MWh per MW of position.
    """
    delta = actual.diff()
    pnl   = signal * delta.loc[signal.index]
    return pnl.dropna()


def sharpe(pnl: pd.Series, periods_per_year: int = 252) -> float:
    """
    Annualised Sharpe ratio.
    Aggregates hourly P&L to daily first, then annualises with sqrt(252).
    This avoids the inflated Sharpe from sqrt(8760) on correlated hourly data.
    No risk-free rate (power trading context).
    """
    if len(pnl) == 0:
        return 0.0
    # resample to daily P&L
    daily = pnl.resample("D").sum()
    if daily.std() == 0:
        return 0.0
    return float(daily.mean() / daily.std() * np.sqrt(periods_per_year))


def max_drawdown(cumulative_pnl: pd.Series) -> float:
    """
    Maximum drawdown in absolute EUR/MW terms.
    Measures the largest peak-to-trough decline in cumulative P&L.
    Robust to strategies that start negative (uses running max from -inf).
    """
    roll_max = cumulative_pnl.cummax()
    drawdown = roll_max - cumulative_pnl
    return float(drawdown.max())


def calmar(pnl: pd.Series, cumulative_pnl: pd.Series,
           periods_per_year: int = 8760) -> float:
    """
    Calmar ratio = annualised P&L (EUR/MW/year) / absolute max drawdown (EUR/MW).
    Dimensionless — higher is better.
    """
    ann_return = pnl.mean() * periods_per_year
    mdd = max_drawdown(cumulative_pnl)
    if mdd == 0:
        return np.inf
    return float(ann_return / mdd)


def win_rate(pnl: pd.Series) -> float:
    """Percentage of hours with strictly positive P&L."""
    return float((pnl > 0).mean() * 100)


def run_strategy(actual: pd.Series, predicted: pd.Series,
                 label: str) -> dict:
    """Full strategy pipeline for one model."""
    signal  = compute_signals(actual, predicted)
    pnl     = compute_pnl(actual, signal)
    cum_pnl = pnl.cumsum()

    ann_pnl = float(pnl.resample("D").sum().mean() * 252)  # EUR/MW/year via daily
    sr      = sharpe(pnl)
    mdd     = max_drawdown(cum_pnl)
    cal     = calmar(pnl, cum_pnl)
    wr      = win_rate(pnl)

    return {
        "label":           label,
        "total_pnl":       float(cum_pnl.iloc[-1]),
        "annualised_pnl":  ann_pnl,
        "sharpe":          sr,
        "max_drawdown":    mdd,
        "calmar":          cal,
        "win_rate":        wr,
        "pnl_series":      pnl,
        "cum_pnl":         cum_pnl,
    }


# ---------------------------------------------------------------------------
# Plotting helpers
# ---------------------------------------------------------------------------
PALETTE = {
    "Naif (lag-168h)":               "#9E9E9E",
    "RF sans meteo":                  "#2196F3",
    "RF avec meteo":                  "#4CAF50",
    "XGBoost":                        "#FF9800",
    "Long-only":                      "#E91E63",
}


def plot_equity_curves(strategies: dict):
    fig, ax = plt.subplots(figsize=(12, 6))
    for key, s in strategies.items():
        color = PALETTE.get(s["label"], "#333333")
        lw    = 2.5 if "RF avec" in s["label"] else 1.5
        ls    = "--" if s["label"] == "Long-only" else "-"
        ax.plot(s["cum_pnl"].index, s["cum_pnl"].values,
                label=s["label"], color=color, linewidth=lw, linestyle=ls)

    ax.axhline(0, color="black", linewidth=0.8, linestyle=":")
    ax.set_title("Courbes de performance cumulee (EUR/MW)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Date")
    ax.set_ylabel("P&L cumule (EUR/MW)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.legend(loc="upper left", fontsize=10)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    path = FIGS_DIR / "equity_curves.png"
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")


def plot_monthly_pnl(strategies: dict):
    """Monthly P&L bar chart comparing RF-weather vs Naive."""
    fig, ax = plt.subplots(figsize=(14, 5))

    keys_to_plot = [k for k in strategies
                    if strategies[k]["label"] in ("RF avec meteo", "Naif (lag-168h)")]

    n_models = len(keys_to_plot)
    bar_width = 0.35
    months = None

    for i, key in enumerate(keys_to_plot):
        s   = strategies[key]
        mp  = s["pnl_series"].resample("ME").sum()
        mp.index = mp.index.to_period("M").to_timestamp()
        if months is None:
            months = mp.index
        x = np.arange(len(mp))
        offset = (i - n_models / 2 + 0.5) * bar_width
        color = PALETTE.get(s["label"], "#888")
        ax.bar(x + offset, mp.values, bar_width,
               label=s["label"], color=color, alpha=0.85)

    if months is not None:
        ax.set_xticks(np.arange(len(months)))
        ax.set_xticklabels([m.strftime("%b %Y") for m in months],
                           rotation=45, ha="right", fontsize=8)
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_title("P&L mensuel par modele (EUR/MW)", fontsize=13, fontweight="bold")
    ax.set_ylabel("P&L (EUR/MW)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3, axis="y")
    plt.tight_layout()
    path = FIGS_DIR / "monthly_pnl.png"
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")


def plot_pnl_distribution(strategies: dict):
    """Hourly P&L distribution boxplot for all models."""
    fig, ax = plt.subplots(figsize=(10, 5))

    data   = [s["pnl_series"].values for s in strategies.values()]
    labels = [s["label"]             for s in strategies.values()]
    colors = [PALETTE.get(s["label"], "#888") for s in strategies.values()]

    bp = ax.boxplot(data, patch_artist=True, tick_labels=labels,
                    whis=1.5, showfliers=False, medianprops={"color": "black"})
    for patch, color in zip(bp["boxes"], colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.7)

    ax.axhline(0, color="black", linewidth=0.8, linestyle=":")
    ax.set_title("Distribution du P&L horaire (EUR/MW)", fontsize=13, fontweight="bold")
    ax.set_ylabel("P&L horaire (EUR/MW)")
    ax.grid(True, alpha=0.3, axis="y")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    path = FIGS_DIR / "pnl_distribution.png"
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("FR Power Thesis -- Phase 3: Trading Backtest")
    print("=" * 60)

    # -----------------------------------------------------------------------
    # 1. Load predictions
    # -----------------------------------------------------------------------
    print(f"\n[1/4] Loading predictions from {PREDS_PATH}...")
    if not PREDS_PATH.exists():
        raise FileNotFoundError(
            f"Predictions file not found: {PREDS_PATH}\n"
            "Run scripts/run_full_pipeline.py first."
        )
    df = pd.read_parquet(PREDS_PATH)
    print(f"  {len(df):,} hourly observations")
    print(f"  Period: {df.index.min().date()} -> {df.index.max().date()}")
    print(f"  Columns: {list(df.columns)}")

    actual = df["actual"]

    # Long-only benchmark: always long (signal = +1 always)
    long_only_signal = pd.Series(1.0, index=actual.index[1:])
    pnl_long_only    = (actual.diff().dropna())

    # -----------------------------------------------------------------------
    # 2. Run strategies
    # -----------------------------------------------------------------------
    print("\n[2/4] Running trading strategies...")

    model_map = {
        "naive":         ("Naif (lag-168h)",  df["naive"]),
        "rf_no_weather": ("RF sans meteo",    df["rf_no_weather"]),
        "rf_weather":    ("RF avec meteo",    df["rf_weather"]),
        "xgboost":       ("XGBoost",          df["xgboost"]),
    }

    strategies = {}
    for key, (label, predicted) in model_map.items():
        strategies[key] = run_strategy(actual, predicted, label)
        s = strategies[key]
        print(f"  {label:<28} | P&L: {s['total_pnl']:>8,.0f} EUR/MW "
              f"| Sharpe: {s['sharpe']:>6.3f} "
              f"| MDD: {s['max_drawdown']:>8,.0f} EUR/MW "
              f"| WinRate: {s['win_rate']:>5.1f}%")

    # Long-only
    cum_long = pnl_long_only.cumsum()
    strategies["long_only"] = {
        "label":           "Long-only",
        "total_pnl":       float(cum_long.iloc[-1]),
        "annualised_pnl":  float(pnl_long_only.resample("D").sum().mean() * 252),
        "sharpe":          sharpe(pnl_long_only),
        "max_drawdown":    max_drawdown(cum_long),
        "calmar":          calmar(pnl_long_only, cum_long),
        "win_rate":        win_rate(pnl_long_only),
        "pnl_series":      pnl_long_only,
        "cum_pnl":         cum_long,
    }
    s = strategies["long_only"]
    print(f"  {'Long-only (ref)':<28} | P&L: {s['total_pnl']:>8,.0f} EUR/MW "
          f"| Sharpe: {s['sharpe']:>6.3f} "
          f"| MDD: {s['max_drawdown']:>8,.0f} EUR/MW "
          f"| WinRate: {s['win_rate']:>5.1f}%")

    # -----------------------------------------------------------------------
    # 3. Save results
    # -----------------------------------------------------------------------
    print("\n[3/4] Saving results...")

    # JSON — exclude series objects
    json_out = {}
    for key, s in strategies.items():
        json_out[key] = {k: v for k, v in s.items()
                         if k not in ("pnl_series", "cum_pnl")}
    bt_path = METRICS_DIR / "backtest_results.json"
    with open(bt_path, "w") as f:
        json.dump(json_out, f, indent=2)
    print(f"  Saved: {bt_path}")

    # CSV comparison table
    rows = []
    col_order = ["label", "total_pnl", "annualised_pnl",
                 "sharpe", "max_drawdown", "calmar", "win_rate"]
    col_labels = {
        "label":          "Modele",
        "total_pnl":      "P&L total (EUR/MW)",
        "annualised_pnl": "P&L annualise (EUR/MW/an)",
        "sharpe":         "Sharpe (daily, ann.)",
        "max_drawdown":   "Max Drawdown (EUR/MW)",
        "calmar":         "Ratio Calmar",
        "win_rate":       "Win Rate (%)",
    }
    for key in ["naive", "rf_no_weather", "rf_weather", "xgboost", "long_only"]:
        s = strategies[key]
        rows.append({col_labels[c]: round(s[c], 3) if isinstance(s[c], float)
                     else s[c]
                     for c in col_order})
    comp_df = pd.DataFrame(rows)
    csv_path = TABLES_DIR / "backtest_comparison.csv"
    comp_df.to_csv(csv_path, index=False)
    print(f"  Saved: {csv_path}")
    print()
    print(comp_df.to_string(index=False))

    # -----------------------------------------------------------------------
    # 4. Plots
    # -----------------------------------------------------------------------
    print("\n[4/4] Generating plots...")
    plot_equity_curves(strategies)
    plot_monthly_pnl(strategies)
    plot_pnl_distribution(strategies)

    print("\n✓ Phase 3 backtest complete. All outputs saved to outputs/")


if __name__ == "__main__":
    main()
