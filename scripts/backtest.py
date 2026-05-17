"""
Phase 3 - Trading Backtest (French Day-Ahead Electricity Market)
================================================================

DESIGN RATIONALE
----------------
EPEX SPOT France day-ahead prices are set in a single batch auction at noon
(12:00 CET) on day D for all 24 hours of day D+1.  Any executable strategy
must therefore use only information available before that gate closure.

Signal (day-ahead correct)
--------------------------
  At noon on day D we know all hourly prices of day D.
  For each hour h of day D+1 we compare our model forecast to the same
  hour on day D (the last comparable observation):

      signal(h, D+1) = sign( pred(h, D+1) - actual(h, D) )

  +1 = expect the price for hour h tomorrow to be HIGHER than today -> LONG
  -1 = expect it to be LOWER                                         -> SHORT
   0 = |pred - ref| < THRESHOLD  (dead-band, no trade)

  Dead-band: 2 EUR/MWh  -- roughly the 20th percentile of |day-on-day delta|
  on EPEX France.  Below this threshold the edge does not cover even minimum
  execution costs.

P&L (day-on-day, consistent with signal)
-----------------------------------------
  gross_pnl(h, D+1) = signal(h, D+1) * ( actual(h, D+1) - actual(h, D) )

  Long position profits if tomorrow's price > today's price for the same hour.
  Short position profits if tomorrow's price < today's price for the same hour.
  Units: EUR per MWh per MW of position (1 MW assumed throughout).

Transaction costs (EPEX SPOT France fee schedule 2024)
-------------------------------------------------------
  Costs apply to EVERY active hour (you submit bids every day for each hour
  you want to trade -- not just when you change direction).

  Optimistic  0.10 EUR/MWh  EPEX matching fee only (~0.04) + minimal slippage
  Central     0.30 EUR/MWh  + bid-ask proxy + admin overhead
  Pessimistic 0.60 EUR/MWh  + imbalance settlement risk in low-liquidity hours

  net_pnl(h) = gross_pnl(h) - cost * |signal(h)|

Risk metrics (methodology)
--------------------------
  Sharpe : aggregated to DAILY, annualised with sqrt(365) -- electricity
           trades every calendar day.  No risk-free rate (power convention).
  Max Drawdown  : peak-to-trough on cumulative NET P&L.
  Calmar        : annualised daily P&L / |MDD|.
  Profit Factor : sum(positive P&L) / |sum(negative P&L)| -- measures
                  consistency of gains vs losses.
  Monthly Sharpe: rolling month-by-month net Sharpe -- key robustness check.
  Peak / Off-peak split: hours 8-20 (peak) vs rest (off-peak).

Models
------
  A  Naive lag-168h (benchmark)
  B  Random Forest without weather features
  C  Random Forest with weather features  (main model)
  D  XGBoost with weather features
  L  Long-only reference (always long, buy-and-hold equivalent)

Outputs
-------
  outputs/metrics/backtest_results.json
  outputs/tables/backtest_comparison.csv       central scenario
  outputs/tables/backtest_sensitivity.csv      3 scenarios x 4 models
  outputs/tables/backtest_monthly_sharpe.csv   monthly Sharpe by model
  outputs/figures/equity_curves.png            gross (no costs)
  outputs/figures/equity_curves_net.png        net of costs, central
  outputs/figures/monthly_pnl.png
  outputs/figures/pnl_distribution.png
  outputs/figures/cost_sensitivity.png
  outputs/figures/rolling_sharpe.png           30-day rolling Sharpe
"""

import json
import warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning)

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------
BASE_DIR    = Path(__file__).resolve().parents[1]
PREDS_PATH  = BASE_DIR / "outputs" / "predictions" / "test_predictions.parquet"
METRICS_DIR = BASE_DIR / "outputs" / "metrics"
TABLES_DIR  = BASE_DIR / "outputs" / "tables"
FIGS_DIR    = BASE_DIR / "outputs" / "figures"

for d in [METRICS_DIR, TABLES_DIR, FIGS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

COST_SCENARIOS = {
    "optimiste":   0.10,
    "central":     0.30,
    "pessimiste":  0.60,
}
COST_CENTRAL   = COST_SCENARIOS["central"]
THRESHOLD      = 2.0    # EUR/MWh dead-band
PEAK_HOURS     = range(8, 20)   # hours 8-19 inclusive = peak
ANN_FACTOR     = 365    # calendar days (electricity market)

PALETTE = {
    "Naif lag-168h":  "#9E9E9E",
    "RF sans meteo":  "#2196F3",
    "RF avec meteo":  "#4CAF50",
    "XGBoost":        "#FF9800",
    "Long-only":      "#E91E63",
}


# ---------------------------------------------------------------------------
# Core functions
# ---------------------------------------------------------------------------

def build_reference(actual: pd.Series) -> pd.Series:
    """
    Day-ahead reference price = same hour, previous day (shift 24h).
    This is the last comparable observation available at gate closure (noon D)
    for each hour h of D+1.
    """
    return actual.shift(24)


def compute_signals(predicted: pd.Series,
                    reference: pd.Series,
                    threshold: float = THRESHOLD) -> pd.Series:
    """
    signal(h) = sign(pred(h) - reference(h))
    Dead-band: 0 when |pred - ref| < threshold (no trade).
    """
    diff   = predicted - reference
    signal = np.sign(diff)
    signal[np.abs(diff) < threshold] = 0
    return signal.dropna()


def compute_pnl(actual: pd.Series,
                reference: pd.Series,
                signal: pd.Series,
                cost_per_mwh: float = COST_CENTRAL) -> tuple:
    """
    gross_pnl(h) = signal(h) * (actual(h) - reference(h))
                 = day-on-day price change in the direction of the signal.
    cost(h)      = cost_per_mwh * |signal(h)|  [per active hour]
    """
    idx        = signal.index
    day_delta  = (actual - reference).loc[idx]
    gross_pnl  = signal * day_delta
    cost_series = cost_per_mwh * signal.abs()
    net_pnl    = gross_pnl - cost_series
    return gross_pnl.dropna(), net_pnl.dropna()


def sharpe(pnl: pd.Series, ann: int = ANN_FACTOR) -> float:
    """Annualised Sharpe on daily P&L, sqrt(365), no risk-free rate."""
    daily = pnl.resample("D").sum()
    if daily.std() == 0 or len(daily) < 5:
        return 0.0
    return float(daily.mean() / daily.std() * np.sqrt(ann))


def max_dd(cum: pd.Series) -> float:
    return float((cum.cummax() - cum).max())


def calmar(pnl: pd.Series, cum: pd.Series, ann: int = ANN_FACTOR) -> float:
    ann_ret = pnl.resample("D").sum().mean() * ann
    mdd     = max_dd(cum)
    return float(ann_ret / mdd) if mdd > 0 else np.inf


def profit_factor(pnl: pd.Series) -> float:
    pos = pnl[pnl > 0].sum()
    neg = pnl[pnl < 0].abs().sum()
    return float(pos / neg) if neg > 0 else np.inf


def win_rate(pnl: pd.Series, signal: pd.Series) -> float:
    active = pnl.loc[signal.loc[pnl.index] != 0]
    return float((active > 0).mean() * 100) if len(active) > 0 else 0.0


def monthly_sharpe(net_pnl: pd.Series) -> pd.Series:
    """Sharpe computed independently on each calendar month."""
    results = {}
    for period, grp in net_pnl.groupby(net_pnl.index.to_period("M")):
        daily = grp.resample("D").sum()
        if daily.std() > 0 and len(daily) >= 5:
            results[str(period)] = daily.mean() / daily.std() * np.sqrt(ANN_FACTOR)
        else:
            results[str(period)] = np.nan
    return pd.Series(results)


def peak_offpeak_split(pnl: pd.Series, signal: pd.Series) -> dict:
    """Net P&L and win rate separately for peak (8-19) and off-peak hours."""
    peak_mask = pnl.index.hour.isin(PEAK_HOURS)
    out = {}
    for name, mask in [("peak", peak_mask), ("offpeak", ~peak_mask)]:
        p = pnl[mask]
        s = signal.loc[p.index]
        out[name] = {
            "net_pnl":  float(p.sum()),
            "win_rate": win_rate(p, s),
            "n_active": int((s != 0).sum()),
        }
    return out


def run_strategy(actual: pd.Series,
                 reference: pd.Series,
                 predicted: pd.Series,
                 label: str,
                 cost: float = COST_CENTRAL) -> dict:
    signal              = compute_signals(predicted, reference)
    gross_pnl, net_pnl = compute_pnl(actual, reference, signal, cost)
    cum_gross           = gross_pnl.cumsum()
    cum_net             = net_pnl.cumsum()

    total_cost    = cost * signal.abs().sum()
    n_active      = int((signal != 0).sum())
    cost_drag_pct = (total_cost / cum_gross.iloc[-1] * 100
                     if cum_gross.iloc[-1] != 0 else np.nan)

    return {
        "label":             label,
        "cost_eur_mwh":      cost,
        # Gross
        "gross_pnl":         float(cum_gross.iloc[-1]),
        "gross_sharpe":      sharpe(gross_pnl),
        # Net
        "net_pnl":           float(cum_net.iloc[-1]),
        "net_sharpe":        sharpe(net_pnl),
        "net_ann_pnl":       float(net_pnl.resample("D").sum().mean() * ANN_FACTOR),
        "max_drawdown":      max_dd(cum_net),
        "calmar":            calmar(net_pnl, cum_net),
        "win_rate":          win_rate(net_pnl, signal),
        "profit_factor":     profit_factor(net_pnl),
        "n_active_hours":    n_active,
        "active_pct":        float(n_active / len(signal) * 100),
        "n_flat_hours":      int((signal == 0).sum()),
        "cost_drag_pct":     cost_drag_pct,
        "monthly_sharpe":    monthly_sharpe(net_pnl),
        "peak_offpeak":      peak_offpeak_split(net_pnl, signal),
        # Series
        "_gross_pnl_s":      gross_pnl,
        "_net_pnl_s":        net_pnl,
        "_cum_gross":        cum_gross,
        "_cum_net":          cum_net,
        "_signal":           signal,
    }


# ---------------------------------------------------------------------------
# Plotting
# ---------------------------------------------------------------------------

def _ax_setup(ax, title, ylabel):
    ax.set_title(title, fontsize=12, fontweight="bold")
    ax.set_ylabel(ylabel)
    ax.grid(True, alpha=0.3)
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))


def plot_equity(strats: dict, net: bool, suffix: str):
    key   = "_cum_net" if net else "_cum_gross"
    cost_note = f" (frais : {COST_CENTRAL} EUR/MWh/h active)" if net else " (sans frais)"
    title = ("Performance cumulee NETTE" if net else "Performance cumulee BRUTE") + cost_note

    fig, ax = plt.subplots(figsize=(13, 6))
    for k, s in strats.items():
        c  = PALETTE.get(s["label"], "#333")
        lw = 2.5 if "RF avec" in s["label"] else 1.5
        ls = "--" if s["label"] == "Long-only" else "-"
        ax.plot(s[key].index, s[key].values,
                label=s["label"], color=c, linewidth=lw, linestyle=ls)
    ax.axhline(0, color="black", linewidth=0.8, linestyle=":")
    _ax_setup(ax, title, "P&L cumule (EUR/MW)")
    ax.set_xlabel("Date")
    ax.legend(loc="upper left", fontsize=10)
    plt.tight_layout()
    p = FIGS_DIR / f"equity_curves{suffix}.png"
    fig.savefig(p, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {p}")


def plot_monthly_pnl(strats: dict):
    keys = [k for k in strats if strats[k]["label"] in
            ("RF avec meteo", "Naif lag-168h", "XGBoost")]
    fig, ax = plt.subplots(figsize=(14, 5))
    bw      = 0.25
    months  = None
    for i, key in enumerate(keys):
        s  = strats[key]
        mp = s["_net_pnl_s"].resample("ME").sum()
        mp.index = mp.index.to_period("M").to_timestamp()
        if months is None:
            months = mp.index
        x = np.arange(len(mp))
        ax.bar(x + (i - len(keys)/2 + 0.5) * bw, mp.values, bw,
               label=s["label"], color=PALETTE.get(s["label"], "#888"), alpha=0.85)
    if months is not None:
        ax.set_xticks(np.arange(len(months)))
        ax.set_xticklabels([m.strftime("%b %Y") for m in months],
                           rotation=45, ha="right", fontsize=8)
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_title(f"P&L mensuel NET (EUR/MW) - scenario central ({COST_CENTRAL} EUR/MWh)",
                 fontsize=12, fontweight="bold")
    ax.set_ylabel("P&L (EUR/MW)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3, axis="y")
    plt.tight_layout()
    p = FIGS_DIR / "monthly_pnl.png"
    fig.savefig(p, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {p}")


def plot_pnl_distribution(strats: dict):
    labels = [s["label"] for s in strats.values()]
    data   = [s["_net_pnl_s"].values for s in strats.values()]
    colors = [PALETTE.get(l, "#888") for l in labels]
    fig, ax = plt.subplots(figsize=(11, 5))
    bp = ax.boxplot(data, patch_artist=True, tick_labels=labels,
                    whis=1.5, showfliers=False,
                    medianprops={"color": "black", "linewidth": 2})
    for patch, color in zip(bp["boxes"], colors):
        patch.set_facecolor(color); patch.set_alpha(0.7)
    ax.axhline(0, color="black", linewidth=0.8, linestyle=":")
    ax.set_title(f"Distribution P&L journalier NET (EUR/MW) - central ({COST_CENTRAL} EUR/MWh)",
                 fontsize=12, fontweight="bold")
    ax.set_ylabel("P&L (EUR/MW)")
    ax.grid(True, alpha=0.3, axis="y")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    p = FIGS_DIR / "pnl_distribution.png"
    fig.savefig(p, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {p}")


def plot_cost_sensitivity(sens: pd.DataFrame):
    models    = [m for m in sens["Modele"].unique() if "Long" not in m]
    scenarios = list(COST_SCENARIOS.keys())
    x         = np.arange(len(models))
    w         = 0.25
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    for ax, col, ylabel, title in [
        (axes[0], "P&L net (EUR/MW)",    "P&L net (EUR/MW)",    "P&L net par scenario de frais"),
        (axes[1], "Sharpe net (daily)",  "Sharpe annualise",     "Sharpe net par scenario de frais"),
    ]:
        for i, sc in enumerate(scenarios):
            sub  = sens[sens["Scenario"] == sc]
            vals = [sub[sub["Modele"] == m][col].values[0]
                    if len(sub[sub["Modele"] == m]) > 0 else 0
                    for m in models]
            ax.bar(x + (i-1)*w, vals, w,
                   label=f"{sc} ({COST_SCENARIOS[sc]} EUR/MWh)", alpha=0.85)
        ax.set_xticks(x)
        ax.set_xticklabels(models, rotation=15, ha="right", fontsize=9)
        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.set_ylabel(ylabel)
        ax.axhline(0, color="black", linewidth=0.7)
        ax.legend(fontsize=9)
        ax.grid(True, alpha=0.3, axis="y")
    fig.suptitle("Analyse de sensibilite aux couts de transaction",
                 fontsize=13, fontweight="bold")
    plt.tight_layout()
    p = FIGS_DIR / "cost_sensitivity.png"
    fig.savefig(p, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {p}")


def plot_rolling_sharpe(strats: dict, window_days: int = 30):
    """30-day rolling Sharpe on net P&L — shows temporal stability."""
    fig, ax = plt.subplots(figsize=(13, 5))
    for key, s in strats.items():
        if s["label"] == "Long-only":
            continue
        daily = s["_net_pnl_s"].resample("D").sum()
        roll  = daily.rolling(window_days, min_periods=15)
        rs    = roll.mean() / roll.std() * np.sqrt(ANN_FACTOR)
        ax.plot(rs.index, rs.values,
                label=s["label"],
                color=PALETTE.get(s["label"], "#333"),
                linewidth=1.6)
    ax.axhline(0, color="black", linewidth=0.8, linestyle=":")
    ax.axhline(2, color="grey", linewidth=0.8, linestyle="--", alpha=0.6)
    ax.text(ax.get_xlim()[0], 2.1, "Sharpe = 2 (ref. hedge fund)",
            fontsize=8, color="grey")
    ax.set_title(f"Sharpe glissant {window_days}j (NET, annualise sqrt(365))",
                 fontsize=12, fontweight="bold")
    ax.set_ylabel("Sharpe")
    ax.set_xlabel("Date")
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    p = FIGS_DIR / "rolling_sharpe.png"
    fig.savefig(p, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {p}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("=" * 65)
    print("FR Power Thesis -- Phase 3: Trading Backtest (day-ahead)")
    print("=" * 65)

    # 1. Load predictions
    print(f"\n[1/6] Loading {PREDS_PATH}...")
    if not PREDS_PATH.exists():
        raise FileNotFoundError(
            f"{PREDS_PATH} not found.\nRun scripts/run_full_pipeline.py first.")
    df = pd.read_parquet(PREDS_PATH)
    print(f"  {len(df):,} obs | {df.index.min().date()} -> {df.index.max().date()}")

    actual    = df["actual"]
    reference = build_reference(actual)   # same hour, day D (shift 24h)

    # Drop first 24 hours (reference = NaN)
    valid_idx = reference.dropna().index
    actual    = actual.loc[valid_idx]
    reference = reference.loc[valid_idx]
    df        = df.loc[valid_idx]

    print(f"  After 24h warm-up drop: {len(actual):,} obs "
          f"({actual.index.min().date()} -> {actual.index.max().date()})")
    print(f"\n  Signal design: sign(pred(h,D+1) - actual(h,D))  [day-ahead executable]")
    print(f"  P&L design   : signal * (actual(h,D+1) - actual(h,D))")
    print(f"  Dead-band    : {THRESHOLD} EUR/MWh")
    print(f"  Cost model   : {COST_CENTRAL} EUR/MWh per active hour (central scenario)")
    print(f"  Annualisation: sqrt({ANN_FACTOR}) (calendar days)")

    model_map = {
        "naive":         ("Naif lag-168h", df["naive"]),
        "rf_no_weather": ("RF sans meteo", df["rf_no_weather"]),
        "rf_weather":    ("RF avec meteo", df["rf_weather"]),
        "xgboost":       ("XGBoost",       df["xgboost"]),
    }

    # 2. Central scenario
    print(f"\n[2/6] Central scenario (cost = {COST_CENTRAL} EUR/MWh)...")
    strats = {}
    for key, (label, pred) in model_map.items():
        strats[key] = run_strategy(actual, reference, pred, label, COST_CENTRAL)

    # Long-only reference
    lo_pnl = (actual - reference).dropna()
    lo_cum  = lo_pnl.cumsum()
    lo_sig  = pd.Series(1.0, index=lo_pnl.index)
    strats["long_only"] = {
        "label": "Long-only", "cost_eur_mwh": 0.0,
        "gross_pnl": float(lo_cum.iloc[-1]), "gross_sharpe": sharpe(lo_pnl),
        "net_pnl":   float(lo_cum.iloc[-1]), "net_sharpe":   sharpe(lo_pnl),
        "net_ann_pnl": float(lo_pnl.resample("D").sum().mean() * ANN_FACTOR),
        "max_drawdown": max_dd(lo_cum), "calmar": calmar(lo_pnl, lo_cum),
        "win_rate": win_rate(lo_pnl, lo_sig),
        "profit_factor": profit_factor(lo_pnl),
        "n_active_hours": len(lo_pnl), "active_pct": 100.0,
        "n_flat_hours": 0, "cost_drag_pct": 0.0,
        "monthly_sharpe": monthly_sharpe(lo_pnl),
        "peak_offpeak": peak_offpeak_split(lo_pnl, lo_sig),
        "_gross_pnl_s": lo_pnl, "_net_pnl_s": lo_pnl,
        "_cum_gross": lo_cum,   "_cum_net":   lo_cum,
        "_signal": lo_sig,
    }

    # Print summary table
    print(f"\n  {'Modele':<22} {'GrossP&L':>9} {'NetP&L':>9} "
          f"{'Sharpe(net)':>12} {'MDD':>8} {'PF':>6} "
          f"{'Win%':>6} {'Active%':>8} {'CostDrag':>9}")
    print("  " + "-" * 96)
    for key in ["naive", "rf_no_weather", "rf_weather", "xgboost", "long_only"]:
        s = strats[key]
        print(f"  {s['label']:<22} "
              f"{s['gross_pnl']:>9,.0f} "
              f"{s['net_pnl']:>9,.0f} "
              f"{s['net_sharpe']:>12.3f} "
              f"{s['max_drawdown']:>8,.0f} "
              f"{s['profit_factor']:>6.2f} "
              f"{s['win_rate']:>5.1f}% "
              f"{s['active_pct']:>7.1f}% "
              f"{s['cost_drag_pct']:>8.1f}%")

    # Peak / off-peak
    print(f"\n  Peak vs Off-Peak NET P&L (central scenario):")
    print(f"  {'Modele':<22} {'Peak P&L':>10} {'Peak Win%':>10} "
          f"{'OffPk P&L':>10} {'OffPk Win%':>11}")
    print("  " + "-" * 66)
    for key in ["naive", "rf_no_weather", "rf_weather", "xgboost"]:
        s = strats[key]
        pk = s["peak_offpeak"]["peak"]
        op = s["peak_offpeak"]["offpeak"]
        print(f"  {s['label']:<22} "
              f"{pk['net_pnl']:>10,.0f} "
              f"{pk['win_rate']:>9.1f}% "
              f"{op['net_pnl']:>10,.0f} "
              f"{op['win_rate']:>10.1f}%")

    # Monthly Sharpe
    print(f"\n  Sharpe mensuel NET (scenario central) :")
    ms_df = pd.DataFrame({
        strats[k]["label"]: strats[k]["monthly_sharpe"]
        for k in ["naive", "rf_no_weather", "rf_weather", "xgboost"]
    }).round(2)
    print("  " + ms_df.to_string().replace("\n", "\n  "))

    # 3. Sensitivity analysis
    print(f"\n[3/6] Sensitivity analysis ({len(COST_SCENARIOS)} scenarios)...")
    sens_rows = []
    for sc, cost in COST_SCENARIOS.items():
        for key, (label, pred) in model_map.items():
            s = run_strategy(actual, reference, pred, label, cost)
            sens_rows.append({
                "Scenario":           sc,
                "Cout (EUR/MWh)":     cost,
                "Modele":             label,
                "P&L brut (EUR/MW)":  round(s["gross_pnl"], 0),
                "P&L net (EUR/MW)":   round(s["net_pnl"],   0),
                "Sharpe brut":        round(s["gross_sharpe"], 2),
                "Sharpe net (daily)": round(s["net_sharpe"],  2),
                "Max Drawdown":       round(s["max_drawdown"],0),
                "Calmar net":         round(s["calmar"], 2),
                "Win Rate (%)":       round(s["win_rate"], 1),
                "Profit Factor":      round(s["profit_factor"], 2),
                "Heures actives":     s["n_active_hours"],
                "Cost Drag (%)":      round(s["cost_drag_pct"], 1),
            })
    sens_df = pd.DataFrame(sens_rows)
    sens_df.to_csv(TABLES_DIR / "backtest_sensitivity.csv", index=False)
    print(f"  Saved: {TABLES_DIR / 'backtest_sensitivity.csv'}")

    pivot_sharpe = sens_df.pivot_table(
        index="Modele", columns="Scenario",
        values="Sharpe net (daily)", aggfunc="first"
    )[list(COST_SCENARIOS.keys())]
    print("\n  Sharpe NET par scenario :")
    print("  " + pivot_sharpe.round(2).to_string().replace("\n", "\n  "))

    # 4. Monthly Sharpe table
    ms_df.to_csv(TABLES_DIR / "backtest_monthly_sharpe.csv")
    print(f"\n  Saved: {TABLES_DIR / 'backtest_monthly_sharpe.csv'}")

    # 5. Save JSON + CSV
    print("\n[4/6] Saving outputs...")
    skip = {"_gross_pnl_s", "_net_pnl_s", "_cum_gross",
            "_cum_net", "_signal", "monthly_sharpe", "peak_offpeak"}
    json_out = {}
    for key, s in strats.items():
        json_out[key] = {k: (round(v, 4) if isinstance(v, float) else v)
                         for k, v in s.items() if k not in skip}
    with open(METRICS_DIR / "backtest_results.json", "w") as f:
        json.dump(json_out, f, indent=2)
    print(f"  Saved: {METRICS_DIR / 'backtest_results.json'}")

    col_map = {
        "label":          "Modele",
        "gross_pnl":      "P&L brut (EUR/MW)",
        "net_pnl":        "P&L net (EUR/MW)",
        "net_sharpe":     "Sharpe net (daily,sqrt365)",
        "max_drawdown":   "Max Drawdown (EUR/MW)",
        "calmar":         "Calmar",
        "win_rate":       "Win Rate (%)",
        "profit_factor":  "Profit Factor",
        "n_active_hours": "Heures actives",
        "active_pct":     "Actif (%)",
        "cost_drag_pct":  "Cost Drag (%)",
    }
    rows = []
    for key in ["naive", "rf_no_weather", "rf_weather", "xgboost", "long_only"]:
        s = strats[key]
        rows.append({v: (round(s[k], 2) if isinstance(s[k], float) else s[k])
                     for k, v in col_map.items()})
    comp_df = pd.DataFrame(rows)
    comp_df.to_csv(TABLES_DIR / "backtest_comparison.csv", index=False)
    print(f"  Saved: {TABLES_DIR / 'backtest_comparison.csv'}")
    print()
    print(comp_df.to_string(index=False))

    # 6. Plots
    print("\n[5/6] Generating plots...")
    plot_equity(strats, net=False, suffix="")
    plot_equity(strats, net=True,  suffix="_net")
    plot_monthly_pnl(strats)
    plot_pnl_distribution(strats)
    plot_cost_sensitivity(sens_df)
    plot_rolling_sharpe(strats)

    print("\n[6/6] Done. Summary:")
    print(f"  Signal    : sign(pred(h,D+1) - actual(h,D))  [day-ahead executable]")
    print(f"  P&L       : signal * (actual(h,D+1) - actual(h,D))")
    print(f"  Dead-band : {THRESHOLD} EUR/MWh")
    print(f"  Cost      : {COST_CENTRAL} EUR/MWh per active hour (central)")
    print(f"  Ann.      : sqrt({ANN_FACTOR}) on daily aggregation")
    print(f"\n  NOTE ON SHARPE: Sharpe > 2 reflects the strategy's directional")
    print(f"  consistency over this specific test period. It is NOT comparable")
    print(f"  to equity fund Sharpe ratios (different asset class, no leverage,")
    print(f"  different cost structure). The monthly Sharpe table shows whether")
    print(f"  performance is robust across months or driven by a few outliers.")
    print(f"\nAll outputs saved to outputs/")


if __name__ == "__main__":
    main()
