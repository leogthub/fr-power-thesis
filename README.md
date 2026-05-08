# Forecasting Day-Ahead French Power Prices using Weather-Based Random Forest Models

**EDHEC Business School — MSc Data Analysis & Artificial Intelligence**  
Master's Thesis

---

## Overview

This project develops a machine learning framework to forecast day-ahead electricity prices on the French power market (EPEX SPOT France). The core approach combines weather-based features (temperature, wind speed, solar irradiance, precipitation) with fundamental power market drivers (load forecasts, cross-border flows, generation mix, fuel prices) to train Random Forest and gradient-boosted models.

The methodology is inspired by the literature on weather-driven electricity price forecasting (see Ziel & Weron 2018, Lago et al. 2021) and adapted to the specificities of the French market, which is heavily influenced by nuclear availability and seasonal heating demand.

---

## Project Structure

```
fr-power-thesis/
├── data/
│   ├── raw/            # Original unmodified data from ENTSO-E, ERA5, etc.
│   ├── interim/        # Intermediate cleaned/merged data
│   └── processed/      # Final feature-engineered datasets ready for modelling
├── notebooks/          # Exploratory and reporting notebooks
├── src/
│   ├── __init__.py
│   ├── config.py       # Paths, constants, model hyperparameters
│   ├── preprocessing.py# Data loading and cleaning routines
│   ├── features.py     # Feature engineering (lags, weather indices, calendar)
│   ├── models.py       # Model training (Random Forest, XGBoost, baselines)
│   ├── evaluate.py     # Metrics: MAE, RMSE, MAPE, pinball loss
│   ├── backtest.py     # Walk-forward cross-validation and backtesting
│   └── plots.py        # Visualisation helpers
├── outputs/
│   ├── figures/        # Generated charts and plots
│   ├── tables/         # LaTeX / CSV result tables
│   └── metrics/        # Saved evaluation metrics (JSON/CSV)
├── scripts/
│   └── run_full_pipeline.py  # End-to-end pipeline runner
├── tests/
│   └── test_features.py      # Unit tests for feature engineering
├── .env.example        # API key template (copy to .env)
├── requirements.txt    # Python dependencies
└── README.md
```

---

## Data Sources

| Source | API / Tool | Content |
|--------|-----------|---------|
| ENTSO-E Transparency Platform | `entsoe-py` | Day-ahead prices, load, generation by type, cross-border flows |
| ERA5 (Copernicus Climate) | `cdsapi` | Temperature, wind speed, solar radiation, precipitation |
| EEX / ICE | Manual / requests | Gas (TTF), coal, carbon (EUA) futures |

---

## Quickstart

### 1. Clone and install

```bash
git clone https://github.com/leogthub/fr-power-thesis.git
cd fr-power-thesis
pip install -r requirements.txt
```

### 2. Configure API keys

```bash
cp .env.example .env
# Edit .env and fill in your ENTSOE_API_KEY and CDS_API_KEY
```

### 3. Run the full pipeline

```bash
python scripts/run_full_pipeline.py
```

### 4. Explore notebooks

```bash
jupyter notebook notebooks/
```

---

## Models

| Model | Description |
|-------|-------------|
| Naive | Yesterday's price (benchmark) |
| ARIMA | Time-series baseline |
| Random Forest | Main model — weather + fundamental features |
| XGBoost | Gradient-boosted tree alternative |
| Random Forest + Weather Index | Extended model with composite weather stress index |

---

## Evaluation

Models are evaluated on an out-of-sample test set (last 12 months) using a walk-forward expanding window. Metrics reported:

- **MAE** — Mean Absolute Error (€/MWh)
- **RMSE** — Root Mean Square Error
- **MAPE** — Mean Absolute Percentage Error
- **R²** — Coefficient of determination

---

## References

- Ziel, F., & Weron, R. (2018). Day-ahead electricity price forecasting with high-dimensional structures. *Energy Economics*.
- Lago, J., Marcjasz, G., De Schutter, B., & Weron, R. (2021). Forecasting day-ahead electricity prices. *Applied Energy*.
- ENTSO-E (2024). Transparency Platform. https://transparency.entsoe.eu
- Copernicus Climate Change Service (2024). ERA5 hourly data. https://cds.climate.copernicus.eu

---

## Author

Leo Camberleng — EDHEC MSc Data Analysis & AI  
Contact: leo.cmbrlng@gmail.com
