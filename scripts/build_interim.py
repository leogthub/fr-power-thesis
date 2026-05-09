"""
Merge all raw data sources into a single hourly DataFrame.
Output: data/interim/merged.parquet

Sources merged:
  - prices_da.parquet          : Day-ahead prices (ENTSO-E)
  - load_forecast.parquet      : Load forecast (ENTSO-E)
  - generation_actual.parquet  : Production by fuel type (ENTSO-E)
  - generation_forecast.parquet: Renewable generation forecast (ENTSO-E)
  - era5_france.parquet        : Weather data (ERA5 / Copernicus)
  - crossborder_flows.parquet  : Net cross-border flows FR→neighbours (ENTSO-E)
  - nuclear_capacity.parquet   : Declared nuclear installed capacity (ENTSO-E)
  - fuel_prices.parquet        : TTF gas, EUA CO2, ARA coal (free sources)
"""
import sys
sys.path.insert(0, ".")

import pandas as pd
import numpy as np
from src.config import RAW_DIR, INTERIM_DIR, TIMEZONE

INTERIM_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_prices() -> pd.Series:
    df = pd.read_parquet(RAW_DIR / "prices_da.parquet")
    s = df.iloc[:, 0]
    s.name = "price_da_eur_mwh"
    s = s.resample("h").mean()
    s = s.clip(lower=-500, upper=3000)
    s = s.interpolate(method="time", limit=3)
    return s


def load_load() -> pd.Series:
    df = pd.read_parquet(RAW_DIR / "load_forecast.parquet")
    s = df.iloc[:, 0]
    s.name = "load_forecast_mw"
    return s.resample("h").mean()


def load_generation() -> pd.DataFrame:
    df = pd.read_parquet(RAW_DIR / "generation_actual.parquet")
    if isinstance(df.columns, pd.MultiIndex):
        df = df.xs("Actual Aggregated", level=1, axis=1, drop_level=True)
    df = df.resample("h").mean()
    rename = {}
    for col in df.columns:
        col_lower = str(col).lower()
        if "nuclear" in col_lower:
            rename[col] = "gen_nuclear_mw"
        elif "wind onshore" in col_lower:
            rename[col] = "gen_wind_onshore_mw"
        elif "wind offshore" in col_lower:
            rename[col] = "gen_wind_offshore_mw"
        elif "solar" in col_lower:
            rename[col] = "gen_solar_mw"
        elif "hydro" in col_lower and "run" in col_lower:
            rename[col] = "gen_hydro_ror_mw"
        elif "hydro" in col_lower and "reservoir" in col_lower:
            rename[col] = "gen_hydro_reservoir_mw"
        elif "gas" in col_lower or "fossil gas" in col_lower:
            rename[col] = "gen_gas_mw"
    df = df.rename(columns=rename)
    keep = [c for c in df.columns if c.startswith("gen_")]
    return df[keep]


def load_generation_forecast() -> pd.DataFrame:
    """Renewable generation forecast — forward-looking signal for traders."""
    path = RAW_DIR / "generation_forecast.parquet"
    if not path.exists():
        print("  WARNING: generation_forecast.parquet not found, skipping.")
        return pd.DataFrame()
    df = pd.read_parquet(path)
    if isinstance(df.columns, pd.MultiIndex):
        # Keep only the forecast column (level=1)
        try:
            df = df.xs("Forecasted Generation", level=1, axis=1, drop_level=True)
        except KeyError:
            df = df.iloc[:, ::2]  # fallback: take every other column
    df = df.resample("h").mean()
    rename = {}
    for col in df.columns:
        col_lower = str(col).lower()
        if "wind" in col_lower:
            rename[col] = "forecast_wind_mw"
        elif "solar" in col_lower:
            rename[col] = "forecast_solar_mw"
    df = df.rename(columns=rename)
    keep = [c for c in df.columns if c.startswith("forecast_")]
    return df[keep] if keep else pd.DataFrame()


def load_weather() -> pd.DataFrame:
    path = RAW_DIR / "era5_france.parquet"
    if not path.exists():
        print("  WARNING: era5_france.parquet not found, skipping weather data.")
        return pd.DataFrame()
    df = pd.read_parquet(path)
    if df.index.tz is None:
        df.index = df.index.tz_localize("UTC").tz_convert(TIMEZONE)
    else:
        df.index = df.index.tz_convert(TIMEZONE)
    drop_cols = [c for c in df.columns if c in ("number", "expver") or df[c].dtype == object]
    df = df.drop(columns=drop_cols, errors="ignore")
    return df.resample("h").mean()


def load_crossborder_flows() -> pd.DataFrame:
    path = RAW_DIR / "crossborder_flows.parquet"
    if not path.exists():
        print("  WARNING: crossborder_flows.parquet not found, skipping.")
        return pd.DataFrame()
    df = pd.read_parquet(path)
    if df.index.tz is None:
        df.index = df.index.tz_localize("UTC").tz_convert(TIMEZONE)
    else:
        df.index = df.index.tz_convert(TIMEZONE)
    return df.resample("h").mean()


def load_nuclear_capacity() -> pd.Series:
    """Declared nuclear installed capacity — used to refine availability ratio in features.py."""
    path = RAW_DIR / "nuclear_capacity.parquet"
    if not path.exists():
        print("  INFO: nuclear_capacity.parquet not found (will use fixed 63 GW in features.py).")
        return pd.Series(dtype=float)
    df = pd.read_parquet(path)
    s = df.iloc[:, 0]
    s.name = "nuclear_installed_capacity_mw"
    if s.index.tz is None:
        s.index = s.index.tz_localize("UTC").tz_convert(TIMEZONE)
    else:
        s.index = s.index.tz_convert(TIMEZONE)
    return s.resample("h").ffill()


def load_fuels() -> pd.DataFrame:
    """Daily fuel prices forward-filled to hourly."""
    path = RAW_DIR / "fuel_prices.parquet"
    if not path.exists():
        print("  WARNING: fuel_prices.parquet not found. Run scripts/fetch_fuels.py first.")
        return pd.DataFrame()
    df = pd.read_parquet(path)
    if df.index.tz is None:
        df.index = df.index.tz_localize("UTC").tz_convert(TIMEZONE)
    else:
        df.index = df.index.tz_convert(TIMEZONE)
    # Upsample daily → hourly with forward fill (prices are constant within a day)
    df = df.resample("h").ffill()
    return df


def align_tz(obj):
    """Ensure a Series or DataFrame index is in Europe/Paris timezone."""
    if hasattr(obj, "index"):
        if obj.index.tz is None:
            obj.index = obj.index.tz_localize("UTC")
        obj.index = obj.index.tz_convert(TIMEZONE)
    return obj


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
print("Loading prices...")
prices = align_tz(load_prices())

print("Loading load forecast...")
load = align_tz(load_load())

print("Loading generation mix (actual)...")
gen = align_tz(load_generation())

print("Loading generation forecast (renewables)...")
gen_fc = load_generation_forecast()
if not gen_fc.empty:
    gen_fc = align_tz(gen_fc)

print("Loading weather data (ERA5)...")
weather = load_weather()

print("Loading cross-border flows...")
flows = load_crossborder_flows()

print("Loading nuclear installed capacity...")
nuc_cap = load_nuclear_capacity()

print("Loading fuel prices...")
fuels = load_fuels()

# ---------------------------------------------------------------------------
# Merge everything on the price index (left join)
# ---------------------------------------------------------------------------
print("\nMerging all sources...")
df = prices.to_frame()
df = df.join(load, how="left")
df = df.join(gen, how="left")

if not gen_fc.empty:
    df = df.join(gen_fc, how="left")

if not weather.empty:
    df = df.join(weather, how="left")

if not flows.empty:
    df = df.join(flows, how="left")

if len(nuc_cap) > 0:
    df = df.join(nuc_cap.to_frame(), how="left")

if not fuels.empty:
    df = df.join(fuels, how="left")

df = df.sort_index()
df = df.dropna(subset=["price_da_eur_mwh"])

out = INTERIM_DIR / "merged.parquet"
df.to_parquet(out)

print(f"\nSaved -> {out}")
print(f"  Shape  : {df.shape[0]:,} rows x {df.shape[1]} columns")
print(f"  Period : {df.index.min()} -> {df.index.max()}")
print(f"  Columns: {list(df.columns)}")
print("\nMissing values per column:")
print(df.isnull().sum()[df.isnull().sum() > 0])
print("\nDescriptive statistics:")
print(df.describe().round(2))
