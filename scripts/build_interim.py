"""
Merge raw ENTSO-E and ERA5 data into a single hourly DataFrame.
Output: data/interim/merged.parquet
"""
import sys
sys.path.insert(0, ".")

import pandas as pd
import numpy as np
from src.config import RAW_DIR, INTERIM_DIR, TIMEZONE

INTERIM_DIR.mkdir(parents=True, exist_ok=True)


def load_prices() -> pd.Series:
    df = pd.read_parquet(RAW_DIR / "prices_da.parquet")
    s = df.iloc[:, 0]
    s.name = "price_da_eur_mwh"
    s = s.resample("h").mean()
    # clip extreme spikes (keep EPEX valid range)
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
    # Flatten MultiIndex columns: keep only "Actual Aggregated" level
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
    # Drop ERA5 metadata columns
    drop_cols = [c for c in df.columns if c in ("number", "expver") or df[c].dtype == object]
    df = df.drop(columns=drop_cols, errors="ignore")
    return df.resample("h").mean()


def align_timezone(s):
    if hasattr(s, "index"):
        if s.index.tz is None:
            s.index = s.index.tz_localize("UTC")
        s.index = s.index.tz_convert(TIMEZONE)
    return s


print("Loading prices...")
prices = align_timezone(load_prices())

print("Loading load forecast...")
load = align_timezone(load_load())

print("Loading generation mix...")
gen = align_timezone(load_generation())

print("Loading weather data...")
weather = load_weather()

print("Merging...")
df = prices.to_frame()
df = df.join(load, how="left")
df = df.join(gen, how="left")
if not weather.empty:
    df = df.join(weather, how="left")

df = df.sort_index()

# Drop rows with no price at all
df = df.dropna(subset=["price_da_eur_mwh"])

out = INTERIM_DIR / "merged.parquet"
df.to_parquet(out)
print(f"Saved -> {out} ({len(df):,} rows x {df.shape[1]} columns)")
print(df.describe().round(2))
