import numpy as np
import pandas as pd

# French nuclear installed capacity (MW) — source: RTE Bilan électrique 2023
FR_NUCLEAR_CAPACITY_MW = 63_000

# Clean spread parameters — CCGT gas, coal plant, source: RTE/ENTSO-E
HEAT_RATE_GAS = 2.0       # MWh_gas per MWh_elec (CCGT)
HEAT_RATE_COAL = 2.5      # MWh_coal per MWh_elec
EMISSION_FACTOR_GAS = 0.35  # tCO2 per MWh_elec (CCGT)
EMISSION_FACTOR_COAL = 0.95  # tCO2 per MWh_elec (coal)

# HDD threshold — 17°C is the standard French heating threshold (source: RTE)
HDD_THRESHOLD = 17.0
CDD_THRESHOLD = 22.0


def add_calendar_features(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["hour"] = df.index.hour
    df["dayofweek"] = df.index.dayofweek
    df["month"] = df.index.month
    df["is_weekend"] = (df.index.dayofweek >= 5).astype(int)
    df["hour_sin"] = np.sin(2 * np.pi * df["hour"] / 24)
    df["hour_cos"] = np.cos(2 * np.pi * df["hour"] / 24)
    df["month_sin"] = np.sin(2 * np.pi * df["month"] / 12)
    df["month_cos"] = np.cos(2 * np.pi * df["month"] / 12)
    return df


def add_price_lags(df: pd.DataFrame, target: str = "price_da_eur_mwh") -> pd.DataFrame:
    df = df.copy()
    for lag in [24, 48, 168]:
        df[f"{target}_lag{lag}h"] = df[target].shift(lag)
    df[f"{target}_roll24h_mean"] = df[target].shift(24).rolling(24).mean()
    df[f"{target}_roll168h_mean"] = df[target].shift(24).rolling(168).mean()
    return df


def add_weather_features(df: pd.DataFrame) -> pd.DataFrame:
    """Compute derived weather features from ERA5 columns.

    HDD threshold = 17°C — standard French heating threshold (RTE).
    This is higher than the German threshold (~15°C) reflecting the stronger
    thermosensitivity of French electric heating (~2.4 GW/°C in winter).
    """
    df = df.copy()
    if "temperature_2m" in df.columns:
        df["hdd"] = np.maximum(0, HDD_THRESHOLD - df["temperature_2m"])
        df["cdd"] = np.maximum(0, df["temperature_2m"] - CDD_THRESHOLD)
    if "wind_speed_10m" in df.columns:
        df["wind_power_proxy"] = df["wind_speed_10m"] ** 3
    return df


def add_weather_stress_index(df: pd.DataFrame) -> pd.DataFrame:
    """Composite index combining cold spells and low wind — high values signal
    peak demand with low renewable generation, a key driver of French price spikes."""
    df = df.copy()
    cold = df.get("hdd", pd.Series(0, index=df.index))
    wind = df.get("wind_speed_10m", pd.Series(1, index=df.index)).clip(lower=0.1)
    df["weather_stress_index"] = cold / wind
    return df


def add_nuclear_availability(df: pd.DataFrame) -> pd.DataFrame:
    """Nuclear availability ratio = actual generation / installed capacity.

    Proxy for the fraction of the French nuclear fleet online.
    France has ~63 GW installed (RTE 2023). A ratio below 0.7 signals
    significant outages and is associated with higher prices.
    """
    df = df.copy()
    if "gen_nuclear_mw" in df.columns:
        df["nuclear_avail_ratio"] = (
            df["gen_nuclear_mw"] / FR_NUCLEAR_CAPACITY_MW
        ).clip(0, 1)
    return df


def add_fuel_spreads(df: pd.DataFrame, target: str = "price_da_eur_mwh") -> pd.DataFrame:
    """Compute Clean Spark Spread and Clean Dark Spread when fuel price columns exist.

    CSS = Price_power − (TTF_gas × heat_rate_gas) − (EUA × emission_factor_gas)
    CDS = Price_power − (Coal × heat_rate_coal) − (EUA × emission_factor_coal)

    Columns expected (from fetch_fuels.py):
        ttf_eur_mwh  — TTF natural gas day-ahead price in EUR/MWh
        eua_eur_t    — EU ETS carbon price in EUR/tCO2
        coal_eur_t   — ARA coal price in EUR/tonne
    """
    df = df.copy()
    price = df.get(target)
    ttf = df.get("ttf_eur_mwh")
    eua = df.get("eua_eur_t")
    coal = df.get("coal_eur_t")

    if price is not None and ttf is not None and eua is not None:
        df["clean_spark_spread"] = (
            price - (ttf * HEAT_RATE_GAS) - (eua * EMISSION_FACTOR_GAS)
        )

    if price is not None and coal is not None and eua is not None:
        # coal_eur_t → coal_eur_mwh via heat_rate (1 tonne ≈ 6.978 MWh_th)
        coal_eur_mwh = coal / 6.978
        df["clean_dark_spread"] = (
            price - (coal_eur_mwh * HEAT_RATE_COAL) - (eua * EMISSION_FACTOR_COAL)
        )

    return df


# Columns produced by weather-related feature functions — used to split
# Modèle B (no weather) from Modèle C (with weather) in the pipeline.
WEATHER_FEATURE_COLS = [
    "temperature_2m",
    "wind_speed_10m",
    "solar_radiation",
    "precipitation",
    "hdd",
    "cdd",
    "wind_power_proxy",
    "weather_stress_index",
]


def build_feature_matrix(
    df: pd.DataFrame,
    target: str = "price_da_eur_mwh",
    include_weather: bool = True,
) -> tuple[pd.DataFrame, pd.Series]:
    """Build the full feature matrix.

    Args:
        df: Merged interim DataFrame (output of build_interim.py).
        target: Name of the price column to predict.
        include_weather: If False, ERA5/weather-derived columns are dropped
            (used to train Modèle B — RF sans météo).

    Returns:
        (X, y) ready for model training.
    """
    df = add_calendar_features(df)
    df = add_price_lags(df, target)
    df = add_weather_features(df)
    df = add_weather_stress_index(df)
    df = add_nuclear_availability(df)
    df = add_fuel_spreads(df, target)

    if not include_weather:
        drop_weather = [c for c in WEATHER_FEATURE_COLS if c in df.columns]
        df = df.drop(columns=drop_weather)

    # Drop columns that are more than 50% NaN (e.g. offshore wind in France)
    thresh = len(df) * 0.5
    df = df.dropna(axis=1, thresh=int(thresh))

    # Forward-fill remaining small gaps (hydro, generation, fuel prices)
    df = df.ffill(limit=3)

    # Only drop rows where target or core lag features are missing
    key_cols = [target, f"{target}_lag24h", f"{target}_lag168h"]
    key_cols = [c for c in key_cols if c in df.columns]
    df = df.dropna(subset=key_cols)

    X = df.drop(columns=[target])
    y = df[target]
    return X, y
