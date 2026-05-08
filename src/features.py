import numpy as np
import pandas as pd


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
    """Compute derived weather features from ERA5 columns."""
    df = df.copy()
    if "temperature_2m" in df.columns:
        df["hdd"] = np.maximum(0, 15.5 - df["temperature_2m"])
        df["cdd"] = np.maximum(0, df["temperature_2m"] - 22)
    if "wind_speed_10m" in df.columns:
        df["wind_power_proxy"] = df["wind_speed_10m"] ** 3
    return df


def add_weather_stress_index(df: pd.DataFrame) -> pd.DataFrame:
    """Composite index combining cold spells and low wind periods."""
    df = df.copy()
    cold = df.get("hdd", pd.Series(0, index=df.index))
    wind = df.get("wind_speed_10m", pd.Series(1, index=df.index)).clip(lower=0.1)
    df["weather_stress_index"] = cold / wind
    return df


def build_feature_matrix(df: pd.DataFrame, target: str = "price_da_eur_mwh") -> tuple[pd.DataFrame, pd.Series]:
    df = add_calendar_features(df)
    df = add_price_lags(df, target)
    df = add_weather_features(df)
    df = add_weather_stress_index(df)
    df = df.dropna()
    X = df.drop(columns=[target])
    y = df[target]
    return X, y
