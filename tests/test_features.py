import pytest
import pandas as pd
import numpy as np
from src.features import (
    add_calendar_features,
    add_price_lags,
    add_weather_features,
    add_weather_stress_index,
)

FREQ = "h"
IDX = pd.date_range("2023-01-01", periods=200, freq=FREQ, tz="Europe/Paris")


def base_df():
    return pd.DataFrame(
        {
            "price_da_eur_mwh": np.random.uniform(50, 200, len(IDX)),
            "temperature_2m": np.random.uniform(-5, 35, len(IDX)),
            "wind_speed_10m": np.random.uniform(0, 20, len(IDX)),
        },
        index=IDX,
    )


def test_calendar_features_columns():
    df = add_calendar_features(base_df())
    for col in ["hour", "dayofweek", "month", "is_weekend", "hour_sin", "hour_cos"]:
        assert col in df.columns


def test_calendar_hour_range():
    df = add_calendar_features(base_df())
    assert df["hour"].between(0, 23).all()


def test_price_lags_created():
    df = add_price_lags(base_df())
    assert "price_da_eur_mwh_lag24h" in df.columns
    assert "price_da_eur_mwh_lag168h" in df.columns


def test_weather_hdd_non_negative():
    df = add_weather_features(base_df())
    assert (df["hdd"] >= 0).all()


def test_weather_stress_index_positive():
    df = add_weather_features(base_df())
    df = add_weather_stress_index(df)
    assert (df["weather_stress_index"] >= 0).all()


def test_is_weekend_binary():
    df = add_calendar_features(base_df())
    assert set(df["is_weekend"].unique()).issubset({0, 1})
