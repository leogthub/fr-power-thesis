import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from src.config import RF_PARAMS, XGB_PARAMS


def train_random_forest(X_train: pd.DataFrame, y_train: pd.Series) -> RandomForestRegressor:
    model = RandomForestRegressor(**RF_PARAMS)
    model.fit(X_train, y_train)
    return model


def train_xgboost(X_train: pd.DataFrame, y_train: pd.Series) -> XGBRegressor:
    model = XGBRegressor(**XGB_PARAMS)
    model.fit(X_train, y_train, verbose=False)
    return model


def naive_forecast(y: pd.Series, lag: int = 24) -> pd.Series:
    """Same-hour previous day benchmark."""
    return y.shift(lag)
