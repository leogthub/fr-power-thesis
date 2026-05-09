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


def naive_forecast(y: pd.Series, lag: int = 168) -> pd.Series:
    """Same-hour previous week benchmark (lag-168h).

    The roadmap defines the naive benchmark as price(h) = price(h-168),
    i.e. the same hour from the previous week. This is a stronger baseline
    than lag-24h because electricity prices exhibit strong weekly seasonality.
    """
    return y.shift(lag)
