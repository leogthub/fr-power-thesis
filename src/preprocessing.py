import pandas as pd
from entsoe import EntsoePandasClient
from src.config import ENTSOE_API_KEY, BIDDING_ZONE, TIMEZONE, RAW_DIR


def fetch_day_ahead_prices(start: pd.Timestamp, end: pd.Timestamp) -> pd.Series:
    client = EntsoePandasClient(api_key=ENTSOE_API_KEY)
    prices = client.query_day_ahead_prices(BIDDING_ZONE, start=start, end=end)
    prices.name = "price_da_eur_mwh"
    return prices


def fetch_load_forecast(start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    client = EntsoePandasClient(api_key=ENTSOE_API_KEY)
    return client.query_load_forecast(BIDDING_ZONE, start=start, end=end)


def fetch_generation_forecast(start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    client = EntsoePandasClient(api_key=ENTSOE_API_KEY)
    return client.query_generation_forecast(BIDDING_ZONE, start=start, end=end)


def load_era5_weather(path=None) -> pd.DataFrame:
    """Load ERA5 weather data from a parquet file downloaded via cdsapi."""
    if path is None:
        path = RAW_DIR / "era5_france.parquet"
    return pd.read_parquet(path)


def clean_prices(prices: pd.Series) -> pd.Series:
    prices = prices.resample("h").mean()
    prices = prices.clip(lower=-500, upper=3000)
    prices = prices.interpolate(method="time", limit=3)
    return prices


def merge_features(prices: pd.Series, load: pd.DataFrame, weather: pd.DataFrame) -> pd.DataFrame:
    df = prices.to_frame()
    df = df.join(load.resample("h").mean(), how="left")
    df = df.join(weather.resample("h").mean(), how="left")
    df.index = df.index.tz_convert(TIMEZONE)
    return df.sort_index()
