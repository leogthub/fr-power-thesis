"""
Fetch French day-ahead prices, load forecast, and generation mix from ENTSO-E.
Saves raw data to data/raw/ as parquet files.
"""
import pandas as pd
from entsoe import EntsoePandasClient
from src.config import ENTSOE_API_KEY, BIDDING_ZONE, RAW_DIR

RAW_DIR.mkdir(parents=True, exist_ok=True)

START = pd.Timestamp("2018-01-01", tz="Europe/Paris")
END = pd.Timestamp("2024-12-31", tz="Europe/Paris")

client = EntsoePandasClient(api_key=ENTSOE_API_KEY)


def fetch_and_save(name, fn, *args, **kwargs):
    print(f"Fetching {name}...")
    try:
        data = fn(*args, **kwargs)
        path = RAW_DIR / f"{name}.parquet"
        data.to_frame().to_parquet(path) if isinstance(data, pd.Series) else data.to_parquet(path)
        print(f"  Saved {len(data)} rows -> {path}")
    except Exception as e:
        print(f"  ERROR: {e}")


fetch_and_save(
    "prices_da",
    client.query_day_ahead_prices,
    BIDDING_ZONE, start=START, end=END,
)

fetch_and_save(
    "load_forecast",
    client.query_load_forecast,
    BIDDING_ZONE, start=START, end=END,
)

fetch_and_save(
    "generation_forecast",
    client.query_generation_forecast,
    BIDDING_ZONE, start=START, end=END,
)

fetch_and_save(
    "generation_actual",
    client.query_generation,
    BIDDING_ZONE, start=START, end=END,
)

print("ENTSO-E download complete.")
