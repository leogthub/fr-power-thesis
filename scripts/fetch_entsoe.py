"""
Fetch French day-ahead prices, load forecast, generation mix,
cross-border flows, and nuclear availability from ENTSO-E.
Saves raw data to data/raw/ as parquet files.
"""
import ssl, urllib3
ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import pandas as pd
from entsoe import EntsoePandasClient
from src.config import ENTSOE_API_KEY, BIDDING_ZONE, RAW_DIR

RAW_DIR.mkdir(parents=True, exist_ok=True)

START = pd.Timestamp("2018-01-01", tz="Europe/Paris")
END   = pd.Timestamp("2025-04-30", tz="Europe/Paris")

client = EntsoePandasClient(api_key=ENTSOE_API_KEY)

# ---------------------------------------------------------------------------
# Neighbours for cross-border flows (ENTSO-E bidding zone codes)
# ---------------------------------------------------------------------------
NEIGHBOURS = {
    "DE": "10Y1001A1001A83F",   # Germany (DE-LU)
    "ES": "10YES-REE------0",   # Spain
    "IT": "10YIT-GRTN-----W",   # Italy — use two-letter code if zone fails
    "GB": "10YGB----------A",   # Great Britain
    "CH": "10YCH-SWISSGRID--D", # Switzerland — use two-letter code if zone fails
    "BE": "10YBE----------2",   # Belgium
}

# Fallback two-letter country codes for borders that fail with zone codes
NEIGHBOURS_CC = {
    "IT": "IT",
    "CH": "CH",
}


def fetch_and_save(name, fn, *args, **kwargs):
    print(f"Fetching {name}...")
    try:
        data = fn(*args, **kwargs)
        path = RAW_DIR / f"{name}.parquet"
        if isinstance(data, pd.Series):
            data.to_frame().to_parquet(path)
        else:
            data.to_parquet(path)
        print(f"  Saved {len(data)} rows -> {path}")
        return data
    except Exception as e:
        print(f"  ERROR fetching {name}: {e}")
        return None


# ---------------------------------------------------------------------------
# Core ENTSO-E data (unchanged from Phase 1, extended to 2025-04-30)
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Cross-border physical flows (net exports FR → neighbour)
# ---------------------------------------------------------------------------
print("\nFetching cross-border flows...")
flow_frames = {}
for country, zone in NEIGHBOURS.items():
    # Try zone code first, then two-letter CC fallback
    codes_to_try = [zone]
    if country in NEIGHBOURS_CC:
        codes_to_try.append(NEIGHBOURS_CC[country])
    success = False
    for code in codes_to_try:
        try:
            exports = client.query_crossborder_flows(BIDDING_ZONE, code, start=START, end=END)
            imports = client.query_crossborder_flows(code, BIDDING_ZONE, start=START, end=END)
            net = exports - imports
            net.name = f"flow_net_fr_{country.lower()}_mw"
            flow_frames[country] = net
            print(f"  FR-{country}: {len(net)} rows (code: {code})")
            success = True
            break
        except Exception as e:
            pass
    if not success:
        print(f"  WARNING FR-{country}: could not fetch with any code")

if flow_frames:
    df_flows = pd.concat(flow_frames.values(), axis=1)
    df_flows = df_flows.resample("h").mean()
    df_flows.to_parquet(RAW_DIR / "crossborder_flows.parquet")
    print(f"  Saved crossborder_flows.parquet ({df_flows.shape})")

# ---------------------------------------------------------------------------
# Nuclear unavailability / availability proxy
# Using installed generation capacity (aggregated) to get declared capacity,
# then computing availability ratio vs actual generation.
# ---------------------------------------------------------------------------
print("\nFetching nuclear installed capacity (aggregated)...")
try:
    nuclear_capacity = client.query_installed_generation_capacity_aggregated(
        BIDDING_ZONE, start=START, end=END, psr_type="B14"  # B14 = Nuclear
    )
    # Result is typically annual/semi-annual — forward-fill to hourly
    if isinstance(nuclear_capacity, pd.DataFrame):
        nuclear_capacity = nuclear_capacity.iloc[:, 0]
    nuclear_capacity.name = "nuclear_installed_capacity_mw"
    nuclear_capacity = nuclear_capacity.resample("h").ffill()
    nuclear_capacity.to_frame().to_parquet(RAW_DIR / "nuclear_capacity.parquet")
    print(f"  Saved nuclear_capacity.parquet ({len(nuclear_capacity)} rows)")
except Exception as e:
    print(f"  WARNING nuclear capacity: {e}")
    print("  Will use fixed 63,000 MW as fallback in features.py")

print("\nENTSO-E download complete.")
