"""
Fetch fuel price data (TTF natural gas, EUA CO2, ARA coal) from free sources.
Saves to data/raw/fuel_prices.parquet with daily granularity,
forward-filled to hourly in build_interim.py.

Sources (all free / open):
  TTF gas  : Yahoo Finance  (ticker TTF=F  — front-month futures, EUR/MWh)
  EUA CO2  : Ember Climate  (open CSV — monthly update)
  Coal ARA : Yahoo Finance  (ticker MTF=F  — ARA coal front-month, USD/t → EUR/t)

Fallback strategy:
  If yfinance fails for any ticker, the column is skipped with a warning.
  Clean Spreads in features.py are computed only when all three columns exist.
"""
import ssl
import urllib3
import requests
import pandas as pd
import numpy as np
from src.config import RAW_DIR

# ---------------------------------------------------------------------------
# SSL workaround — corporate/Windows environments often block cert verification
# ---------------------------------------------------------------------------
ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_SESSION = requests.Session()
_SESSION.verify = False

RAW_DIR.mkdir(parents=True, exist_ok=True)

START = "2018-01-01"
END   = "2025-04-30"


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------
def _yf_daily(ticker: str, start: str, end: str, col_name: str) -> pd.Series | None:
    """Download a daily close series from Yahoo Finance via yfinance (curl_cffi backend)."""
    try:
        import yfinance as yf
        raw = yf.download(
            ticker, start=start, end=end,
            auto_adjust=True, progress=False,
        )
        if raw.empty:
            print(f"  WARNING: yfinance returned empty data for {ticker}")
            return None
        s = raw["Close"].squeeze()
        s.name = col_name
        s.index = pd.to_datetime(s.index).tz_localize("UTC")
        print(f"  {ticker} ({col_name}): {len(s)} daily rows")
        return s
    except Exception as e:
        print(f"  WARNING: yfinance failed for {ticker}: {e}")
        return None


def _ember_eua(start: str, end: str) -> pd.Series | None:
    """Download EU ETS (EUA) carbon price from Ember Climate open data.

    Ember publishes a daily EUA price CSV updated monthly.
    URL: https://ember-climate.org/data/data-tools/carbon-price-viewer/
    Fallback direct link to the underlying CSV.
    """
    urls = [
        # Ember direct CSV
        "https://ember-climate.org/app/uploads/2022/01/EU-ETS-carbon-prices.csv",
        # Backup: Sandbag / Carbon Monitor aggregation
        "https://raw.githubusercontent.com/openclimatedata/eu-ets-carbon-price/main/data/eu-ets-carbon-price.csv",
    ]
    for url in urls:
        try:
            resp = _SESSION.get(url, timeout=30)
            resp.raise_for_status()
            from io import StringIO
            df = pd.read_csv(StringIO(resp.text))
            # Detect date and price columns flexibly
            date_col = next((c for c in df.columns if "date" in c.lower()), df.columns[0])
            price_col = next(
                (c for c in df.columns if any(x in c.lower() for x in ["price", "eur", "eua", "carbon"])),
                df.columns[-1],
            )
            s = pd.to_numeric(df[price_col], errors="coerce")
            s.index = pd.to_datetime(df[date_col], errors="coerce")
            s = s.dropna()
            s.index = s.index.tz_localize("UTC")
            s.name = "eua_eur_t"
            s = s.sort_index().loc[start:end]
            if len(s) > 100:
                print(f"  EUA (Ember): {len(s)} daily rows from {url}")
                return s
        except Exception as e:
            print(f"  WARNING Ember URL {url}: {e}")

    # Final fallback: yfinance EUA futures
    print("  Falling back to yfinance for EUA (EMISS.L or CO2.DE)...")
    for ticker in ["CO2.DE", "EMISS.L", "EUA=F"]:
        s = _yf_daily(ticker, start, end, "eua_eur_t")
        if s is not None and len(s) > 100:
            return s

    return None


# ---------------------------------------------------------------------------
# Alternative: FRED API (St. Louis Fed) — uses requests, no curl_cffi
# ---------------------------------------------------------------------------
FRED_BASE = "https://fred.stlouisfed.org/graph/fredgraph.csv"

def _fred_series(series_id: str, col_name: str, start: str, end: str) -> pd.Series | None:
    """Fetch a monthly series from FRED (no API key required for CSV).
    FRED CSV format: observation_date,<series_id>
    """
    try:
        url = f"{FRED_BASE}?id={series_id}"
        resp = _SESSION.get(url, timeout=30)
        resp.raise_for_status()
        from io import StringIO
        df = pd.read_csv(StringIO(resp.text))
        # FRED uses 'observation_date' as the date column
        date_col = df.columns[0]
        val_col  = df.columns[1]
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.dropna(subset=[date_col])
        df = df.set_index(date_col)
        s = pd.to_numeric(df[val_col], errors="coerce").dropna()
        s.index = s.index.tz_localize("UTC")
        s = s.loc[start:end]
        s.name = col_name
        if len(s) > 10:
            print(f"  FRED {series_id} ({col_name}): {len(s)} rows")
            return s
        return None
    except Exception as e:
        print(f"  WARNING FRED {series_id}: {e}")
        return None


# ---------------------------------------------------------------------------
# Main fetch
# ---------------------------------------------------------------------------
print("Fetching fuel prices...")

frames = []

# TTF gas — try yfinance first, fallback to FRED EU gas price index
ttf = _yf_daily("TTF=F", START, END, "ttf_eur_mwh")
if ttf is None:
    # FRED: PNGASEUUSDM = EU natural gas import price (USD/MMBtu) → convert to EUR/MWh
    # 1 MMBtu ≈ 0.2931 MWh → divide by 0.2931 to get USD/MWh → multiply by 0.92 EUR/USD
    ttf_usd = _fred_series("PNGASEUUSDM", "ttf_eur_mwh", START, END)
    if ttf_usd is not None:
        ttf = (ttf_usd / 0.2931 * 0.92).resample("D").ffill()
        ttf.name = "ttf_eur_mwh"
        print(f"  TTF (FRED proxy, EUR/MWh): {len(ttf)} rows")
if ttf is not None:
    frames.append(ttf)

# EUA CO2 — try Ember then yfinance
eua = _ember_eua(START, END)
if eua is None:
    # FRED doesn't have EUA — use a fixed rolling estimate based on known ranges
    # EUA ranged ~15-30 EUR/t (2018-2020), ~25-50 (2021), ~60-100 (2022-2023), ~50-70 (2024)
    # This is a rough proxy — acknowledge limitation in thesis
    print("  INFO: EUA not available from automated sources. Skipping clean spreads.")
if eua is not None:
    frames.append(eua)

# ARA coal — try yfinance, fallback to FRED coal price
coal_usd = _yf_daily("MTF=F", START, END, "coal_eur_t")
if coal_usd is None:
    # FRED: PCOALAUUSDM = Australian coal price (USD/metric ton) — closest proxy to ARA
    coal_usd = _fred_series("PCOALAUUSDM", "coal_eur_t", START, END)
    if coal_usd is not None:
        coal_usd = (coal_usd * 0.92).resample("D").ffill()
        coal_usd.name = "coal_eur_t"
        print(f"  Coal (FRED proxy, EUR/t): {len(coal_usd)} rows")
if coal_usd is not None:
    frames.append(coal_usd)

# ---------------------------------------------------------------------------
# Merge & save
# ---------------------------------------------------------------------------
if frames:
    df_fuels = pd.concat(frames, axis=1).sort_index()
    # Keep only the requested date range
    df_fuels = df_fuels.loc[START:END]
    # Ensure daily index (some tickers have duplicate timestamps)
    df_fuels = df_fuels[~df_fuels.index.duplicated(keep="last")]
    df_fuels = df_fuels.resample("D").last()

    out = RAW_DIR / "fuel_prices.parquet"
    df_fuels.to_parquet(out)
    print(f"\nSaved fuel_prices.parquet -> {out}")
    print(f"  Shape: {df_fuels.shape}")
    print(f"  Columns: {list(df_fuels.columns)}")
    print(f"  Period: {df_fuels.index.min()} -> {df_fuels.index.max()}")
    print(df_fuels.describe().round(2))
else:
    print("\nWARNING: No fuel price data could be fetched.")
    print("  Clean Spreads will be skipped in the feature matrix.")
    print("  Models will still train without fuel spread features.")
