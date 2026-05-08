"""
Download ERA5 hourly weather data for France via the Copernicus CDS API.
Variables: 2m temperature, 10m u/v wind components, surface solar radiation,
           total precipitation.
Saves to data/raw/era5_france_<year>.nc then merges to data/raw/era5_france.parquet
"""
import os
import cdsapi
import xarray as xr
import pandas as pd
from pathlib import Path
from src.config import CDS_API_KEY, CDS_API_URL, RAW_DIR

RAW_DIR.mkdir(parents=True, exist_ok=True)

# France bounding box [N, W, S, E]
AREA = [51.5, -5.5, 41.5, 10.0]

VARIABLES = [
    "2m_temperature",
    "10m_u_component_of_wind",
    "10m_v_component_of_wind",
    "surface_solar_radiation_downwards",
    "total_precipitation",
]

YEARS = [str(y) for y in range(2018, 2025)]

client = cdsapi.Client(url=CDS_API_URL, key=CDS_API_KEY)


def download_year(year: str) -> Path:
    out_path = RAW_DIR / f"era5_france_{year}.nc"
    if out_path.exists():
        print(f"  {year} already downloaded, skipping.")
        return out_path

    print(f"  Downloading ERA5 {year}...")
    client.retrieve(
        "reanalysis-era5-single-levels",
        {
            "product_type": "reanalysis",
            "variable": VARIABLES,
            "year": year,
            "month": [f"{m:02d}" for m in range(1, 13)],
            "day": [f"{d:02d}" for d in range(1, 32)],
            "time": [f"{h:02d}:00" for h in range(24)],
            "area": AREA,
            "format": "netcdf",
        },
        str(out_path),
    )
    return out_path


def nc_to_dataframe(nc_path: Path) -> pd.DataFrame:
    ds = xr.open_dataset(nc_path)
    # Spatial mean over France bounding box
    df = ds.mean(dim=["latitude", "longitude"]).to_dataframe()
    df.index = df.index.tz_localize("UTC").tz_convert("Europe/Paris")

    # Derive wind speed from u/v components
    if "u10" in df.columns and "v10" in df.columns:
        df["wind_speed_10m"] = (df["u10"] ** 2 + df["v10"] ** 2) ** 0.5
        df = df.drop(columns=["u10", "v10"])

    df = df.rename(columns={
        "t2m": "temperature_2m",
        "ssrd": "solar_radiation",
        "tp": "precipitation",
    })
    # Convert temperature from Kelvin to Celsius
    if "temperature_2m" in df.columns:
        df["temperature_2m"] = df["temperature_2m"] - 273.15

    return df


print("Downloading ERA5 data for France (2018–2024)...")
frames = []
for year in YEARS:
    nc_path = download_year(year)
    frames.append(nc_to_dataframe(nc_path))

df_all = pd.concat(frames).sort_index()
out_parquet = RAW_DIR / "era5_france.parquet"
df_all.to_parquet(out_parquet)
print(f"ERA5 merged dataset saved → {out_parquet} ({len(df_all):,} rows)")
