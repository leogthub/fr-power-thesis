"""
Download ERA5 hourly weather data for France via the Copernicus CDS API.
Requests are split month-by-month to stay within the CDS size limit.
Saves to data/raw/era5_YYYY_MM.nc, then merges to data/raw/era5_france.parquet.
"""
import sys
sys.path.insert(0, ".")

import cdsapi
import xarray as xr
import pandas as pd
from pathlib import Path
from src.config import CDS_API_KEY, CDS_API_URL, RAW_DIR

RAW_DIR.mkdir(parents=True, exist_ok=True)

AREA = [51.5, -5.5, 41.5, 10.0]  # [N, W, S, E] — France bounding box

VARIABLES = [
    "2m_temperature",
    "10m_u_component_of_wind",
    "10m_v_component_of_wind",
    "surface_solar_radiation_downwards",
    "total_precipitation",
]

YEARS = range(2018, 2025)
MONTHS = range(1, 13)

client = cdsapi.Client(url=CDS_API_URL, key=CDS_API_KEY, quiet=True)


def nc_to_dataframe(nc_path: Path) -> pd.DataFrame:
    import zipfile, io
    # CDS API v2 returns zip archives containing instant + accum nc files
    if zipfile.is_zipfile(nc_path):
        frames = []
        with zipfile.ZipFile(nc_path) as zf:
            for name in zf.namelist():
                if not name.endswith(".nc"):
                    continue
                data = zf.read(name)
                ds = xr.open_dataset(io.BytesIO(data))
                frames.append(ds.mean(dim=["latitude", "longitude"]).to_dataframe())
        df = pd.concat(frames, axis=1)
        df = df.loc[:, ~df.columns.duplicated()]
    else:
        ds = xr.open_dataset(nc_path)
        df = ds.mean(dim=["latitude", "longitude"]).to_dataframe()

    df.index = pd.to_datetime(df.index).tz_localize("UTC").tz_convert("Europe/Paris")
    if "u10" in df.columns and "v10" in df.columns:
        df["wind_speed_10m"] = (df["u10"] ** 2 + df["v10"] ** 2) ** 0.5
        df = df.drop(columns=["u10", "v10"])
    df = df.rename(columns={
        "t2m": "temperature_2m",
        "ssrd": "solar_radiation",
        "tp": "precipitation",
    })
    if "temperature_2m" in df.columns:
        df["temperature_2m"] = df["temperature_2m"] - 273.15
    return df


frames = []
total = len(YEARS) * len(list(MONTHS))
done = 0

for year in YEARS:
    for month in MONTHS:
        nc_path = RAW_DIR / f"era5_{year}_{month:02d}.nc"
        done += 1
        label = f"{year}-{month:02d} [{done}/{total}]"

        if nc_path.exists():
            print(f"  {label} already downloaded, loading...")
        else:
            import calendar
            _, n_days = calendar.monthrange(year, month)
            print(f"  Downloading ERA5 {label}...", flush=True)
            try:
                client.retrieve(
                    "reanalysis-era5-single-levels",
                    {
                        "product_type": "reanalysis",
                        "variable": VARIABLES,
                        "year": str(year),
                        "month": f"{month:02d}",
                        "day": [f"{d:02d}" for d in range(1, n_days + 1)],
                        "time": [f"{h:02d}:00" for h in range(24)],
                        "area": AREA,
                        "format": "netcdf",
                    },
                    str(nc_path),
                )
            except Exception as e:
                print(f"  ERROR {label}: {e}")
                continue

        try:
            frames.append(nc_to_dataframe(nc_path))
        except Exception as e:
            print(f"  Parse error {label}: {e}")

if frames:
    df_all = pd.concat(frames).sort_index()
    out_parquet = RAW_DIR / "era5_france.parquet"
    df_all.to_parquet(out_parquet)
    print(f"ERA5 merged -> {out_parquet} ({len(df_all):,} rows)")
else:
    print("No data downloaded.")
