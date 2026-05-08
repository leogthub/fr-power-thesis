from pathlib import Path
from dotenv import load_dotenv
import os

load_dotenv()

ROOT = Path(__file__).resolve().parents[1]

# Data paths
RAW_DIR = ROOT / "data" / "raw"
INTERIM_DIR = ROOT / "data" / "interim"
PROCESSED_DIR = ROOT / "data" / "processed"

# Output paths
FIGURES_DIR = ROOT / "outputs" / "figures"
TABLES_DIR = ROOT / "outputs" / "tables"
METRICS_DIR = ROOT / "outputs" / "metrics"

# API keys
ENTSOE_API_KEY = os.getenv("ENTSOE_API_KEY", "")
CDS_API_KEY = os.getenv("CDS_API_KEY", "")
CDS_API_URL = os.getenv("CDS_API_URL", "https://cds.climate.copernicus.eu/api")

# Market config
BIDDING_ZONE = "10YFR-RTE------C"  # France RTE
TIMEZONE = "Europe/Paris"

# Modelling
TARGET_COL = "price_da_eur_mwh"
TEST_MONTHS = 12
RANDOM_SEED = 42

RF_PARAMS = {
    "n_estimators": 500,
    "max_depth": 10,
    "min_samples_leaf": 5,
    "n_jobs": -1,
    "random_state": RANDOM_SEED,
}

XGB_PARAMS = {
    "n_estimators": 500,
    "max_depth": 6,
    "learning_rate": 0.05,
    "subsample": 0.8,
    "colsample_bytree": 0.8,
    "random_state": RANDOM_SEED,
}
