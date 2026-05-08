"""
Build final feature matrix from interim merged data.
Output: data/processed/features.parquet
"""
import sys
sys.path.insert(0, ".")

import pandas as pd
from src.config import INTERIM_DIR, PROCESSED_DIR
from src.features import build_feature_matrix

PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

print("Loading interim data...")
df = pd.read_parquet(INTERIM_DIR / "merged.parquet")
print(f"  {len(df):,} rows, {df.shape[1]} columns")

print("Engineering features...")
X, y = build_feature_matrix(df)
print(f"  Feature matrix: {X.shape[0]:,} rows x {X.shape[1]} features")
print(f"  Features: {list(X.columns)}")

out = PROCESSED_DIR / "features.parquet"
df_out = X.copy()
df_out["price_da_eur_mwh"] = y
df_out.to_parquet(out)
print(f"Saved -> {out}")
