import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
import warnings; warnings.filterwarnings('ignore')

df = pd.read_parquet('data/processed/features.parquet')
target_col = 'price_da_eur_mwh'
feature_cols = [c for c in df.columns if c != target_col]
train = df[df.index <= '2024-04-30 23:00:00'].dropna(subset=[target_col])
X_train = train[feature_cols].ffill().fillna(0)
y_train = train[target_col]

rf = RandomForestRegressor(n_estimators=500, max_depth=10, min_samples_leaf=5,
                           max_features='sqrt', random_state=42, n_jobs=-1)
rf.fit(X_train, y_train)
imp = pd.Series(rf.feature_importances_, index=feature_cols).sort_values(ascending=False)

wx = {'temperature_2m','wind_speed_10m','solar_radiation','precipitation',
      'hdd','cdd','wind_power_proxy','weather_stress_index'}
lag_roll_cols = [c for c in feature_cols if 'lag' in c or 'roll' in c]
lag_total = sum(imp[k] for k in lag_roll_cols if k in imp.index)
wx_total = sum(imp[k] for k in wx if k in imp.index)

print("Price lag 24h:", round(imp['price_da_eur_mwh_lag24h']*100, 1))
print("Roll mean 24h:", round(imp['price_da_eur_mwh_roll24h_mean']*100, 1))
print("Price lag 48h:", round(imp['price_da_eur_mwh_lag48h']*100, 1))
print("Roll mean 168h:", round(imp['price_da_eur_mwh_roll168h_mean']*100, 1))
print("Price lag 168h:", round(imp['price_da_eur_mwh_lag168h']*100, 1))
print("TTF:", round(imp['ttf_eur_mwh']*100, 1))
print("Coal:", round(imp['coal_eur_t']*100, 1))
print("Nuclear avail:", round(imp['nuclear_avail_ratio']*100, 1))
print("All lags+rolls total:", round(lag_total*100, 1))
print("All weather total:", round(wx_total*100, 1))
