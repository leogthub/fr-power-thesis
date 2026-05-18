"""Regenerate feature_importance_rf.png with 500 trees and correct lag24h ordering."""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from sklearn.ensemble import RandomForestRegressor
import warnings
warnings.filterwarnings('ignore')

print('Loading data...')
df = pd.read_parquet('data/processed/features.parquet')
target_col = 'price_da_eur_mwh'
feature_cols = [c for c in df.columns if c != target_col]

TRAIN_END = '2024-04-30 23:00:00'
train = df[df.index <= TRAIN_END].dropna(subset=[target_col])
X_train = train[feature_cols].ffill().fillna(0)
y_train = train[target_col]

print('Training RF (500 trees)...')
rf = RandomForestRegressor(
    n_estimators=500, max_depth=10, min_samples_leaf=5,
    max_features='sqrt', random_state=42, n_jobs=-1
)
rf.fit(X_train, y_train)

imp = pd.Series(rf.feature_importances_, index=feature_cols).sort_values(ascending=False)
top20 = imp.head(20)

WEATHER_COLS = {
    'temperature_2m', 'wind_speed_10m', 'solar_radiation', 'precipitation',
    'hdd', 'cdd', 'wind_power_proxy', 'weather_stress_index'
}

colors = ['#c0392b' if c in WEATHER_COLS else '#2563a8' for c in top20.index]

label_map = {
    'price_da_eur_mwh_lag24h':       'Price lag 24h',
    'price_da_eur_mwh_roll24h_mean': 'Rolling mean 24h',
    'price_da_eur_mwh_roll168h_mean':'Rolling mean 168h',
    'price_da_eur_mwh_lag168h':      'Price lag 168h',
    'price_da_eur_mwh_lag48h':       'Price lag 48h',
    'ttf_eur_mwh':                   'TTF gas price',
    'coal_eur_t':                    'ARA coal price',
    'nuclear_avail_ratio':           'Nuclear avail. ratio',
    'gen_nuclear_mw':                'Nuclear generation',
    'gen_gas_mw':                    'Gas generation',
    'flow_net_fr_de_mw':             'Flow FR-DE',
    'flow_net_fr_es_mw':             'Flow FR-ES',
    'flow_net_fr_gb_mw':             'Flow FR-GB',
    'flow_net_fr_be_mw':             'Flow FR-BE',
    'load_forecast_mw':              'Load forecast',
    'dayofweek':                     'Day of week',
    'temperature_2m':                'Temperature (2m)',
    'wind_speed_10m':                'Wind speed (10m)',
    'hdd':                           'HDD',
    'gen_hydro_reservoir_mw':        'Hydro reservoir',
    'gen_wind_onshore_mw':           'Wind generation',
    'gen_solar_mw':                  'Solar generation',
    'solar_radiation':               'Solar radiation',
    'weather_stress_index':          'Weather Stress Index',
    'wind_power_proxy':              'Wind power proxy',
    'month_sin':                     'Month (sin)',
    'hour_sin':                      'Hour (sin)',
    'hour_cos':                      'Hour (cos)',
    'gen_hydro_ror_mw':              'Run-of-river hydro',
    'is_weekend':                    'Is weekend',
}
labels = [label_map.get(c, c) for c in top20.index]

fig, ax = plt.subplots(figsize=(10, 7))
bars = ax.barh(range(len(top20)), top20.values * 100,
               color=colors, edgecolor='white', linewidth=0.5)
ax.set_yticks(range(len(top20)))
ax.set_yticklabels(labels, fontsize=11)
ax.invert_yaxis()
ax.set_xlabel('MDI Importance (%)', fontsize=12)
ax.set_title(
    'Top 20 Features by MDI Importance\nRandom Forest with Weather (Model C)',
    fontsize=13, fontweight='bold', pad=15
)
ax.grid(axis='x', alpha=0.3, linestyle='--')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for i, (bar, val) in enumerate(zip(bars, top20.values)):
    ax.text(val * 100 + 0.1, i, f'{val*100:.1f}%',
            va='center', fontsize=9, color='#333333')

legend_elements = [
    Patch(facecolor='#2563a8', label='Non-weather features'),
    Patch(facecolor='#c0392b', label='Weather features'),
]
ax.legend(handles=legend_elements, loc='lower right', fontsize=10)

plt.tight_layout()
plt.savefig('outputs/figures/feature_importance_rf.png', dpi=150, bbox_inches='tight')
plt.savefig('thesis/figures/feature_importance_rf.png', dpi=150, bbox_inches='tight')
print('Saved feature_importance_rf.png')
print('Top 5:')
for feat, val in zip(labels[:5], top20.values[:5]):
    print(f'  {feat}: {val*100:.1f}%')
