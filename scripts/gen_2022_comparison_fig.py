"""Generate model_comparison_2022_vs_2024.png comparing MAE and R2 across regimes."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

models = ['Naive\n(A)', 'RF no wx\n(B)', 'RF wx\n(C)', 'XGBoost\n(D)']

mae_2022  = [72.74, 67.28, 66.55, 76.32]
r2_2022   = [0.419, 0.464, 0.475, 0.282]
mae_2024  = [33.09, 16.95, 16.94, 19.14]
r2_2024   = [0.160, 0.775, 0.776, 0.659]

x = np.arange(len(models))
w = 0.35

fig, axes = plt.subplots(1, 2, figsize=(13, 5))
fig.suptitle(
    'Model Performance: 2022 Energy Crisis vs. 2024-25 Stable Period',
    fontsize=14, fontweight='bold', y=1.02
)

COLOR_CRISIS = '#c0392b'
COLOR_STABLE = '#2563a8'

# ── MAE panel ──────────────────────────────────────────────────────────────
ax = axes[0]
b1 = ax.bar(x - w/2, mae_2022, w, label='2022 Crisis',
            color=COLOR_CRISIS, alpha=0.85, edgecolor='white')
b2 = ax.bar(x + w/2, mae_2024, w, label='2024-25 Stable',
            color=COLOR_STABLE, alpha=0.85, edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(models, fontsize=11)
ax.set_ylabel('MAE (EUR/MWh)', fontsize=12)
ax.set_title('Mean Absolute Error', fontsize=12, fontweight='bold')
ax.legend(fontsize=10)
ax.grid(axis='y', alpha=0.3, linestyle='--')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for bar in b1:
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
            f'{bar.get_height():.1f}', ha='center', va='bottom',
            fontsize=9, fontweight='bold')
for bar in b2:
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
            f'{bar.get_height():.1f}', ha='center', va='bottom',
            fontsize=9, fontweight='bold')

ax.annotate('DM: p<0.001***\n(C vs B in 2022)',
            xy=(1.5, 68.5), fontsize=8.5, color=COLOR_CRISIS, ha='center',
            style='italic',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='#fdecea',
                      edgecolor=COLOR_CRISIS, alpha=0.9))
ax.annotate('DM: p=0.572\n(C vs B in 2024-25)',
            xy=(2.5, 18), fontsize=8.5, color=COLOR_STABLE, ha='center',
            style='italic',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='#e8f0fb',
                      edgecolor=COLOR_STABLE, alpha=0.9))

# ── R2 panel ───────────────────────────────────────────────────────────────
ax = axes[1]
b3 = ax.bar(x - w/2, r2_2022, w, label='2022 Crisis',
            color=COLOR_CRISIS, alpha=0.85, edgecolor='white')
b4 = ax.bar(x + w/2, r2_2024, w, label='2024-25 Stable',
            color=COLOR_STABLE, alpha=0.85, edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(models, fontsize=11)
ax.set_ylabel('R² (Coefficient of Determination)', fontsize=12)
ax.set_title('R² Score', fontsize=12, fontweight='bold')
ax.legend(fontsize=10)
ax.set_ylim(0, 1.0)
ax.grid(axis='y', alpha=0.3, linestyle='--')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

for bar in b3:
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
            f'{bar.get_height():.3f}', ha='center', va='bottom',
            fontsize=9, fontweight='bold')
for bar in b4:
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
            f'{bar.get_height():.3f}', ha='center', va='bottom',
            fontsize=9, fontweight='bold')

plt.tight_layout()
plt.savefig('outputs/figures/model_comparison_2022_vs_2024.png',
            dpi=150, bbox_inches='tight')
plt.savefig('thesis/figures/model_comparison_2022_vs_2024.png',
            dpi=150, bbox_inches='tight')
print('Saved model_comparison_2022_vs_2024.png')
