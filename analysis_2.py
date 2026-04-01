import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')

# ── Load and clean ────────────────────────────────────────────────────────────
df = pd.read_excel('Individual Assignment Data File Electronic_sales.xlsx')
df['Payment Method'] = df['Payment Method'].str.strip().replace('Paypal', 'PayPal')
df['Gender'] = df['Gender'].fillna(df['Gender'].mode()[0])
df['Cancelled'] = (df['Order Status'] == 'Cancelled').astype(int)
df['Has Add-on'] = df['Add-ons Purchased'].notna().astype(int)
df['Price Band'] = pd.cut(
    df['Total Price'],
    bins=[0, 200, 500, 1000, 2000, 99999],
    labels=['<$200', '$200–500', '$500–1k', '$1k–2k', '>$2k']
)
df['Month'] = pd.to_datetime(df['Purchase Date']).dt.to_period('M')

overall = df['Cancelled'].mean() * 100

# ── Figure 1: 6-panel bar chart ───────────────────────────────────────────────
fig, axes = plt.subplots(2, 3, figsize=(16, 10))
fig.suptitle(
    'Figure 1: Cancellation Rate by Key Transaction Variables\n'
    'Dashed line = overall rate (32.8%). Red = above average, Blue = below average.',
    fontsize=13, fontweight='bold', y=1.02
)

plot_vars = [
    ('Product Type',   'χ²=0.87, p=0.928, V=0.007'),
    ('Shipping Type',  'χ²=3.50, p=0.478, V=0.013'),
    ('Payment Method', 'χ²=8.42, p=0.077, V=0.021'),
    ('Price Band',     'χ²=1.47, p=0.833, V=0.009'),
    ('Rating',         'χ²=1.60, p=0.809, V=0.009'),
    ('Quantity',       'χ²=6.32, p=0.707, V=0.018'),
]

for ax, (col, stats) in zip(axes.flatten(), plot_vars):
    rates = df.groupby(col)['Cancelled'].mean().mul(100).sort_values(ascending=False)
    colors = ['#c0392b' if r > overall else '#2980b9' for r in rates.values]

    bars = ax.bar(
        rates.index.astype(str), rates.values,
        color=colors, edgecolor='white', linewidth=0.8, width=0.6
    )

    # Overall rate reference line
    ax.axhline(
        overall, color='black', linestyle='--',
        linewidth=1.2, label=f'Overall: {overall:.1f}%'
    )

    # Value labels on top of each bar
    for bar in bars:
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.15,
            f'{bar.get_height():.1f}%',
            ha='center', va='bottom', fontsize=8, fontweight='bold'
        )

    # Titles, labels, stats
    ax.set_title(col, fontsize=11, fontweight='bold', pad=8)
    ax.set_xlabel(stats, fontsize=8, color='dimgrey', labelpad=6)
    ax.set_ylabel('Cancellation Rate (%)', fontsize=9)
    ax.set_ylim(28, 38)
    ax.tick_params(axis='x', labelsize=8, rotation=20)
    ax.tick_params(axis='y', labelsize=8)
    ax.legend(fontsize=7, loc='upper right')
    ax.yaxis.grid(True, linestyle=':', alpha=0.5)
    ax.set_axisbelow(True)

plt.tight_layout()
plt.savefig('Figure1_RQ1_cancellation_by_variable.png', dpi=300, bbox_inches='tight')
#plt.show()
plt.close() 
print("Figure 1 saved.")

# ── Figure 2: Payment Method stacked bar ─────────────────────────────────────
fig2, ax2 = plt.subplots(figsize=(9, 5))

pm_ct = pd.crosstab(
    df['Payment Method'], df['Order Status'], normalize='index'
).mul(100).sort_values('Cancelled', ascending=False)

pm_ct.plot(
    kind='bar', stacked=True, ax=ax2,
    color=['#c0392b', '#27ae60'], edgecolor='white', width=0.6
)

# Value labels inside each bar segment
for bar_group in ax2.containers:
    ax2.bar_label(
        bar_group, fmt='%.1f%%',
        label_type='center', fontsize=9,
        fontweight='bold', color='white'
    )

ax2.axhline(overall, color='black', linestyle='--',
            linewidth=1.2, label=f'Overall cancel rate: {overall:.1f}%')

ax2.set_title(
    'Figure 2: Order Status Distribution by Payment Method\n'
    '(χ²=8.42, p=0.077 — closest to significance but not significant at p<0.05)',
    fontsize=11, fontweight='bold'
)
ax2.set_ylabel('Percentage of Orders (%)', fontsize=10)
ax2.set_xlabel('Payment Method', fontsize=10)
ax2.set_ylim(0, 110)
ax2.tick_params(axis='x', rotation=20, labelsize=9)
ax2.legend(fontsize=9, loc='upper right')
ax2.yaxis.grid(True, linestyle=':', alpha=0.4)
ax2.set_axisbelow(True)

plt.tight_layout()
plt.savefig('Figure2_RQ1_payment_method.png', dpi=300, bbox_inches='tight')
#plt.show()
plt.close() 
print("Figure 2 saved.")

# ── Figure 3: Monthly trend line ─────────────────────────────────────────────
monthly = df.groupby('Month')['Cancelled'].mean().mul(100)
monthly.index = monthly.index.astype(str)
mean_rate = monthly.mean()

fig3, ax3 = plt.subplots(figsize=(13, 5))

ax3.plot(
    monthly.index, monthly.values,
    marker='o', color='#c0392b',
    linewidth=2, markersize=7, zorder=3
)

# Shading above/below mean
ax3.fill_between(
    monthly.index, monthly.values, mean_rate,
    where=(monthly.values > mean_rate),
    alpha=0.12, color='#c0392b', label='Above average'
)
ax3.fill_between(
    monthly.index, monthly.values, mean_rate,
    where=(monthly.values <= mean_rate),
    alpha=0.12, color='#2980b9', label='Below average'
)

# Mean reference line
ax3.axhline(
    mean_rate, color='gray', linestyle='--',
    linewidth=1.2, label=f'Mean: {mean_rate:.1f}%'
)

# Value labels on each data point
for x, y in zip(monthly.index, monthly.values):
    ax3.annotate(
        f'{y:.1f}%', (x, y),
        textcoords='offset points', xytext=(0, 10),
        ha='center', fontsize=8, fontweight='bold', color='#c0392b'
    )

ax3.set_title(
    'Figure 3: Monthly Cancellation Rate (Sep 2023 – Sep 2024)\n'
    '(χ²=6.46, p=0.891 — no significant temporal trend)',
    fontsize=12, fontweight='bold'
)
ax3.set_ylabel('Cancellation Rate (%)', fontsize=10)
ax3.set_xlabel('Month', fontsize=10)
ax3.set_ylim(28, 40)
ax3.tick_params(axis='x', rotation=35, labelsize=8)
ax3.tick_params(axis='y', labelsize=9)
ax3.legend(fontsize=9)
ax3.yaxis.grid(True, linestyle=':', alpha=0.4)
ax3.set_axisbelow(True)

plt.tight_layout()
plt.savefig('Figure3_RQ1_monthly_trend.png', dpi=300, bbox_inches='tight')
#plt.show()
plt.close() 
print("Figure 3 saved.")