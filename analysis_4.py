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
df['Add-on Total'] = df['Add-on Total'].fillna(0)
df['Order Value'] = df['Total Price'] + df['Add-on Total']

bins   = [0, 25, 35, 50, 65, 120]
labels = ['18–25', '26–35', '36–50', '51–65', '65+']
df['Age Group'] = pd.cut(df['Age'], bins=bins, labels=labels, right=True)

# Segment-level summary we already have from RQ2
df['Order Value'] = df['Total Price'] + df['Add-on Total'].fillna(0)

bins   = [0, 25, 35, 50, 65, 120]
labels = ['18–25', '26–35', '36–50', '51–65', '65+']
df['Age Group'] = pd.cut(df['Age'], bins=bins, labels=labels, right=True)

# Average order value overall
avg_order_value = df[df['Cancelled'] == 1]['Order Value'].mean()
print(f"Average cancelled order value: ${avg_order_value:.2f}")

# ── Intervention 1: Loyalty re-engagement (36-50 Loyalty Members) ─────────────
seg1 = df[(df['Age Group'] == '36–50') & (df['Loyalty Member'] == 'Yes')]
seg1_cancellations = seg1['Cancelled'].sum()
seg1_revenue_lost  = seg1[seg1['Cancelled'] == 1]['Order Value'].sum()
seg1_recovery      = 0.25  # 25% recovery assumption
seg1_recoverable   = seg1_revenue_lost * seg1_recovery
seg1_cost          = 15000  # estimated campaign cost
seg1_roi           = (seg1_recoverable - seg1_cost) / seg1_cost * 100

# ── Intervention 2: Pre-cancellation popup (all orders) ───────────────────────
all_cancellations  = df['Cancelled'].sum()
all_revenue_lost   = df[df['Cancelled'] == 1]['Order Value'].sum()
seg2_recovery      = 0.15
seg2_recoverable   = all_revenue_lost * seg2_recovery
seg2_cost          = 5000   # one-time development cost
seg2_roi           = (seg2_recoverable - seg2_cost) / seg2_cost * 100

# ── Intervention 3: Targeted discount (Non-loyalty 51-65 and 65+) ─────────────
seg3 = df[
    (df['Loyalty Member'] == 'No') &
    (df['Age Group'].isin(['51–65', '65+']))
]
seg3_revenue_lost  = seg3[seg3['Cancelled'] == 1]['Order Value'].sum()
seg3_recovery      = 0.20
seg3_recoverable   = seg3_revenue_lost * seg3_recovery
seg3_cost          = 25000  # discount budget
seg3_roi           = (seg3_recoverable - seg3_cost) / seg3_cost * 100

# ── Intervention 4: Young male engagement nudge (18-25 Males) ─────────────────
seg4 = df[(df['Age Group'] == '18–25') & (df['Gender'] == 'Male')]
seg4_revenue_lost  = seg4[seg4['Cancelled'] == 1]['Order Value'].sum()
seg4_recovery      = 0.20
seg4_recoverable   = seg4_revenue_lost * seg4_recovery
seg4_cost          = 10000  # targeted campaign cost
seg4_roi           = (seg4_recoverable - seg4_cost) / seg4_cost * 100

print(f"\nIntervention 1 — Loyalty re-engagement:")
print(f"  Cancellations: {seg1_cancellations}, Revenue lost: ${seg1_revenue_lost:,.0f}")
print(f"  Recoverable: ${seg1_recoverable:,.0f}, Cost: ${seg1_cost:,.0f}, ROI: {seg1_roi:.1f}%")

print(f"\nIntervention 2 — Pre-cancellation popup:")
print(f"  Cancellations: {all_cancellations}, Revenue lost: ${all_revenue_lost:,.0f}")
print(f"  Recoverable: ${seg2_recoverable:,.0f}, Cost: ${seg2_cost:,.0f}, ROI: {seg2_roi:.1f}%")

print(f"\nIntervention 3 — Targeted discount:")
print(f"  Revenue lost: ${seg3_revenue_lost:,.0f}")
print(f"  Recoverable: ${seg3_recoverable:,.0f}, Cost: ${seg3_cost:,.0f}, ROI: {seg3_roi:.1f}%")

print(f"\nIntervention 4 — Young male nudge:")
print(f"  Revenue lost: ${seg4_revenue_lost:,.0f}")
print(f"  Recoverable: ${seg4_recoverable:,.0f}, Cost: ${seg4_cost:,.0f}, ROI: {seg4_roi:.1f}%")

import matplotlib.pyplot as plt
import numpy as np

interventions = [
    'Pre-cancellation\nPopup',
    'Targeted Discount\n(51-65 & 65+ Non-Loyalty)',
    'Loyalty Re-engagement\n(36-50)',
    'Young Male\nEngagement Nudge'
]

recoverable = [seg2_recoverable, seg3_recoverable,
               seg1_recoverable, seg4_recoverable]

costs = [seg2_cost, seg3_cost, seg1_cost, seg4_cost]

# Sort by recoverable revenue descending
sorted_idx  = np.argsort(recoverable)[::-1]
interventions_s = [interventions[i] for i in sorted_idx]
recoverable_s   = [recoverable[i]   for i in sorted_idx]
costs_s         = [costs[i]         for i in sorted_idx]

# Cumulative percentage for Pareto line
cumulative_pct = np.cumsum(recoverable_s) / sum(recoverable_s) * 100

fig, ax1 = plt.subplots(figsize=(12, 6))

# Bars
bars = ax1.bar(
    interventions_s, recoverable_s,
    color=['#c0392b', '#e67e22', '#2980b9', '#27ae60'],
    edgecolor='white', linewidth=0.8, width=0.5
)

# Value labels on bars
for bar in bars:
    ax1.text(
        bar.get_x() + bar.get_width() / 2,
        bar.get_height() + 20000,
        f'${bar.get_height():,.0f}',
        ha='center', va='bottom', fontsize=9, fontweight='bold'
    )

ax1.set_ylabel('Recoverable Revenue ($)', fontsize=10)
ax1.set_xlabel('Intervention', fontsize=10)
ax1.yaxis.grid(True, linestyle=':', alpha=0.4)
ax1.set_axisbelow(True)

# Pareto cumulative line on second axis
ax2 = ax1.twinx()
ax2.plot(
    interventions_s, cumulative_pct,
    color='black', marker='o', linewidth=2,
    markersize=7, label='Cumulative %'
)
ax2.axhline(80, color='gray', linestyle='--',
            linewidth=1, label='80% threshold')
ax2.set_ylabel('Cumulative Recovery (%)', fontsize=10)
ax2.set_ylim(0, 110)

# Annotate cumulative % on line
for x, y in zip(interventions_s, cumulative_pct):
    ax2.annotate(
        f'{y:.1f}%', (x, y),
        textcoords='offset points',
        xytext=(0, 10), ha='center', fontsize=8, color='black'
    )

ax2.legend(loc='center right', fontsize=9)

plt.title(
    'Figure 6: Pareto Analysis — Recoverable Revenue by Intervention\n'
    'Bars = recoverable revenue. Line = cumulative % of total recoverable revenue.',
    fontsize=11, fontweight='bold'
)
plt.tight_layout()
plt.savefig('Figure6_RQ3_pareto.png', dpi=300, bbox_inches='tight')
plt.close()
print("Figure 6 saved.")

roi_values = [seg1_roi, seg2_roi, seg3_roi, seg4_roi]
roi_labels = [
    'Loyalty\nRe-engagement',
    'Pre-cancellation\nPopup',
    'Targeted\nDiscount',
    'Young Male\nNudge'
]

colors = ['#c0392b' if r > 1000 else '#2980b9' for r in roi_values]

fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.bar(
    roi_labels, roi_values,
    color=colors, edgecolor='white', linewidth=0.8, width=0.5
)

for bar in bars:
    ax.text(
        bar.get_x() + bar.get_width() / 2,
        bar.get_height() + 50,
        f'{bar.get_height():,.0f}%',
        ha='center', va='bottom', fontsize=9, fontweight='bold'
    )

ax.set_ylabel('Return on Investment (%)', fontsize=10)
ax.set_xlabel('Intervention', fontsize=10)
ax.yaxis.grid(True, linestyle=':', alpha=0.4)
ax.set_axisbelow(True)
plt.title(
    'Figure 7: Projected ROI by Intervention\n'
    '(Based on conservative recovery rate assumptions)',
    fontsize=11, fontweight='bold'
)
plt.tight_layout()
plt.savefig('Figure7_RQ3_roi.png', dpi=300, bbox_inches='tight')
plt.close()
print("Figure 7 saved.")