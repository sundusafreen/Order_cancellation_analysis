import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import chi2_contingency
import warnings
warnings.filterwarnings('ignore')

# ── Load and clean ────────────────────────────────────────────────────────────
df = pd.read_excel('Individual Assignment Data File Electronic_sales.xlsx')
df['Payment Method'] = df['Payment Method'].str.strip().replace('Paypal', 'PayPal')
df['Gender'] = df['Gender'].fillna(df['Gender'].mode()[0])
df['Cancelled'] = (df['Order Status'] == 'Cancelled').astype(int)
df['Has Add-on'] = df['Add-ons Purchased'].notna().astype(int)
df['Order Value'] = df['Total Price'] + df['Add-on Total'].fillna(0)

bins   = [0, 25, 35, 50, 65, 120]
labels = ['18–25', '26–35', '36–50', '51–65', '65+']
df['Age Group'] = pd.cut(df['Age'], bins=bins, labels=labels, right=True)

# ── Block 1: Pivot tables ─────────────────────────────────────────────────────
pivot1 = df.pivot_table(
    values='Cancelled',
    index='Age Group',
    columns='Loyalty Member',
    aggfunc='mean'
).mul(100).round(2)

pivot2 = df.pivot_table(
    values='Cancelled',
    index='Age Group',
    columns='Gender',
    aggfunc='mean'
).mul(100).round(2)

print("Risk Matrix 1: Age Group × Loyalty Member")
print(pivot1)
print("\nRisk Matrix 2: Age Group × Gender")
print(pivot2)

# ── Block 2: Heatmaps ─────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle(
    'Figure 4: Customer Segment Risk Matrices\n'
    'Cell values = cancellation rate (%). Red = high risk, Green = low risk.',
    fontsize=13, fontweight='bold'
)

# Heatmap 1: Age Group × Loyalty Member
sns.heatmap(
    pivot1,
    annot=True,
    fmt='.1f',
    cmap='RdYlGn_r',
    vmin=30, vmax=36,
    linewidths=0.8,
    linecolor='white',
    annot_kws={'size': 12, 'weight': 'bold'},
    ax=axes[0],
    cbar_kws={'label': 'Cancellation Rate (%)'}
)
axes[0].set_title('Age Group × Loyalty Member', fontsize=11, fontweight='bold', pad=10)
axes[0].set_ylabel('Age Group', fontsize=10)
axes[0].set_xlabel('Loyalty Member', fontsize=10)
axes[0].tick_params(axis='x', labelsize=9)
axes[0].tick_params(axis='y', labelsize=9, rotation=0)

# Heatmap 2: Age Group × Gender
sns.heatmap(
    pivot2,
    annot=True,
    fmt='.1f',
    cmap='RdYlGn_r',
    vmin=30, vmax=36,
    linewidths=0.8,
    linecolor='white',
    annot_kws={'size': 12, 'weight': 'bold'},
    ax=axes[1],
    cbar_kws={'label': 'Cancellation Rate (%)'}
)
axes[1].set_title('Age Group × Gender', fontsize=11, fontweight='bold', pad=10)
axes[1].set_ylabel('Age Group', fontsize=10)
axes[1].set_xlabel('Gender', fontsize=10)
axes[1].tick_params(axis='x', labelsize=9)
axes[1].tick_params(axis='y', labelsize=9, rotation=0)

plt.tight_layout()
plt.savefig('Figure4_RQ2_risk_matrices.png', dpi=300, bbox_inches='tight')
plt.close()
print("\nFigure 4 saved.")

# ── Block 3: Financial impact ─────────────────────────────────────────────────
revenue = df[df['Cancelled'] == 1].groupby(
    ['Age Group', 'Loyalty Member']
)['Order Value'].sum().reset_index()
revenue.columns = ['Age Group', 'Loyalty Member', 'Revenue Lost ($)']
revenue['Revenue Lost ($)'] = revenue['Revenue Lost ($)'].round(2)
revenue = revenue.sort_values('Revenue Lost ($)', ascending=False)

print("\nRevenue Lost by Segment (Age Group × Loyalty Member):")
print(revenue.to_string(index=False))

# Cancel rate per segment
rates = df.groupby(['Age Group', 'Loyalty Member']).agg(
    Cancel_Rate=('Cancelled', 'mean'),
    Total_Orders=('Cancelled', 'count'),
    Cancelled_Orders=('Cancelled', 'sum')
).reset_index()
rates['Cancel_Rate'] = rates['Cancel_Rate'].mul(100).round(2)
full = rates.merge(revenue, on=['Age Group', 'Loyalty Member'])
full = full.sort_values('Revenue Lost ($)', ascending=False)

print("\nFull Segment Summary:")
print(full.to_string(index=False))

# ── Block 4: Revenue bar chart ────────────────────────────────────────────────
fig2, ax = plt.subplots(figsize=(12, 6))

full['Segment'] = full['Age Group'].astype(str) + ' / ' + full['Loyalty Member'].astype(str)
colors = ['#c0392b' if r > 33 else '#2980b9' for r in full['Cancel_Rate']]

bars = ax.bar(
    full['Segment'], full['Revenue Lost ($)'],
    color=colors, edgecolor='white', linewidth=0.8
)

# Value labels on bars
for bar in bars:
    ax.text(
        bar.get_x() + bar.get_width() / 2,
        bar.get_height() + 5000,
        f'${bar.get_height():,.0f}',
        ha='center', va='bottom', fontsize=8, fontweight='bold'
    )

ax.set_title(
    'Figure 5: Revenue Lost to Cancellations by Customer Segment\n'
    'Red = cancel rate above 33%, Blue = below 33%',
    fontsize=12, fontweight='bold'
)
ax.set_ylabel('Total Revenue Lost ($)', fontsize=10)
ax.set_xlabel('Segment (Age Group / Loyalty Member)', fontsize=10)
ax.tick_params(axis='x', rotation=30, labelsize=8)
ax.yaxis.grid(True, linestyle=':', alpha=0.4)
ax.set_axisbelow(True)

plt.tight_layout()
plt.savefig('Figure5_RQ2_revenue_lost.png', dpi=300, bbox_inches='tight')
plt.close()
print("Figure 5 saved.")

# ── Block 5: Chi-square tests ─────────────────────────────────────────────────
for col, label in [('Age Group', 'Age Group'), ('Loyalty Member', 'Loyalty Member'), ('Gender', 'Gender')]:
    ct = pd.crosstab(df[col], df['Order Status'])
    chi2, p, dof, _ = chi2_contingency(ct)
    n = ct.sum().sum()
    v = np.sqrt(chi2 / (n * (min(ct.shape) - 1)))
    print(f"\n{label}: χ²={chi2:.2f}, df={dof}, p={p:.4f}, Cramér's V={v:.4f}")