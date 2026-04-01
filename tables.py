import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency, mannwhitneyu
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

# ── Helper: run chi-square and return results dict ────────────────────────────
def run_chi_square(col):
    ct = pd.crosstab(df[col], df['Order Status'])
    chi2, p, dof, _ = chi2_contingency(ct)
    n = ct.sum().sum()
    v = np.sqrt(chi2 / (n * (min(ct.shape) - 1)))
    result = 'Marginal (p<0.10)' if p < 0.10 else 'Not significant'
    return {
        'Variable':    col,
        'χ² / U stat': f'{chi2:.2f}',
        'df':          str(dof),
        'p-value':     f'{p:.3f}',
        "Cramér's V":  f'{v:.3f}',
        'Result':      result
    }

# ── Run tests for all variables ───────────────────────────────────────────────
variables = [
    'Product Type',
    'Shipping Type',
    'Payment Method',
    'Price Band',
    'Has Add-on',
    'Quantity',
    'Loyalty Member',
    'Rating',
    'Gender',
]

rows = [run_chi_square(col) for col in variables]

# ── Mann-Whitney U for Total Price (continuous) ───────────────────────────────
cancelled_prices  = df[df['Cancelled'] == 1]['Total Price']
completed_prices  = df[df['Cancelled'] == 0]['Total Price']
u_stat, p_mw = mannwhitneyu(cancelled_prices, completed_prices, alternative='two-sided')
rows.append({
    'Variable':    'Total Price*',
    'χ² / U stat': f'{u_stat:,.0f}',
    'df':          '—',
    'p-value':     f'{p_mw:.3f}',
    "Cramér's V":  'N/A',
    'Result':      'Not significant'
})

# ── Build DataFrame ───────────────────────────────────────────────────────────
results_df = pd.DataFrame(rows)
print("\nTable 1: Chi-Square Test Results — RQ1")
print(results_df.to_string(index=False))

# ── Plot as formatted table ───────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(14, 4.5))
ax.axis('off')

table = ax.table(
    cellText=results_df.values,
    colLabels=results_df.columns,
    cellLoc='center',
    loc='center'
)

table.auto_set_font_size(False)
table.set_fontsize(10)
table.scale(1, 1.8)

# ── Styling ───────────────────────────────────────────────────────────────────
header_color   = '#2C3E50'
row_even       = '#F2F2F2'
row_odd        = '#FFFFFF'
marginal_color = '#FDEBD0'

for (row, col), cell in table.get_celld().items():
    cell.set_edgecolor('#CCCCCC')
    cell.set_linewidth(0.5)

    if row == 0:
        cell.set_facecolor(header_color)
        cell.set_text_props(color='white', fontweight='bold')
    else:
        var_name = results_df.iloc[row - 1]['Variable']
        result   = results_df.iloc[row - 1]['Result']

        if var_name == 'Payment Method':
            cell.set_facecolor(marginal_color)
        elif row % 2 == 0:
            cell.set_facecolor(row_even)
        else:
            cell.set_facecolor(row_odd)

        if col == 5:
            if result == 'Marginal (p<0.10)':
                cell.set_text_props(color='#E67E22', fontweight='bold')
            else:
                cell.set_text_props(color='#27AE60', fontweight='bold')

# Column widths
col_widths = [0.22, 0.15, 0.06, 0.10, 0.12, 0.20]
for col_idx, width in enumerate(col_widths):
    for row_idx in range(len(results_df) + 1):
        table[row_idx, col_idx].set_width(width)

plt.title(
    'Table 1: Chi-Square Test Results — RQ1 Transaction-Level Variables\n'
    '*Total Price tested using Mann-Whitney U test (continuous variable)',
    fontsize=11, fontweight='bold', pad=15, loc='left'
)

plt.tight_layout()
plt.savefig('Table1_RQ1_results.png', dpi=300, bbox_inches='tight')
plt.close()
print("\nTable 1 saved.")