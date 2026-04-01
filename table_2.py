import pandas as pd
import matplotlib.pyplot as plt

# ── Data ──────────────────────────────────────────────────────────────────────
data = {
    'Intervention': [
        'Pre-cancellation\nConfirmation Popup',
        'Targeted Retention\nDiscount',
        'Loyalty Re-engagement\nCampaign',
        'Young Male\nEngagement Nudge'
    ],
    'Target Segment': [
        'All orders',
        'Non-Loyalty,\naged 51–65 & 65+',
        'Loyalty Members,\naged 36–50',
        'Male customers,\naged 18–25'
    ],
    'Rationale': [
        'Uniform cancellation rate\nsuggests behavioural cause',
        'Highest absolute revenue\nloss (>$8M combined)',
        'Cancel rate (34.1%) exceeds\nnon-loyalty peers (32.8%)',
        'Highest gender-based\ncancel rate (34.5%)'
    ],
    'Recovery\nRate': ['15%', '20%', '25%', '20%'],
    'Est. Cost': ['$5,000', '$25,000', '$15,000', '$10,000'],
    'Recoverable\nRevenue': [
        '$3,207,353', '$1,608,981', '$295,076', '$256,716'
    ],
    'Projected\nROI': ['64,047%', '6,336%', '1,867%', '2,467%']
}

df_table = pd.DataFrame(data)

# ── Figure ────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(18, 4.5))
ax.axis('off')

table = ax.table(
    cellText=df_table.values,
    colLabels=df_table.columns,
    cellLoc='center',
    loc='center'
)

table.auto_set_font_size(False)
table.set_fontsize(9)
table.scale(1, 2.8)

# ── Styling ───────────────────────────────────────────────────────────────────
header_color = '#2C3E50'
row_colors   = ['#FFFFFF', '#F2F2F2']
roi_color    = '#EBF5FB'

for (row, col), cell in table.get_celld().items():
    cell.set_edgecolor('#CCCCCC')
    cell.set_linewidth(0.5)

    if row == 0:
        # Header
        cell.set_facecolor(header_color)
        cell.set_text_props(color='white', fontweight='bold')
    else:
        # Alternate row shading
        cell.set_facecolor(row_colors[(row - 1) % 2])

        # Highlight ROI column
        if col == 6:
            cell.set_facecolor(roi_color)
            cell.set_text_props(color='#1A5276', fontweight='bold')

        # Highlight Intervention column
        if col == 0:
            cell.set_text_props(fontweight='bold')

# Column widths
col_widths = [0.18, 0.16, 0.22, 0.09, 0.09, 0.14, 0.12]
for col_idx, width in enumerate(col_widths):
    for row_idx in range(len(df_table) + 1):
        table[row_idx, col_idx].set_width(width)

plt.title(
    'Table 2: Proposed Interventions — RQ3 Summary\n'
    'Recovery rates are conservative estimates. Costs are indicative.',
    fontsize=11, fontweight='bold', pad=15, loc='left'
)

plt.tight_layout()
plt.savefig('Table2_RQ3_interventions.png', dpi=300, bbox_inches='tight')
plt.close()
print("Table 2 saved.")