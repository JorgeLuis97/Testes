import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import StrMethodFormatter

# Read the Excel file
file_path = r'C:\Users\jorge.rabelo\OneDrive - Konstroi\Área de Trabalho\Análise de Tickets.xlsx'
df = pd.read_excel(file_path)

# Filter the required columns
df_filtered = df.loc[:, ['Número', 'Status', 'Dias_Abertos']]

# Add a new column indicating if Dias_Abertos > 30
df_filtered.loc[:, 'Dias_Abertos_gt_30'] = df_filtered['Dias_Abertos'] > 30

# Group by Status and count the number of True values in Dias_Abertos_gt_30
status_counts = df_filtered[df_filtered['Dias_Abertos_gt_30']].groupby('Status').size()

# Create a figure with two subplots
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))

# Plot the bar plot on the first subplot
bars = ax1.bar(status_counts.index, status_counts.values)
ax1.set_xlabel('Status')
ax1.set_ylabel('Count')
ax1.set_title('Tickets Analysis - Status with Dias_Abertos > 30')
ax1.set_xticklabels(status_counts.index, rotation=45)
ax1.yaxis.set_major_formatter(StrMethodFormatter('{x:.0f}'))

# Add data labels to the bar plot
for bar in bars:
    height = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width() / 2, height, height, ha='center', va='bottom')

# Create a table on the second subplot
table_data = df_filtered[['Número', 'Status']]
table = ax2.table(cellText=table_data.values, colLabels=table_data.columns, loc='center')
table.auto_set_font_size(False)
table.set_fontsize(8)
table.scale(1.2, 1.2)

# Hide axis and labels on the second subplot
ax2.axis('off')

# Adjust the spacing between subplots
plt.subplots_adjust(wspace=0.3)

# Add a scrollbar to the table
table.auto_set_column_width([0, 1])
fig.tight_layout(rect=[0, 0, 0.8, 1])

# Show the plot
plt.show()
