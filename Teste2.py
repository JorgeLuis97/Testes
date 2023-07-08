import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import StrMethodFormatter

# Read the Excel file
file_path = r'C:\Users\jorge.rabelo\OneDrive - Konstroi\Área de Trabalho\Análise de Tickets.xlsx'
df = pd.read_excel(file_path)

# Filter the required columns
df_filtered = df[['Número', 'Dias_Abertos', 'Dias_sem_ação']]

# Create a figure with two subplots
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))

# Plot the scatter plot on the first subplot
ax1.scatter(df_filtered['Dias_Abertos'], df_filtered['Dias_sem_ação'])
ax1.set_xlabel('Dias Abertos')
ax1.set_ylabel('Dias sem ação')
ax1.set_title('Tickets Analysis')
ax1.grid(True)

# Add ticket numbers as annotations
for i, num in enumerate(df_filtered['Número']):
    ax1.annotate(num, (df_filtered['Dias_Abertos'][i], df_filtered['Dias_sem_ação'][i]), fontsize=8, alpha=0.7)

# Create a table on the second subplot
table = ax2.table(cellText=df_filtered.values, colLabels=df_filtered.columns, loc='center')
table.auto_set_font_size(False)
table.set_fontsize(8)
table.scale(1.2, 1.2)

# Format numbers as integers in the table
for i, key in enumerate(df_filtered.columns):
    if key != 'Número':
        col = table.get_celld()[(0, i)]
        col.set_formatter(StrMethodFormatter("{x:.0f}"))

# Hide axis and labels on the second subplot
ax2.axis('off')

# Adjust the spacing between subplots
plt.subplots_adjust(wspace=0.3)

# Add a scroll bar to the table
table.auto_set_column_width([0, 1, 2])
fig.tight_layout()
scrollable_area = plt.Rectangle((0, 0), 0.1, 0.1, fill=False, edgecolor='none', visible=False)
fig.add_artist(scrollable_area)
ax2.set_position([0.3, 0.1, 0.6, 0.8])
ax2.add_artist(table)

# Show the plot
plt.show()
