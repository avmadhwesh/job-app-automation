import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import timedelta, date
import mplcursors

# Load the Excel file
df = pd.read_excel('Applications_Fall.xlsx')

# Initialize the frequency map
freqmap = {}

# Iterate through rows
for index, row in df.iterrows():
    if pd.isna(row['Date Applied']):  # Check for empty or NaN values
        continue
    # Convert 'Date Applied' to datetime.date
    date_applied = pd.to_datetime(row['Date Applied']).date()
    if date_applied not in freqmap:
        freqmap[date_applied] = 1
    else:
        freqmap[date_applied] += 1


# Define the date range: one year ago to today
end_date = date.today()
start_date = end_date - timedelta(days=365)

# Prepare data for the heatmap
num_days = (end_date - start_date).days + 1

# Create a grid for the contribution graph
heatmap_data = np.zeros((7, (num_days // 7) + 1))  # 7 days a week
dates_grid = [[None for _ in range((num_days // 7) + 1)] for _ in range(7)]  # To store dates for hover info

# Fill the grid with frequencies and dates
for i in range(num_days):
    current_date = start_date + timedelta(days=i)
    week = i // 7
    day = current_date.weekday()  # Monday=0, Sunday=6
    heatmap_data[day, week] = freqmap.get(current_date, 0)
    dates_grid[day][week] = current_date  # Store the date for hover info

# Plot the heatmap
fig, ax = plt.subplots(figsize=(12, 3))  # Adjust figure size for a compact graph
im = ax.imshow(heatmap_data, cmap='Greens', aspect='equal', interpolation='none')

# Customize axes
ax.set_yticks([0, 2, 4])  # Show only Monday, Wednesday, Friday
ax.set_yticklabels(['Mon', 'Wed', 'Fri'], fontsize=8)
ax.set_xticks(range(0, heatmap_data.shape[1], 4))  # Show months on every 4th week
month_labels = [(start_date + timedelta(weeks=i * 4)).strftime('%b') for i in range(heatmap_data.shape[1] // 4 + 1)]
ax.set_xticklabels(month_labels, fontsize=8)


ax.grid(which="minor", color="gray", linestyle='-', linewidth=0.5)


# Remove spines and gridlines for a clean look
ax.spines[:].set_visible(True)
ax.tick_params(axis='both', length=0)

# Add hover functionality
cursor = mplcursors.cursor(im, hover=True)

@cursor.connect("add")
def on_add(sel):

    x, y = int(sel.target[0]), int(sel.target[1])


    # Retrieve the date and contributions
    date_hover = dates_grid[y][x]
    contributions = int(heatmap_data[y, x])


    # Display the information
    if date_hover:
        sel.annotation.set(text=f"{date_hover}: {contributions} contributions")
    else:
        sel.annotation.set(text="No data")

# Adjust the appearance of the cells
for edge, spine in ax.spines.items():
    spine.set_visible(True)


plt.title('Contribution Graph', fontsize=10, pad=10)
plt.tight_layout(pad=1.0)
plt.show()