import pandas as pd
import matplotlib.pyplot as plt

# Read CSV file
file_path = 'data\size_reduction_results.csv'  # Adjust the path to your CSV file
data = pd.read_csv(file_path)

# Sort data by Size Reduction Factor
data_sorted = data.sort_values(by='Size Reduction Factor', ascending=False)

# Set the 'Sheet' as the x-axis labels
x_labels = data_sorted['Sheet']

# Creating indices for x-axis ticks
x_ticks = range(1, len(x_labels) + 1)

# Plotting the area chart
plt.figure(figsize=(7, 5))  # Updated figure size
plt.fill_between(x_ticks, data_sorted['Size Reduction Factor'] * 100, color='#A7C6ED', label='Size Reduction Percentage')  # Pastel blue color

# Adding titles and labels
plt.title('Distribution of Size Reduction Percentages for Excel Sheets')
plt.xlabel('Excel Sheets (Index)')
plt.ylabel('Size Reduction (%)')  # Changed label to show percentage

# Add gridlines for the y-axis, using a subtle style
plt.grid(True, axis='y', linestyle='--', color='gray', alpha=0.6)

# Creating custom ticks: First at 1, then every 1000
tick_positions = [1] + list(range(1000, len(x_labels) + 1, 1000))
# Setting the x-axis ticks to start from 1 and then in intervals of 1000
plt.xticks(ticks=tick_positions, labels=tick_positions, rotation=45, ha='right')

# Show legend
plt.legend()

# Adjust layout to prevent clipping of tick-labels
plt.tight_layout()

# Save the plot
plt.savefig('images/charts/distribution_of_size_reduction.pdf', dpi=300, bbox_inches='tight')  # High DPI for print quality
plt.savefig('images/charts/distribution_of_size_reduction.png', dpi=300, bbox_inches='tight')  # High DPI for print quality

# Close the figure when no longer needed
plt.close()
