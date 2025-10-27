import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import os

# Load data from CSV file
file_path: str = os.path.join(
    "scripts", "analyze_destatis_xlsx_files", "analysis_output.csv"
)
chart_output_path: str = os.path.join(
    "scripts",
    "analyze_destatis_xlsx_files",
    "charts",
    "combined_chart_distribution_of_size_reduction.pdf",
)

data = pd.read_csv(file_path)

# Create a figure with GridSpec to customize subplot layout
fig = plt.figure(figsize=(20, 6))
gs = GridSpec(2, 4, height_ratios=[13, 8], figure=fig)

# Sort data by Reduction Ratio for both plots
data_sorted = data.sort_values(by="Reduction Ratio", ascending=False)
x_labels = data_sorted["Sheet / Workbook"]
x_ticks = range(1, len(x_labels) + 1)

# Define an absolute buffer for padding on the x-axis
x_buffer = 250  # Set this to your desired value for the space on either side of the data (top row)
x_buffer_bottom = 1000  # Additional buffer specifically for the bottom row of subplots

# Define a y-axis buffer for adding padding to the y-limits
y_buffer = 10  # Set this value to control padding on the y-axis (as a percentage of max_y_value)

# Plotting the combined reduction chart on top (spanning all columns)
ax_combined = fig.add_subplot(gs[0, :])  # This spans all columns in the first row
ax_combined.fill_between(
    x_ticks,
    data_sorted["Reduction Ratio"] * 100,
    color="#A7C6ED",
    label="Reduction Ratio",
)
ax_combined.set_title("Combined Reduction Ratio", fontsize=18)
ax_combined.grid(True, axis="y", linestyle="--", color="gray", alpha=0.6)
# Removing the legend from the top chart
# ax_combined.legend()  # Remove or comment out this line to delete the legend

# Removing x-axis label for the top chart
# ax_combined.set_xlabel('Excel Sheets (Index)', fontsize=14)  # Removed or commented out to delete the label

# Set tick positions and labels for x-axis of the top chart
tick_positions = [1] + list(range(2000, len(x_labels) + 1, 2000))
ax_combined.set_xticks(tick_positions)
ax_combined.set_xticklabels(tick_positions, rotation=45, ha="right")
ax_combined.set_xlim(
    1 - x_buffer, len(x_labels) + x_buffer
)  # Set x-axis limits with padding for top plot

# Sort data by individual reduction ratios and plot in the lower row of subplots
# Skipping 'Reduction Ratio' since it's already represented in the combined plot
ratios = [
    ("Empty Cell Reduction Ratio", "#E69F00"),  # Orange
    ("Removed Columns Reduction Ratio", "#D55E00"),  # Red
    ("Rows Outside of Value Range Reduction Ratio", "#F0E442"),  # Yellow
    ("Whitespace Cell Reduction Ratio", "#009E73"),  # Green
]

# Creating individual reduction plots in the second row
axs = []
for i, (ratio, color) in enumerate(ratios):
    ax = fig.add_subplot(gs[1, i])  # Individual subplot for each reduction ratio
    data_sorted = data.sort_values(by=ratio, ascending=False)
    ax.fill_between(
        x_ticks, data_sorted[ratio] * 100, color=color
    )  # Multiply by 100 to convert to percentage
    ax.set_title(ratio, fontsize=14)
    ax.grid(True, axis="y", linestyle="--", color="gray", alpha=0.6)
    # Set y-limits to show values from 0 to 100
    ax.set_ylim(-y_buffer, 100 + y_buffer)  # Use y_buffer for padding on y-axis limits
    ax.set_xlim(
        1 - x_buffer_bottom, len(x_labels) + x_buffer_bottom
    )  # Set x-axis limits with additional padding for bottom row
    ax.set_yticks(
        [0, 20, 40, 60, 80, 100]
    )  # Set specific y-ticks for better readability
    axs.append(ax)

# Adding common x-axis label for the whole figure
fig.text(0.5, 0.03, "Excel Sheets (Index)", ha="center", fontsize=16)

# Adjusting the position of the shared y-axis label to align better with the subplots
fig.text(
    0.015, 0.5, "Size Reduction (%)", va="center", rotation="vertical", fontsize=16
)

# Custom ticks for x-axis, applied to each subplot in the lower row (individual reduction plots)
bottom_tick_positions = [1] + list(range(4000, len(x_labels) + 1, 4000))
for ax in axs:
    ax.set_xticks(bottom_tick_positions)
    ax.set_xticklabels(bottom_tick_positions, rotation=45, ha="right")

# Adjust layout to ensure proper alignment with explicit control of spacing
plt.tight_layout(
    rect=[0.03, 0.05, 0.97, 0.95]
)  # Adding a tight layout to prevent overlap
plt.subplots_adjust(
    hspace=0.6, wspace=0.2
)  # Increased hspace for more distance between rows

# create directory if it does not exist
os.makedirs(os.path.dirname(chart_output_path), exist_ok=True)

# Save the plot
plt.savefig(
    chart_output_path,
    dpi=300,
    bbox_inches="tight",
)  # High DPI for print quality

# Close the figure when no longer needed
plt.close()
