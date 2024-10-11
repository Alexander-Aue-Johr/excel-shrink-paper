import pandas as pd

# Load the CSV file
file_path = 'data/size_reduction_results.csv'  # Adjust the path to your CSV file
data = pd.read_csv(file_path)

# Sort the data in descending order by 'Size Reduction Factor'
data_sorted = data.sort_values(by='Size Reduction Factor', ascending=False)

# Specify the number of rows to include
num_rows = 10  # Change this variable as needed

# Select the top entries and limit to 5 columns
top_entries = data_sorted[['File', 'Sheet', 'Original Size', 'Cleaned Size', 'Size Reduction Factor']].head(num_rows)

# Convert Size Reduction Factor to percentage (0-100 %)
top_entries['Size Reduction Factor'] = top_entries['Size Reduction Factor'] * 100

# Rename the column 'Size Reduction Factor' to 'Size Reduction (%)'
top_entries = top_entries.rename(columns={'Size Reduction Factor': 'Size Reduction (%)'})

# Generate LaTeX code for the table with proper formatting
latex_table = r"""\begin{table}[htbp]
    \caption{Top Size Reduction Results}
    \label{tab:top_size_reduction}
    \centering
    \begin{tabular}{l l r r r}
        \toprule
        File & Sheet & Original Size & Cleaned Size & Size Reduction (\%) \\
        \midrule
""" + top_entries.to_latex(index=False, header=False, float_format="%.2f").replace("\\\\", " \\\\") + r"""
        \bottomrule
    \end{tabular}
\end{table}
"""

# Save the table in a file
with open('tables/top_size_reduction_table.tex', 'w') as f:
    f.write(latex_table)

print("LaTeX table created: 'top_size_reduction_table.tex'")
