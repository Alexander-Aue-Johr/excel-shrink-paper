"""Generate the Enron distributions with the shared corpus plotting code."""
from pathlib import Path
import sys


ANALYSIS_DIRECTORY = Path(__file__).resolve().parent
SHARED_DIRECTORY = ANALYSIS_DIRECTORY.parent / "analyze_destatis_xlsx_files"
sys.path.insert(0, str(SHARED_DIRECTORY))

import generate_distribution_of_size_reduction_chart_rust as chart


def main():
    source = ANALYSIS_DIRECTORY / "excel_shrink_rust_analysis.csv.gz"
    if not source.is_file():
        raise FileNotFoundError(f"The complete Enron analysis table is required at {source}")
    data = chart.read_rust_analysis(source)
    output = ANALYSIS_DIRECTORY / "charts"
    sheets = chart.build_distribution_data(data)
    workbooks = chart.build_file_distribution_data(data)
    for label, frame in (("sheet", sheets), ("workbook", workbooks)):
        target = output / f"enron-reduction-by-{label}.pdf"
        chart.plot_distribution(frame, target, target.with_suffix(".png"),
                                title=f"Enron structural XML reduction by {label}",
                                x_axis_label=f"{label.capitalize()} rank, sorted independently within each panel")
    print(f"Enron analysis contains {len(workbooks)} workbooks and {len(sheets)} sheet/workbook XML entries")


if __name__ == "__main__":
    main()
