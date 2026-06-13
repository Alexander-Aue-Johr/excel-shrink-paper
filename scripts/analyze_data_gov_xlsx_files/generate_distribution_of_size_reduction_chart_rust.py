from __future__ import annotations

import sys
from pathlib import Path


DESTatis_ANALYSIS_DIR = (
    Path(__file__).resolve().parents[1] / "analyze_destatis_xlsx_files"
)
sys.path.insert(0, str(DESTatis_ANALYSIS_DIR))

import generate_distribution_of_size_reduction_chart_rust as chart


DATA_GOV_ANALYSIS_DIR = Path("scripts", "analyze_data_gov_xlsx_files")

chart.INPUT_PATH = DATA_GOV_ANALYSIS_DIR / "excel_shrink_rust_analysis.csv.gz"
chart.CHART_OUTPUT_PATH = (
    DATA_GOV_ANALYSIS_DIR
    / "charts"
    / "combined_chart_distribution_of_size_reduction_rust.pdf"
)
chart.CHART_PREVIEW_OUTPUT_PATH = chart.CHART_OUTPUT_PATH.with_suffix(".png")
chart.FILE_CHART_OUTPUT_PATH = (
    DATA_GOV_ANALYSIS_DIR
    / "charts"
    / "combined_chart_distribution_of_size_reduction_by_file_rust.pdf"
)
chart.FILE_CHART_PREVIEW_OUTPUT_PATH = chart.FILE_CHART_OUTPUT_PATH.with_suffix(".png")
chart.BLOATIEST_EXAMPLES_OUTPUT_PATH = (
    DATA_GOV_ANALYSIS_DIR / "bloatiest_reduction_examples_rust.csv"
)


if __name__ == "__main__":
    chart.main()
