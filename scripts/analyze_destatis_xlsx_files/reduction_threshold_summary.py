import os
import pandas as pd


INPUT_PATH = os.path.join(
    "scripts",
    "analyze_destatis_xlsx_files",
    "analysis_output.csv",
)

OUTPUT_PATH = os.path.join(
    "scripts",
    "analyze_destatis_xlsx_files",
    "reduction_threshold_summary.csv",
)

THRESHOLDS = [
    0.50,
    0.60,
    0.70,
    0.80,
    0.90,
    0.95,
    0.99,
    0.999,
    0.9999,
]


def build_threshold_summary(
    data: pd.DataFrame,
    group_name: str,
    ratio_column: str = "Reduction Ratio",
) -> pd.DataFrame:
    total_rows = len(data)

    rows = []

    for threshold in THRESHOLDS:
        count = (data[ratio_column] >= threshold).sum()
        share = count / total_rows * 100 if total_rows > 0 else 0.0

        rows.append(
            {
                "Group": group_name,
                "Total Rows": total_rows,
                "Threshold": f">= {threshold * 100:.2f}%",
                "Row Count": int(count),
                "Row Share (%)": share,
            }
        )

    return pd.DataFrame(rows)


def print_group_overview(
    group_name: str,
    data: pd.DataFrame,
) -> None:
    print()
    print("=" * 80)
    print(group_name)
    print("=" * 80)
    print(f"Rows: {len(data)}")
    print(f"Unique files: {data['File'].nunique()}")

    if len(data) > 0:
        print(f"Rows with Original Size > 0: {(data['Original Size'] > 0).sum()}")
        print(f"Rows with Original Size == 0: {(data['Original Size'] == 0).sum()}")
        print(f"Rows with Original Size < 0: {(data['Original Size'] < 0).sum()}")

        print()
        print("Most frequent values in 'Sheet / Workbook':")
        print(data["Sheet / Workbook"].value_counts().head(10).to_string())


def main() -> None:
    data = pd.read_csv(INPUT_PATH)

    # Make sure ratio and size columns are numeric.
    numeric_columns = [
        "Original Size",
        "Cleaned Size",
        "Reduction Ratio",
        "Empty Cell Reduction Ratio",
        "Whitespace Cell Reduction Ratio",
        "Rows Outside of Value Range Reduction Ratio",
        "Removed Columns Reduction Ratio",
    ]

    for column in numeric_columns:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0)

    worksheets_only = data[
        data["Sheet / Workbook"] != "Workbook"
    ].copy()

    non_worksheets_only = data[
        data["Sheet / Workbook"] == "Workbook"
    ].copy()

    all_rows = data.copy()

    groups = [
        ("worksheets_only", worksheets_only),
        ("non_worksheets_only", non_worksheets_only),
        ("all_rows", all_rows),
    ]

    summaries = []

    for group_name, group_data in groups:
        print_group_overview(group_name, group_data)

        summary = build_threshold_summary(
            data=group_data,
            group_name=group_name,
            ratio_column="Reduction Ratio",
        )

        summaries.append(summary)

        print()
        print(
            summary[
                ["Threshold", "Row Count", "Row Share (%)"]
            ].to_string(
                index=False,
                formatters={
                    "Row Share (%)": "{:.2f}".format,
                },
            )
        )

    combined_summary = pd.concat(summaries, ignore_index=True)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    combined_summary.to_csv(OUTPUT_PATH, index=False)

    print()
    print("=" * 80)
    print("Saved combined threshold summary")
    print("=" * 80)
    print(OUTPUT_PATH)

    print()
    print("Sanity checks")
    print("-" * 80)
    print(f"All CSV data rows excluding header: {len(data)}")
    print(f"Worksheet rows: {len(worksheets_only)}")
    print(f"Non-worksheet rows: {len(non_worksheets_only)}")
    print(
        "Worksheet rows + non-worksheet rows == all rows:",
        len(worksheets_only) + len(non_worksheets_only) == len(data),
    )
    print(f"Unique files: {data['File'].nunique()}")


if __name__ == "__main__":
    main()