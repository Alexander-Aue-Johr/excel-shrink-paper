from __future__ import annotations

import gzip
import os
import re
import math
import shutil
from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.gridspec import GridSpec


INPUT_PATH = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "excel_shrink_rust_analysis.csv.gz",
)

CHART_OUTPUT_PATH = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "charts",
    "combined_chart_distribution_of_size_reduction_rust.pdf",
)

CHART_PREVIEW_OUTPUT_PATH = CHART_OUTPUT_PATH.with_suffix(".png")

FILE_CHART_OUTPUT_PATH = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "charts",
    "combined_chart_distribution_of_size_reduction_by_file_rust.pdf",
)

FILE_CHART_PREVIEW_OUTPUT_PATH = FILE_CHART_OUTPUT_PATH.with_suffix(".png")

BLOATIEST_EXAMPLES_OUTPUT_PATH = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "bloatiest_reduction_examples_rust.csv",
)


@dataclass(frozen=True)
class CausePanel:
    title: str
    causes: tuple[str, ...]
    color: str


CAUSE_PANELS = [
    CausePanel(
        "Empty Cells",
        (
            "PreprocessEmptySelfClosingCells",
            "EmptyCellsWithoutContent",
            "EmptyHiddenCells",
        ),
        "#E69F00",
    ),
    CausePanel(
        "Empty Rows",
        ("EmptyRowsWithoutVisibleStyleOrHeight",),
        "#EDC948",
    ),
    CausePanel("Trailing Layout Rows", ("TrailingLayoutOnlyRows",), "#D37295"),
    CausePanel(
        "Empty Row Attributes",
        ("EmptyRowAttributeCleanup",),
        "#9C755F",
    ),
    CausePanel("Column Range Merge", ("ExtraneousColumnRangeMerge",), "#D55E00"),
    CausePanel(
        "Unreferenced Whitespace",
        ("UnreferencedWhitespaceCells",),
        "#009E73",
    ),
    CausePanel(
        "Unreferenced Hidden Content",
        ("UnreferencedHiddenContentCells",),
        "#59A14F",
    ),
    CausePanel(
        "Hidden Empty Rows",
        (
            "PreprocessConsecutiveHiddenEmptyRows",
            "ConsecutiveHiddenEmptyRows",
        ),
        "#76B7B2",
    ),
]


PANEL_RATIO_THRESHOLD = 1e-6

PANEL_DISPLAY_TITLES = {
    "Empty Cells": "Empty\nCells",
    "Unreferenced Whitespace": "Unreferenced\nWhitespace",
    "Unreferenced Hidden Content": "Unreferenced\nHidden Content",
    "Hidden Empty Rows": "Hidden\nEmpty Rows",
    "Empty Rows": "Empty\nRows",
    "Trailing Layout Rows": "Trailing\nLayout Rows",
    "Empty Row Attributes": "Empty Row\nAttributes",
    "Column Range Merge": "Column Range\nMerge",
}


def normalized_column_name(column: str) -> str:
    return re.sub(r"[^a-z0-9]", "", column.lower())


def find_column(data: pd.DataFrame, candidates: list[str]) -> str | None:
    columns_by_normalized_name = {
        normalized_column_name(column): column for column in data.columns
    }

    for candidate in candidates:
        column = columns_by_normalized_name.get(normalized_column_name(candidate))
        if column is not None:
            return column

    return None


def refresh_gzip_from_csv(gzip_path: Path) -> Path:
    csv_path = gzip_path.with_suffix("")
    if csv_path.exists():
        os.makedirs(gzip_path.parent, exist_ok=True)
        with csv_path.open("rb") as source, gzip.open(gzip_path, "wb") as target:
            shutil.copyfileobj(source, target)
        print(f"Updated compressed analysis CSV: {gzip_path}")

    return gzip_path


def read_rust_analysis(path: Path) -> pd.DataFrame:
    path = refresh_gzip_from_csv(path)
    data = pd.read_csv(path)

    numeric_columns = [
        "removed_bytes_uncompressed_xml",
        "added_bytes_uncompressed_xml",
        "original_sheet_xml_bytes",
        "cleaned_sheet_xml_bytes",
        "reduction_bytes_uncompressed_xml",
        "reduction_bytes",
        "reduction_ratio",
        "reduction ratio",
        "size_reduction_ratio",
    ]

    for column in numeric_columns:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0)

    return data


def build_long_format_reductions(data: pd.DataFrame) -> pd.DataFrame:
    result = data.copy()

    reduction_bytes_column = find_column(
        result,
        [
            "reduction_bytes_uncompressed_xml",
            "reduction_bytes",
            "size_reduction_bytes",
        ],
    )

    if reduction_bytes_column is not None:
        result["cause_reduction_bytes"] = result[reduction_bytes_column]
    else:
        result["cause_reduction_bytes"] = (
            result["removed_bytes_uncompressed_xml"]
            - result["added_bytes_uncompressed_xml"]
        )

    result["cause_reduction_bytes"] = result["cause_reduction_bytes"].clip(lower=0)

    reduction_ratio_column = find_column(
        result,
        [
            "reduction_ratio",
            "reduction ratio",
            "size_reduction_ratio",
            "cause_reduction_ratio",
        ],
    )

    if reduction_ratio_column is not None:
        result["cause_reduction_ratio"] = result[reduction_ratio_column]
    else:
        original_size = result["original_sheet_xml_bytes"]
        result["cause_reduction_ratio"] = (
            result["cause_reduction_bytes"] / original_size
        ).where(original_size > 0, 0)

    return result


def build_distribution_data(data: pd.DataFrame) -> pd.DataFrame:
    group_keys = ["input_file", "sheet_file"]
    long_data = build_long_format_reductions(data)

    sheet_sizes = (
        long_data.groupby(group_keys, as_index=False)
        .agg(
            original_size=("original_sheet_xml_bytes", "max"),
            cleaned_size=("cleaned_sheet_xml_bytes", "min"),
        )
    )

    cause_ratios = (
        long_data.pivot_table(
            index=group_keys,
            columns="cause",
            values="cause_reduction_ratio",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    result = sheet_sizes.merge(cause_ratios, on=group_keys, how="left")

    original_size = result["original_size"]
    result["Reduction Ratio"] = (
        (result["original_size"] - result["cleaned_size"]) / original_size
    ).where(original_size > 0, 0)

    for panel in CAUSE_PANELS:
        present_causes = [cause for cause in panel.causes if cause in result.columns]
        if present_causes:
            result[panel.title] = result[present_causes].sum(axis=1)
        else:
            result[panel.title] = 0

    result["File"] = result["input_file"].map(lambda value: Path(value).name)
    result["Sheet / Workbook"] = result["sheet_file"].map(
        lambda value: "Workbook" if value == "xl/workbook.xml" else value
    )

    ratio_columns = ["Reduction Ratio", *(panel.title for panel in CAUSE_PANELS)]
    result[ratio_columns] = result[ratio_columns].clip(lower=0, upper=1)

    return result


def build_file_distribution_data(data: pd.DataFrame) -> pd.DataFrame:
    long_data = build_long_format_reductions(data)

    sheet_sizes = (
        long_data.groupby(["input_file", "sheet_file"], as_index=False)
        .agg(
            original_size=("original_sheet_xml_bytes", "max"),
            cleaned_size=("cleaned_sheet_xml_bytes", "min"),
        )
    )

    file_sizes = (
        sheet_sizes.groupby("input_file", as_index=False)
        .agg(
            original_size=("original_size", "sum"),
            cleaned_size=("cleaned_size", "sum"),
        )
    )

    cause_bytes = (
        long_data.pivot_table(
            index="input_file",
            columns="cause",
            values="cause_reduction_bytes",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    result = file_sizes.merge(cause_bytes, on="input_file", how="left")
    original_size = result["original_size"]
    result["Reduction Ratio"] = (
        (result["original_size"] - result["cleaned_size"]) / original_size
    ).where(original_size > 0, 0)

    for panel in CAUSE_PANELS:
        present_causes = [cause for cause in panel.causes if cause in result.columns]
        if present_causes:
            result[panel.title] = (
                result[present_causes].sum(axis=1) / original_size
            ).where(original_size > 0, 0)
        else:
            result[panel.title] = 0

    result["File"] = result["input_file"].map(lambda value: Path(value).name)

    ratio_columns = ["Reduction Ratio", *(panel.title for panel in CAUSE_PANELS)]
    result[ratio_columns] = result[ratio_columns].clip(lower=0, upper=1)

    return result


def build_sheet_example_data(data: pd.DataFrame) -> pd.DataFrame:
    group_keys = ["input_file", "sheet_file"]
    long_data = build_long_format_reductions(data)

    result = (
        long_data.groupby(group_keys, as_index=False)
        .agg(
            original_size=("original_sheet_xml_bytes", "max"),
            cleaned_size=("cleaned_sheet_xml_bytes", "min"),
        )
    )

    cause_bytes = (
        long_data.pivot_table(
            index=group_keys,
            columns="cause",
            values="cause_reduction_bytes",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    result = result.merge(cause_bytes, on=group_keys, how="left")

    result["overall_reduction_bytes"] = (
        result["original_size"] - result["cleaned_size"]
    ).clip(lower=0)
    original_size = result["original_size"]
    result["overall_reduction_ratio"] = (
        result["overall_reduction_bytes"] / original_size
    ).where(original_size > 0, 0)

    for panel in CAUSE_PANELS:
        present_causes = [cause for cause in panel.causes if cause in result.columns]
        byte_column = f"{panel.title} bytes"
        ratio_column = f"{panel.title} ratio"
        if present_causes:
            result[byte_column] = result[present_causes].sum(axis=1).clip(lower=0)
            result[ratio_column] = (result[byte_column] / original_size).where(
                original_size > 0, 0
            )
        else:
            result[byte_column] = 0
            result[ratio_column] = 0

    result["File"] = result["input_file"].map(lambda value: Path(value).name)
    result["Sheet / Workbook"] = result["sheet_file"].map(
        lambda value: "Workbook" if value == "xl/workbook.xml" else value
    )

    return result


def build_file_example_data(data: pd.DataFrame) -> pd.DataFrame:
    sheet_data = build_sheet_example_data(data)

    file_sizes = (
        sheet_data.groupby("input_file", as_index=False)
        .agg(
            original_size=("original_size", "sum"),
            cleaned_size=("cleaned_size", "sum"),
            sheet_or_workbook_count=("sheet_file", "size"),
            mean_sheet_original_size=("original_size", "mean"),
            mean_sheet_cleaned_size=("cleaned_size", "mean"),
            mean_sheet_reduction_bytes=("overall_reduction_bytes", "mean"),
            mean_sheet_reduction_ratio=("overall_reduction_ratio", "mean"),
        )
    )

    long_data = build_long_format_reductions(data)
    cause_bytes = (
        long_data.pivot_table(
            index="input_file",
            columns="cause",
            values="cause_reduction_bytes",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )
    result = file_sizes.merge(cause_bytes, on="input_file", how="left")

    result["overall_reduction_bytes"] = (
        result["original_size"] - result["cleaned_size"]
    ).clip(lower=0)
    original_size = result["original_size"]
    result["overall_reduction_ratio"] = (
        result["overall_reduction_bytes"] / original_size
    ).where(original_size > 0, 0)

    for panel in CAUSE_PANELS:
        present_causes = [cause for cause in panel.causes if cause in result.columns]
        byte_column = f"{panel.title} bytes"
        ratio_column = f"{panel.title} ratio"
        if present_causes:
            result[byte_column] = result[present_causes].sum(axis=1).clip(lower=0)
            result[ratio_column] = (result[byte_column] / original_size).where(
                original_size > 0, 0
            )
        else:
            result[byte_column] = 0
            result[ratio_column] = 0

    result["File"] = result["input_file"].map(lambda value: Path(value).name)
    result["Sheet / Workbook"] = ""

    return result


def _example_row(
    *,
    scope: str,
    category: str,
    category_causes: tuple[str, ...],
    metric: str,
    row: pd.Series,
    bytes_column: str,
    ratio_column: str,
) -> dict[str, object]:
    output = {
        "scope": scope,
        "category": category,
        "category_causes": ";".join(category_causes),
        "metric": metric,
        "input_file": row.get("input_file", ""),
        "file": row.get("File", ""),
        "sheet_file": row.get("sheet_file", ""),
        "sheet_or_workbook": row.get("Sheet / Workbook", ""),
        "original_size_bytes": row.get("original_size", 0),
        "cleaned_size_bytes": row.get("cleaned_size", 0),
        "reduction_bytes": row.get(bytes_column, 0),
        "reduction_ratio": row.get(ratio_column, 0),
        "reduction_percent": row.get(ratio_column, 0) * 100,
    }

    optional_columns = [
        "sheet_or_workbook_count",
        "mean_sheet_original_size",
        "mean_sheet_cleaned_size",
        "mean_sheet_reduction_bytes",
        "mean_sheet_reduction_ratio",
    ]
    for column in optional_columns:
        output[column] = row.get(column, "")

    return output


def _append_top_examples(
    rows: list[dict[str, object]],
    data: pd.DataFrame,
    *,
    scope: str,
    category: str,
    category_causes: tuple[str, ...],
    bytes_column: str,
    ratio_column: str,
) -> None:
    if data.empty:
        return

    absolute_data = data[data[bytes_column] > 0]
    if not absolute_data.empty:
        absolute_row = absolute_data.sort_values(
            [bytes_column, ratio_column, "original_size"],
            ascending=[False, False, False],
        ).iloc[0]
        rows.append(
            _example_row(
                scope=scope,
                category=category,
                category_causes=category_causes,
                metric="absolute_bytes",
                row=absolute_row,
                bytes_column=bytes_column,
                ratio_column=ratio_column,
            )
        )

    relative_data = data[data[ratio_column] > PANEL_RATIO_THRESHOLD]
    if not relative_data.empty:
        relative_row = relative_data.sort_values(
            [ratio_column, bytes_column, "original_size"],
            ascending=[False, False, False],
        ).iloc[0]
        rows.append(
            _example_row(
                scope=scope,
                category=category,
                category_causes=category_causes,
                metric="relative_ratio",
                row=relative_row,
                bytes_column=bytes_column,
                ratio_column=ratio_column,
            )
        )


def build_bloatiest_examples(data: pd.DataFrame) -> pd.DataFrame:
    sheet_data = build_sheet_example_data(data)
    file_data = build_file_example_data(data)
    rows: list[dict[str, object]] = []

    categories: list[tuple[str, tuple[str, ...], str, str]] = [
        (
            "Overall Reduction",
            (),
            "overall_reduction_bytes",
            "overall_reduction_ratio",
        ),
        *(
            (
                panel.title,
                panel.causes,
                f"{panel.title} bytes",
                f"{panel.title} ratio",
            )
            for panel in CAUSE_PANELS
        ),
    ]

    for category, category_causes, bytes_column, ratio_column in categories:
        _append_top_examples(
            rows,
            sheet_data,
            scope="sheet_or_workbook",
            category=category,
            category_causes=category_causes,
            bytes_column=bytes_column,
            ratio_column=ratio_column,
        )
        _append_top_examples(
            rows,
            file_data,
            scope="excel_file",
            category=category,
            category_causes=category_causes,
            bytes_column=bytes_column,
            ratio_column=ratio_column,
        )

    result = pd.DataFrame(rows)
    if "mean_sheet_reduction_ratio" in result.columns:
        result["mean_sheet_reduction_percent"] = pd.to_numeric(
            result["mean_sheet_reduction_ratio"], errors="coerce"
        ) * 100

    return result


def save_bloatiest_examples(data: pd.DataFrame, output_path: Path) -> pd.DataFrame:
    examples = build_bloatiest_examples(data)
    os.makedirs(output_path.parent, exist_ok=True)
    examples.to_csv(output_path, index=False)
    return examples


def report_cause_overview(data: pd.DataFrame) -> None:
    long_data = build_long_format_reductions(data)
    overview = (
        long_data.groupby("cause", as_index=False)
        .agg(
            rows=("cause", "size"),
            removed_bytes=("cause_reduction_bytes", "sum"),
            mean_ratio=("cause_reduction_ratio", "mean"),
            max_ratio=("cause_reduction_ratio", "max"),
        )
        .sort_values("removed_bytes", ascending=False)
    )

    print(f"Cause count: {len(overview)}")
    print(overview.to_string(index=False))


def nice_x_limit(nonzero_count: int, total_count: int) -> int:
    if total_count <= 1:
        return total_count

    buffered_count = max(1, math.ceil(nonzero_count * 1.05))

    if buffered_count <= 1_000:
        step = 100
    elif buffered_count <= 5_000:
        step = 500
    elif buffered_count <= 10_000:
        step = 1_000
    else:
        step = 5_000

    return min(total_count, max(1, math.ceil(buffered_count / step) * step))


def x_ticks_for_limit(x_limit: int) -> list[int]:
    if x_limit <= 1:
        return [1]

    midpoint = max(1, int(round(x_limit / 2)))
    return sorted({1, midpoint, x_limit})


def axis_padding_for_limit(x_limit: int) -> int:
    return max(5, math.ceil(x_limit * 0.03))


def plot_distribution(
    data: pd.DataFrame,
    output_path: Path,
    preview_output_path: Path,
    title: str = "Rust Combined Reduction Ratio",
    x_axis_label: str = "Excel Sheets / Workbook Entries (Index)",
) -> None:
    full_axis_limit = len(data)
    half_axis_limit = max(1, math.ceil(full_axis_limit / 2))
    quarter_axis_limit = max(1, math.ceil(full_axis_limit / 4))
    half_axis_panels = {
        "Empty Rows",
        "Empty Row Attributes",
        "Trailing Layout Rows",
    }
    panel_axis_limits = {
        panel.title: (
            full_axis_limit
            if panel.title == "Empty Cells"
            else half_axis_limit
            if panel.title in half_axis_panels
            else quarter_axis_limit
        )
        for panel in CAUSE_PANELS
    }
    panel_width_ratios = [
        4 if panel.title == "Empty Cells" else 2 if panel.title in half_axis_panels else 1
        for panel in CAUSE_PANELS
    ]

    fig = plt.figure(figsize=(28, 8.5))
    gs = GridSpec(
        2,
        len(CAUSE_PANELS),
        height_ratios=[3.3, 2.4],
        width_ratios=panel_width_ratios,
        figure=fig,
    )

    data_sorted = data.sort_values(by="Reduction Ratio", ascending=False)
    x_ticks = range(1, len(data_sorted) + 1)

    x_buffer = axis_padding_for_limit(len(data_sorted))
    y_buffer = 10

    ax_combined = fig.add_subplot(gs[0, :])
    ax_combined.fill_between(
        x_ticks,
        data_sorted["Reduction Ratio"] * 100,
        color="#A7C6ED",
        label="Reduction Ratio",
    )
    ax_combined.set_title(title, fontsize=18)
    ax_combined.grid(True, axis="y", linestyle="--", color="gray", alpha=0.6)

    if len(data_sorted) <= 2_000:
        tick_positions = x_ticks_for_limit(len(data_sorted))
    else:
        tick_positions = [1] + list(range(2000, len(data_sorted) + 1, 2000))
    ax_combined.set_xticks(tick_positions)
    ax_combined.set_xticklabels(tick_positions, rotation=45, ha="right")
    ax_combined.set_xlim(1 - x_buffer, len(data_sorted) + x_buffer)

    axs = []
    for index, panel in enumerate(CAUSE_PANELS):
        ax = fig.add_subplot(gs[1, index])

        panel_sorted = data.sort_values(by=panel.title, ascending=False).reset_index(
            drop=True
        )
        panel_x_ticks = range(1, len(panel_sorted) + 1)
        x_limit = panel_axis_limits[panel.title]
        x_padding = axis_padding_for_limit(x_limit)

        ax.fill_between(
            panel_x_ticks,
            panel_sorted[panel.title] * 100,
            color=panel.color,
        )
        ax.set_title(PANEL_DISPLAY_TITLES[panel.title], fontsize=10)
        ax.grid(True, axis="y", linestyle="--", color="gray", alpha=0.6)
        ax.set_ylim(-y_buffer, 100 + y_buffer)
        ax.set_xlim(1 - x_padding, x_limit + x_padding)
        ax.set_yticks([0, 20, 40, 60, 80, 100])
        ax.set_xticks(x_ticks_for_limit(x_limit))
        ax.set_xticklabels(x_ticks_for_limit(x_limit), rotation=45, ha="right")

        if index > 0:
            ax.set_yticklabels([])

        axs.append(ax)

    fig.text(
        0.5,
        0.035,
        x_axis_label,
        ha="center",
        fontsize=13,
    )
    fig.text(
        0.015,
        0.5,
        "Size Reduction (%)",
        va="center",
        rotation="vertical",
        fontsize=13,
    )

    plt.tight_layout(rect=[0.03, 0.08, 0.99, 0.96])
    plt.subplots_adjust(hspace=0.45, wspace=0.22)

    os.makedirs(output_path.parent, exist_ok=True)
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.savefig(preview_output_path, dpi=180, bbox_inches="tight")
    plt.close()


def main() -> None:
    data = read_rust_analysis(INPUT_PATH)
    report_cause_overview(data)

    distribution_data = build_distribution_data(data)
    plot_distribution(distribution_data, CHART_OUTPUT_PATH, CHART_PREVIEW_OUTPUT_PATH)

    file_distribution_data = build_file_distribution_data(data)
    plot_distribution(
        file_distribution_data,
        FILE_CHART_OUTPUT_PATH,
        FILE_CHART_PREVIEW_OUTPUT_PATH,
        title="Rust Combined Reduction Ratio by Excel File",
        x_axis_label="Excel Files (Index)",
    )

    bloatiest_examples = save_bloatiest_examples(data, BLOATIEST_EXAMPLES_OUTPUT_PATH)

    print(f"Read rows: {len(data)}")
    print(f"Plotted sheet/workbook entries: {len(distribution_data)}")
    print(f"Plotted file entries: {len(file_distribution_data)}")
    print(f"Saved bloatiest examples: {BLOATIEST_EXAMPLES_OUTPUT_PATH}")
    print(f"Bloatiest example rows: {len(bloatiest_examples)}")
    print(f"Saved chart: {CHART_OUTPUT_PATH}")
    print(f"Saved preview: {CHART_PREVIEW_OUTPUT_PATH}")
    print(f"Saved file chart: {FILE_CHART_OUTPUT_PATH}")
    print(f"Saved file preview: {FILE_CHART_PREVIEW_OUTPUT_PATH}")


if __name__ == "__main__":
    main()
