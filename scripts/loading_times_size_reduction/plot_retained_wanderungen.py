"""Plot retained CSV observations without loading a workbook or benchmark module."""

import argparse
import csv
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt


WORKBOOK = "wanderungen-2010120217005.xlsx"


def main():
    arguments = parse_arguments()
    observations = read_retained_observations(arguments.csv)
    draw_observations(observations, arguments.output)


def parse_arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", type=Path, default=Path(__file__).with_name(
        "excel_library_execution_times.csv"))
    parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args()


def read_retained_observations(source):
    with source.open(encoding="utf-8-sig", newline="") as stream:
        rows = [row for row in csv.DictReader(stream) if row["File"] == WORKBOOK]
    if len(rows) != 5 or len({row["Library"] for row in rows}) != 5:
        raise ValueError("Expected five distinct retained library observations")
    return rows


def completed_duration(row, prefix):
    if row[f"{prefix} Status"] != "ok":
        raise ValueError("A completed duration requires a successful observation")
    duration = sum(float(row[f"{prefix} {operation} Time (s)"])
                   for operation in ("Open", "Save"))
    if duration <= 0:
        raise ValueError("Logarithmic time axes require positive durations")
    return duration


def draw_observations(rows, output):
    plt.rcParams.update({"font.size": 10, "pdf.fonttype": 42})
    figure, axis = plt.subplots(figsize=(8.2, 4.2))
    for position, row in enumerate(rows):
        cleaned = completed_duration(row, "Shrinked")
        axis.scatter(cleaned, position, color="#16735b", marker="o", zorder=3)
        axis.annotate(f"{cleaned:.2f} s", (cleaned, position),
                      xytext=(0, -15), textcoords="offset points", ha="center", fontsize=8)
        if row["Original Status"] == "failed_timeout":
            bound = float(row["Original Failure Time (s)"])
            axis.scatter(bound, position, color="#b05a17", marker=">", zorder=3)
            axis.annotate(f">= {bound:.1f} s (timeout)", (bound, position),
                          xytext=(-6, 9), textcoords="offset points", ha="right", fontsize=8)
        else:
            duration = completed_duration(row, "Original")
            axis.scatter(duration, position, facecolor="white", edgecolor="#285c92", zorder=3)
            axis.annotate(f"{duration:.2f} s", (duration, position),
                          xytext=(0, 9), textcoords="offset points", ha="center", fontsize=8)
    axis.set_xscale("log")
    axis.set_xlim(0.3, 1600)
    axis.set_yticks(range(len(rows)), [row["Library"] for row in rows])
    axis.set_ylim(-0.6, len(rows) - 0.4)
    axis.invert_yaxis()
    axis.set_xlabel("Retained opening + saving time in seconds (logarithmic scale)")
    axis.grid(axis="x", which="major", color="#dddddd", linewidth=0.6)
    axis.set_axisbelow(True)
    for spine in ("top", "right", "left"):
        axis.spines[spine].set_visible(False)
    figure.suptitle("Historical Wanderungen observations", x=0.32, ha="left", fontsize=12)
    figure.text(0.08, 0.025, "Filled green = cleaned; outlined blue = original; orange = censored.\n"
                "Historical environment and producer identity remain unreconciled. Cleanup time excluded.",
                fontsize=8)
    figure.subplots_adjust(left=0.32, right=0.98, bottom=0.25, top=0.85)
    output.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output)
    figure.savefig(output.with_suffix(".png"), dpi=180)
    plt.close(figure)


if __name__ == "__main__":
    main()
