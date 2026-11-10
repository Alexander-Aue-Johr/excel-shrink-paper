"""Plot retained paired cleanup observations without opening any workbook."""
import argparse
import csv
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pairs", type=Path, nargs="+", required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    rows = read_pairs(args.pairs)
    draw_comparison(rows, args.output)


def read_pairs(paths):
    rows = []
    identities = set()
    for path in paths:
        with path.open(encoding="utf-8-sig", newline="") as stream:
            for row in csv.DictReader(stream):
                identity = (row["host_id"], row["comparator"], row["corpus"], row["input_sha256"])
                if identity in identities:
                    raise ValueError(f"Duplicate paired observation: {identity}")
                identities.add(identity)
                row["sheetshrink_seconds"] = float(row["sheetshrink_seconds"])
                row["comparator_seconds"] = float(row["comparator_seconds"])
                if min(row["sheetshrink_seconds"], row["comparator_seconds"]) <= 0:
                    raise ValueError("Logarithmic plots require positive times")
                rows.append(row)
    if not rows:
        raise ValueError("No paired observations were supplied")
    return rows


def draw_comparison(rows, output):
    groups = list(dict.fromkeys((row["host_id"], row["comparator"], row["corpus"]) for row in rows))
    multiple_lanes = len({group[:2] for group in groups}) > 1
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8,
                         "pdf.fonttype": 42, "svg.fonttype": "none"})
    figure, axes = plt.subplots(1, 2, figsize=(7.15, 1.8 + 0.47 * len(groups)))
    colors = {"destatis": "#0072B2", "enron": "#D55E00", "datagov": "#009E73"}
    labels = []
    for position, group in enumerate(groups):
        selected = [row for row in rows if (row["host_id"], row["comparator"], row["corpus"]) == group]
        original = np.array([row["comparator_seconds"] for row in selected])
        cleaned = np.array([row["sheetshrink_seconds"] for row in selected])
        ratios = original / cleaned
        color = colors.get(group[2], "#555555")
        corpus_name = {"destatis": "Destatis", "enron": "Enron", "datagov": "Data.gov"}.get(group[2], group[2])
        label = f"{corpus_name}  n={len(selected)}"
        if multiple_lanes:
            label = f"{group[1]} / {group[0]}\n{label}"
        labels.append(label)
        draw_time_summary(axes[0], position - 0.09, original, "s", "#555555")
        draw_time_summary(axes[0], position + 0.09, cleaned, "o", "#0072B2")
        jitter = np.linspace(-0.17, 0.17, len(ratios))
        np.random.default_rng(20261110).shuffle(jitter)
        axes[1].scatter(ratios, position + jitter, s=11, color=color, alpha=0.45, linewidths=0)
        quartiles = np.percentile(ratios, [25, 50, 75])
        axes[1].plot(quartiles[[0, 2]], [position, position], color="black", linewidth=2)
        axes[1].scatter(quartiles[1], position, color="black", marker="D", s=20, zorder=4)
        axes[1].annotate(f"{quartiles[1]:.1f}x", (quartiles[1], position),
                         xytext=(4, 8), textcoords="offset points", fontsize=7,
                         bbox={"facecolor": "white", "edgecolor": "none", "alpha": 0.8, "pad": 0.5})
    format_axes(axes, labels)
    tools = ", ".join(dict.fromkeys(group[1] for group in groups))
    hosts = ", ".join(dict.fromkeys(group[0] for group in groups))
    heading = "Same-host cleanup workflow observations" if multiple_lanes else f"{tools} / SheetShrink  ·  {hosts}"
    figure.text(0.02, 0.98, heading, va="top", fontsize=8)
    axes[0].scatter([], [], color="#555555", marker="s", label="Comparator")
    axes[0].scatter([], [], color="#0072B2", marker="o", label="SheetShrink")
    axes[0].legend(frameon=False, loc="lower left", bbox_to_anchor=(-0.04, 1.01), ncol=2, fontsize=7)
    figure.text(0.02, 0.025, "Absolute times show workbook medians and quartiles. Ratio dots represent paired workbooks.\n"
                "Black diamonds and lines mark the median and IQR across workbooks.\n"
                "Comparator n=1. SheetShrink n=9 after warm-up. Native visual validation pending.", fontsize=6.8)
    figure.subplots_adjust(left=0.32 if multiple_lanes else 0.18, right=0.985, top=0.80, bottom=0.24, wspace=0.22)
    output.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output)
    figure.savefig(output.with_suffix(".png"), dpi=200)
    plt.close(figure)


def draw_time_summary(axis, position, values, marker, color):
    low, middle, high = np.percentile(values, [25, 50, 75])
    axis.plot([low, high], [position, position], color=color, linewidth=1.7)
    axis.scatter(middle, position, color=color, marker=marker, s=24, zorder=3)
    axis.annotate(f"{middle:.2g} s", (middle, position), xytext=(4, 7 if marker == "s" else -10),
                  textcoords="offset points", fontsize=7, color=color)


def format_axes(axes, labels):
    for axis in axes:
        axis.set_xscale("log")
        axis.set_ylim(len(labels) - 0.5, -0.5)
        axis.grid(axis="x", which="major", color="#dddddd", linewidth=0.5)
        axis.set_axisbelow(True)
        axis.tick_params(axis="y", length=0)
        for side in ("left", "right", "top"):
            axis.spines[side].set_visible(False)
    axes[0].set_yticks(range(len(labels)), labels)
    axes[1].set_yticks(range(len(labels)), [""] * len(labels))
    axes[0].set_xlabel("Observed cleanup workflow time (s)")
    axes[1].set_xlabel("Comparator / SheetShrink time ratio")
    axes[1].axvline(1, color="#999999", linestyle="--", linewidth=0.8)


if __name__ == "__main__":
    main()
