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
    parser.add_argument("--common-inputs", action="store_true",
                        help="Restrict each host/corpus group to input identities shared by its tools")
    args = parser.parse_args()
    rows = read_pairs(args.pairs)
    if args.common_inputs:
        rows = shared_input_pairs(rows)
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


def shared_input_pairs(rows):
    groups = {}
    for row in rows:
        tools = groups.setdefault((row["host_id"], row["corpus"]), {})
        tools.setdefault(row["comparator"], set()).add(row["input_sha256"])
    common = {key: set.intersection(*tools.values()) for key, tools in groups.items()}
    selected = [row for row in rows if row["input_sha256"] in common[(row["host_id"], row["corpus"])]]
    if not selected:
        raise ValueError("The supplied tools have no common input identities")
    return selected


def draw_comparison(rows, output):
    groups = list(dict.fromkeys((row["host_id"], row["comparator"], row["corpus"]) for row in rows))
    tool_names = list(dict.fromkeys(group[1] for group in groups))
    hosts = list(dict.fromkeys(group[0] for group in groups))
    multiple_tools = len(tool_names) > 1
    multiple_hosts = len(hosts) > 1
    corpus_order = {"destatis": 0, "enron": 1, "datagov": 2}
    groups.sort(key=lambda group: (group[0], corpus_order.get(group[2], 3), group[2], tool_names.index(group[1])))
    markers = ["s", "^", "D", "v"]
    colors = ["#D55E00", "#009E73", "#CC79A7", "#E69F00"]
    styles = {name: (markers[index % 4], colors[index % 4]) for index, name in enumerate(tool_names)}
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8,
                         "pdf.fonttype": 42, "svg.fonttype": "none"})
    figure, axes = plt.subplots(1, 2, figsize=(7.15, 1.8 + 0.47 * len(groups)))
    labels = []
    for position, group in enumerate(groups):
        selected = [row for row in rows if (row["host_id"], row["comparator"], row["corpus"]) == group]
        original = np.array([row["comparator_seconds"] for row in selected])
        cleaned = np.array([row["sheetshrink_seconds"] for row in selected])
        ratios = original / cleaned
        marker, color = styles[group[1]]
        corpus_name = {"destatis": "Destatis", "enron": "Enron", "datagov": "Data.gov"}.get(group[2], group[2])
        label = f"{corpus_name}  n={len(selected)}"
        if multiple_tools:
            label = f"{group[1]}\n{label}"
        if multiple_hosts:
            label = f"{group[0]}\n{label}"
        labels.append(label)
        draw_time_summary(axes[0], position - 0.09, original, marker, color, 7)
        draw_time_summary(axes[0], position + 0.09, cleaned, "o", "#0072B2", -10)
        jitter = np.linspace(-0.17, 0.17, len(ratios))
        np.random.default_rng(20261110).shuffle(jitter)
        axes[1].scatter(ratios, position + jitter, s=11, color=color, marker=marker, alpha=0.45, linewidths=0)
        quartiles = np.percentile(ratios, [25, 50, 75])
        axes[1].plot(quartiles[[0, 2]], [position, position], color="black", linewidth=2)
        axes[1].scatter(quartiles[1], position, color="black", marker="D", s=20, zorder=4)
        axes[1].annotate(f"{quartiles[1]:.1f}x", (quartiles[1], position),
                         xytext=(4, 8), textcoords="offset points", fontsize=7,
                         bbox={"facecolor": "white", "edgecolor": "none", "alpha": 0.8, "pad": 0.5})
    format_axes(axes, labels)
    heading = f"Cleanup workflow comparisons  ·  {', '.join(hosts)}"
    figure.text(0.02, 0.98, heading, va="top", fontsize=8)
    for name, (marker, color) in styles.items():
        axes[0].scatter([], [], color=color, marker=marker, label=name)
    axes[0].scatter([], [], color="#0072B2", marker="o", label="SheetShrink")
    figure.legend(*axes[0].get_legend_handles_labels(), frameon=False, loc="upper left",
                  bbox_to_anchor=(0.20, 0.93), ncol=len(tool_names) + 1, fontsize=7)
    figure.text(0.02, 0.025, "Absolute times show workbook medians and quartiles. Ratio dots represent paired workbooks.\n"
                "Black diamonds and lines mark the median and IQR across workbooks.\n"
                "Each external tool n=1. SheetShrink n=9 after warm-up. Native visual validation pending.", fontsize=6.8)
    figure.subplots_adjust(left=0.29 if multiple_hosts else 0.23, right=0.985, top=0.80, bottom=0.20, wspace=0.22)
    output.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output)
    figure.savefig(output.with_suffix(".png"), dpi=200)
    plt.close(figure)


def draw_time_summary(axis, position, values, marker, color, label_offset):
    low, middle, high = np.percentile(values, [25, 50, 75])
    axis.plot([low, high], [position, position], color=color, linewidth=1.7)
    axis.scatter(middle, position, color=color, marker=marker, s=24, zorder=3)
    axis.annotate(f"{middle:.2g} s", (middle, position), xytext=(4, label_offset),
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
    axes[1].set_xlabel("Tool / SheetShrink time ratio")
    axes[1].axvline(1, color="#999999", linestyle="--", linewidth=0.8)


if __name__ == "__main__":
    main()
