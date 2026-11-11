"""Plot retained paired cleanup observations without opening any workbook."""
import argparse
import csv
import gzip
import json
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
    parser.add_argument("--receipts", type=Path, nargs="+",
                        help="Complete per-tool selection receipts for outcome counts")
    parser.add_argument("--cohort-pairs", type=Path, nargs="+",
                        help="Restrict input hashes to the intersection of these reference pair files")
    parser.add_argument("--single-column", action="store_true",
                        help="Group tools by corpus in two compact timing panels")
    parser.add_argument("--a4-landscape", action="store_true",
                        help="Render the corpus-grouped chart on an A4 landscape page")
    parser.add_argument("--two-column", action="store_true",
                        help="Render the corpus-grouped chart at full two-column paper width")
    args = parser.parse_args()
    if args.two_column and (not args.single_column or args.a4_landscape):
        parser.error("--two-column requires --single-column and excludes --a4-landscape")
    if args.a4_landscape and not args.single_column:
        parser.error("--a4-landscape requires --single-column")
    rows = read_pairs(args.pairs)
    if args.cohort_pairs:
        cohorts = [{row["input_sha256"] for row in read_pairs([path])} for path in args.cohort_pairs]
        cohort = set.intersection(*cohorts)
        rows = [row for row in rows if row["input_sha256"] in cohort]
    if args.common_inputs:
        rows = shared_input_pairs(rows)
    outcomes = read_outcomes(args.receipts, rows) if args.receipts else None
    if args.single_column:
        draw_single_column(rows, args.output, args.a4_landscape, args.two_column)
    else:
        draw_comparison(rows, args.output, outcomes)


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


def read_outcomes(paths, rows):
    plotted = {(r["host_id"], r["comparator"], r["corpus"], r["input_sha256"]) for r in rows}
    outcomes = {}
    seen = set()
    for path in paths:
        with gzip.open(path, "rt", encoding="utf-8") as stream:
            receipt = json.load(stream)
        tool = receipt.get("tool_label", "NXPowerLite")
        for record in receipt["records"]:
            group = (receipt["host_id"], tool, record["corpus"])
            identity = (*group, record["input_sha"])
            if identity in seen:
                raise ValueError(f"Duplicate outcome identity: {identity}")
            seen.add(identity)
            counts = outcomes.setdefault(group, [0] * 6)
            if identity in plotted:
                category = 0
            elif record["conversion"] == "partial":
                category = 5
            elif record["conversion"] == "failed":
                category = 2
            elif record["eligible_for_formatting_comparison"]:
                category = 1
            elif tool == "NXPowerLite":
                category = 3
            else:
                category = 4
            counts[category] += 1
    if not plotted.issubset(seen):
        raise ValueError("Outcome receipts omit plotted identities")
    return outcomes


def draw_outcomes(axis, groups, outcomes):
    colors = ["#0072B2", "#56B4E9", "#D55E00", "#E69F00", "#CCCCCC", "#CC79A7"]
    hatches = ["", "//", "xx", "..", "\\\\", "++"]
    for position, group in enumerate(groups):
        counts = outcomes[group]
        left = 0
        for count, color, hatch in zip(counts, colors, hatches):
            axis.barh(position, count, left=left, height=0.44, color=color,
                      hatch=hatch, edgecolor="white", linewidth=0.4)
            left += count
        axis.text(0, position + 0.34, "/".join(str(n) for n in counts) + f" = {sum(counts)}", fontsize=6.5)
    axis.set_ylim(len(groups) - 0.5, -0.5)
    axis.set_xlim(0, max(sum(outcomes[g]) for g in groups))
    axis.set_xticks([0, max(sum(outcomes[g]) for g in groups)])
    axis.set_yticks([])
    axis.set_xlabel("Selection outcomes (files)")
    for side in ("left", "right", "top"):
        axis.spines[side].set_visible(False)


def draw_comparison(rows, output, outcomes=None):
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
    figure, axes = plt.subplots(1, 3 if outcomes else 2,
                                figsize=(7.15, 1.8 + 0.42 * len(groups)),
                                gridspec_kw={"width_ratios": [1, 1, 0.85]} if outcomes else None)
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
            host_label = {"LAPTOP-H86KBIH3": "Lenovo", "LV-ANB-001P-BS": "HP"}.get(group[0], group[0])
            label = f"{host_label} · {label}"
        labels.append(label)
        draw_time_summary(axes[0], position - 0.09, original, marker, color, 5)
        draw_time_summary(axes[0], position + 0.09, cleaned, "o", "#0072B2", -7)
        jitter = np.linspace(-0.17, 0.17, len(ratios))
        np.random.default_rng(20261110).shuffle(jitter)
        axes[1].scatter(ratios, position + jitter, s=11, color=color,
                        marker=marker, alpha=0.45, linewidths=0)
        quartiles = np.percentile(ratios, [25, 50, 75])
        axes[1].plot(quartiles[[0, 2]], [position, position], color="black", linewidth=2)
        axes[1].scatter(quartiles[1], position, color="black", marker="D", s=20, zorder=4)
        axes[1].annotate(f"{quartiles[1]:.1f}x", (quartiles[1], position),
                         xytext=(4, 8), textcoords="offset points", fontsize=7,
                         bbox={"facecolor": "white", "edgecolor": "none", "alpha": 0.8, "pad": 0.5})
    format_axes(axes[:2], labels)
    if outcomes:
        draw_outcomes(axes[2], groups, outcomes)
    heading = "Cleanup workflow observations by host and corpus"
    figure.text(0.02, 0.98, heading, va="top", fontsize=8)
    for name, (marker, color) in styles.items():
        axes[0].scatter([], [], color=color, marker=marker, label=name)
    axes[0].scatter([], [], color="#0072B2", marker="o", label="SheetShrink")
    figure.legend(*axes[0].get_legend_handles_labels(), frameon=False, loc="upper left",
                  bbox_to_anchor=(0.10, 0.93), ncol=3, fontsize=7)
    figure.text(0.02, 0.025, "Absolute times show workbook medians and quartiles. Ratio dots represent paired workbooks.\n"
                "Black diamonds/lines: median/IQR.\n"
                "One primary comparator attempt. SheetShrink: Lenovo 9 repetitions; HP 7–11. Visual validation separate.", fontsize=6.8)
    if outcomes:
        figure.text(0.02, 0.13,
                    "Status: plotted / other eligible / failed / payload excluded / incomplete receipt / partial.\n"
                    "Blue / hatched blue / crosshatched orange / dotted yellow / hatched gray / crosshatched purple.",
                    fontsize=6.8)
    figure.subplots_adjust(left=0.25 if multiple_hosts else 0.23, right=0.985, top=0.80, bottom=0.29 if outcomes else 0.20, wspace=0.40)
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


def draw_single_column(rows, output, a4_landscape=False, two_column=False):
    """Draw matched times and ratios with compact tool labels and shared corpus headings."""
    short_names = {"NXPowerLite": "NXPL", "XLStylesTool": "XLST",
                   "Check Performance": "CP", "Inquire": "INQ"}
    tools = ["NXPowerLite", "XLStylesTool", "Check Performance", "Inquire"]
    styles = dict(zip(tools, zip(["s", "^", "v", "D"],
                                ["#D55E00", "#009E73", "#E69F00", "#CC79A7"])))
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 7,
                         "pdf.fonttype": 42, "svg.fonttype": "none"})
    page_size = (3.5, 4.35)
    if two_column:
        page_size = (7.16, 3.6)
    if a4_landscape:
        page_size = (297 / 25.4, 210 / 25.4)
    label_size = 10 if a4_landscape else 6
    heading_size = 12 if a4_landscape else 7
    figure, axes = plt.subplots(1, 2, figsize=page_size)
    positions, labels = [], []
    for corpus_index, (corpus, title) in enumerate(
            (("destatis", "Destatis"), ("enron", "Enron"), ("datagov", "Data.gov"))):
        heading = corpus_index * 5
        axes[0].text(-0.24 if a4_landscape or two_column else -0.48, heading, title, transform=axes[0].get_yaxis_transform(),
                     fontweight="bold", fontsize=heading_size, va="center", clip_on=False)
        for tool_index, tool in enumerate(tools):
            selected = [r for r in rows if r["corpus"] == corpus and r["comparator"] == tool]
            if not selected:
                continue
            position = heading + tool_index + 1
            positions.append(position)
            labels.append(f"{short_names[tool]} ({len(selected)})")
            marker, color = styles[tool]
            times = np.array([r["comparator_seconds"] for r in selected])
            baseline = np.array([r["sheetshrink_seconds"] for r in selected])
            for values, offset, symbol, ink in ((times, -0.13, marker, color),
                                                 (baseline, 0.13, "o", "#0072B2")):
                q1, median, q3 = np.percentile(values, [25, 50, 75])
                axes[0].plot([q1, q3], [position + offset] * 2, color=ink, linewidth=1)
                axes[0].scatter(median, position + offset, s=15, marker=symbol, color=ink)
            ratios = times / baseline
            jitter = np.linspace(-0.2, 0.2, len(selected))
            np.random.default_rng(20261110).shuffle(jitter)
            axes[1].scatter(ratios, position + jitter, s=9, marker=marker,
                            facecolors=color, edgecolors=color,
                            linewidths=0.5, alpha=0.55, zorder=2)
            q1, median, q3 = np.percentile(ratios, [25, 50, 75])
            axes[1].plot([q1, q3], [position] * 2, color="black", linewidth=1.5)
            axes[1].scatter(median, position, s=12, marker="D", color="black")
    for axis in axes:
        axis.set_xscale("log")
        axis.set_ylim(15, -0.6)
        axis.grid(axis="x", which="major", color="#dddddd", linewidth=0.5)
        axis.set_axisbelow(True)
        axis.tick_params(axis="both", labelsize=label_size)
        axis.tick_params(axis="y", length=0)
        for side in ("left", "right", "top"):
            axis.spines[side].set_visible(False)
    axes[0].set_yticks(positions, labels)
    axes[1].set_yticks(positions, [""] * len(positions))
    axes[0].set_xlabel("Workflow time (s)", fontsize=label_size + 0.5)
    axes[1].set_xlabel("Tool / SheetShrink", fontsize=label_size + 0.5)
    axes[1].axvline(1, color="#999999", linestyle="--", linewidth=0.7)
    axes[0].scatter([], [], marker="o", color="#0072B2", s=15, label="SheetShrink")
    figure.legend(*axes[0].get_legend_handles_labels(), loc="upper right",
                  bbox_to_anchor=(1, 1), ncol=2, frameon=False, fontsize=label_size)
    # Reserve only the compact labels; corpus headings share their left edge.
    figure.subplots_adjust(left=0.12 if a4_landscape or two_column else 0.23,
                           right=0.99, top=0.91, bottom=0.14, wspace=0.28)
    if a4_landscape:
        figure.text(0.12, 0.025,
                    "NXPL: NXPowerLite   XLST: XLStylesTool   CP: Check Performance   INQ: Inquire",
                    fontsize=9)
    output.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output)
    figure.savefig(output.with_suffix(".png"), dpi=250)
    plt.close(figure)


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
