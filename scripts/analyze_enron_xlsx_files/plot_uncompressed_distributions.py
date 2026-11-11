"""Plot ZIP-directory workbook sizes and worksheet XML with common display limits."""
import argparse
from collections import Counter
import json
import math
from pathlib import Path
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import matplotlib.patheffects as effects
import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "analyze_destatis_xlsx_files"))
import generate_distribution_of_size_reduction_chart_rust as chart

CORPORA = (("destatis", "Destatis"), ("enron", "Enron"), ("datagov", "Data.gov"))
TIERS = ((0, 50, "#A7C6ED", ""), (50, 80, "#F0E442", "/"),
         (80, 90, "#E69F00", "//"), (90, 95, "#D55E00", "xx"),
         (95, 100.01, "#7F1D1D", "xxxx"))
COLORS = ("#E69F00", "#56B4E9", "#F0E442", "#0072B2", "#D55E00", "#009E73", "#CC79A7", "#777777")
STYLES = ("-", "--", "-.", ":", "-", "--", "-.", ":")
MARKERS = ("o", "s", "^", "v", "D", "P", "X", "*")


def reduction(original, cleaned):
    return 100 * (original - cleaned) / original.where(original > 0, 1)


def load_frames(args, corpus):
    frame = pd.read_csv(args.worksheets / f"{corpus}-worksheet-ratios.csv.gz")
    series = {"Worksheet XML": reduction(frame.original_size, frame.cleaned_size).to_numpy()}
    if not args.worksheets_only:
        path = args.packages / f"{corpus}-v2/uncompressed-workbooks.csv"
        books = pd.read_csv(path).fillna("")
        books = books[books.status == "processed"]
        expected = pd.read_csv(args.packages / f"{corpus}-v2/workbooks.csv").fillna("")
        expected = expected[expected.status == "processed"]
        identity_columns = ["source_xlsx_sha256", "output_sha256"]
        if Counter(map(tuple, books[identity_columns].to_numpy())) != Counter(map(tuple, expected[identity_columns].to_numpy())):
            raise ValueError(f"ZIP-directory observations differ from the processed cohort for {corpus}")
        if "error" in books and books.error.astype(str).str.len().gt(0).any():
            raise ValueError(f"Incomplete ZIP-directory observations for {corpus}")
        series = {"Workbook payload": reduction(pd.to_numeric(books.original_uncompressed_bytes),
                    pd.to_numeric(books.output_uncompressed_bytes)).to_numpy(), **series}
        books.to_csv(args.output / f"{corpus}-uncompressed-workbooks.csv.gz", index=False,
                     compression={"method": "gzip", "mtime": 0})
    return frame, series


def tier_fill(axis, values):
    ordered = np.sort(values)[::-1]
    x = np.linspace(0, 100, len(ordered))
    for low, high, color, hatch in TIERS:
        axis.fill_between(x, 0, ordered, where=(ordered >= low) & (ordered < high),
                          facecolor=color, edgecolor="#444444", linewidth=0, hatch=hatch)
    axis.plot(x, ordered, color="#333333", linewidth=.65)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("worksheets", type=Path)
    parser.add_argument("packages", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--worksheets-only", action="store_true")
    parser.add_argument("--visibility-threshold", type=float, default=.5)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    frames = {corpus: load_frames(args, corpus) for corpus, _ in CORPORA}
    distributions = [v for _, series in frames.values() for v in series.values()]
    causes = [100 * frame[p.title].to_numpy() for frame, _ in frames.values() for p in chart.CAUSE_PANELS]
    extent = max(100 * np.count_nonzero(v >= args.visibility_threshold) / len(v) for v in distributions + causes)
    limit = min(100, max(5, 5 * math.ceil(extent / 5)))
    rows = len(next(iter(frames.values()))[1]) + 1
    plt.rcParams.update({"font.size": 8, "axes.titlesize": 9, "legend.fontsize": 6.7, "hatch.linewidth": .3})
    figure, axes = plt.subplots(rows, 3, figsize=(7.1, 1.35 * rows + 1.1), sharex=True, sharey="row")
    metadata = {"common_percentile_limit": limit, "visibility_threshold_percentage_points": args.visibility_threshold, "corpora": {}}
    ticks = sorted(set([0, limit] + [tick for tick in range(20, limit, 20) if limit-tick >= 10]))
    for column, (corpus, label) in enumerate(CORPORA):
        frame, series = frames[corpus]
        observations = {}
        for row, (name, values) in enumerate(series.items()):
            tier_fill(axes[row, column], values)
            observations[name] = {"count": len(values), "grew": int((values < 0).sum()),
                "minimum_reduction_percent": float(values.min()),
                "tier_counts": {str(low): int(((values >= low) & (values < high)).sum()) for low, high, _, _ in TIERS}}
            axes[row, column].set_title(f"{label} · {len(values):,}" if row == 0 else f"{len(values):,} worksheets")
            axes[row, column].set_ylim(0, 102)
            if column == 0:
                axes[row, column].set_ylabel(f"{name}\nreduction (%)")
        for index, panel in enumerate(chart.CAUSE_PANELS):
            values = np.sort(100 * frame[panel.title].to_numpy())[::-1]
            x = np.linspace(0, 100, len(values))
            visible = x <= limit
            line, = axes[-1, column].plot(x[visible], values[visible], color=COLORS[index],
                linestyle=STYLES[index], marker=MARKERS[index], markersize=2,
                markeredgecolor="#333333", markeredgewidth=.3,
                markevery=max(1, np.count_nonzero(visible)//7), linewidth=1, label=panel.title)
            line.set_path_effects([effects.Stroke(linewidth=1.4, foreground="#444444"), effects.Normal()])
        for axis in axes[:, column]:
            axis.set_xlim(0, limit)
            axis.set_xticks(ticks)
            axis.set_yticks([0, 50, 100])
            axis.grid(axis="y", alpha=.18)
            axis.spines[["top", "right"]].set_visible(False)
        axes[-1, column].set_ylim(0, 102)
        axes[-1, column].set_xlabel("Independently ranked percentile")
        metadata["corpora"][corpus] = observations
    axes[-1, 0].set_ylabel("Reduction by cause (%)")
    tier_handles = [Patch(facecolor=color, edgecolor="#444444", hatch=hatch,
                         label=f"{low}–{high:g}%" if high <= 100 else f"≥{low}%") for low, high, color, hatch in TIERS]
    figure.legend(handles=tier_handles, loc="lower center", ncol=5, bbox_to_anchor=(.5, .115), frameon=False)
    handles, labels = axes[-1, 0].get_legend_handles_labels()
    figure.legend(handles, labels, loc="lower center", ncol=3, bbox_to_anchor=(.5, -.002), frameon=False)
    figure.tight_layout(rect=(0, .21, 1, 1), h_pad=.7)
    for extension in ("pdf", "png"):
        figure.savefig(args.output / f"corpus-uncompressed-distributions.{extension}", dpi=220)
    (args.output / "coverage.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    print(json.dumps(metadata), flush=True)


if __name__ == "__main__":
    main()
