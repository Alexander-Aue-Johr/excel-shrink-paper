"""Render corpus worksheet distributions at the manuscript's printed size."""
import argparse
import json
from pathlib import Path
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "analyze_destatis_xlsx_files"))
import generate_distribution_of_size_reduction_chart_rust as chart


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("transfer_root", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--preaggregated", action="store_true")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.size": 8, "axes.titlesize": 9, "legend.fontsize": 7})
    figure, axes = plt.subplots(2, 3, figsize=(7.1, 4.0), sharex=True, sharey="row")
    metadata = {}
    for column, (corpus, label) in enumerate((("destatis", "Destatis"), ("enron", "Enron"), ("datagov", "Data.gov"))):
        if args.preaggregated:
            frame = chart.pd.read_csv(args.transfer_root / f"{corpus}-worksheet-ratios.csv.gz")
        else:
            data = chart.read_rust_analysis(args.transfer_root / f"{corpus}-v2/analysis.csv.gz")
            data = data[data["sheet_file"].str.startswith("xl/worksheets/")]
            frame = chart.build_distribution_data(data)
        ratio = (frame["original_size"]-frame["cleaned_size"]) / frame["original_size"].where(frame["original_size"] > 0, 1)
        positions = np.linspace(0, 100, len(frame))
        axes[0, column].fill_between(positions, np.sort(ratio.to_numpy())[::-1]*100, color="#4477aa", alpha=.75)
        axes[0, column].set_title(f"{label}\n{len(frame):,} worksheets")
        for panel in chart.CAUSE_PANELS:
            axes[1, column].plot(positions, np.sort(frame[panel.title].to_numpy())[::-1]*100,
                                 label=panel.title, color=panel.color, linewidth=1)
        for axis in axes[:, column]:
            axis.set_xlim(0, 100)
            axis.set_xticks([0, 50, 100])
            axis.axhline(0, color="#444444", linewidth=.4)
            axis.grid(axis="y", alpha=.2)
            axis.spines[["top", "right"]].set_visible(False)
        axes[1, column].set_ylim(-2, 102)
        axes[1, column].set_xlabel("Worksheet percentile")
        metadata[corpus] = {"worksheets": len(frame), "workbooks": frame["input_file"].nunique()}
        frame.to_csv(args.output / f"{corpus}-worksheet-ratios.csv.gz", index=False, compression={"method":"gzip", "mtime":0})
    axes[0, 0].set_ylabel("XML reduction (%)")
    axes[1, 0].set_ylabel("Reduction by cause (%)")
    handles, labels = axes[1, 0].get_legend_handles_labels()
    figure.legend(handles, labels, loc="lower center", ncol=3, frameon=False, bbox_to_anchor=(.5, -.01))
    figure.tight_layout(rect=(0, .18, 1, 1), h_pad=.8)
    for extension in ("pdf", "png"):
        figure.savefig(args.output / f"corpus-worksheet-distributions.{extension}", dpi=200)
    (args.output / "coverage.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
