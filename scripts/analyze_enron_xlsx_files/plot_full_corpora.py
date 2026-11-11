"""Verify transferred full-corpus tables and reproduce reduction distributions."""
import argparse
import csv
import hashlib
import json
from pathlib import Path
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "analyze_destatis_xlsx_files"))
import generate_distribution_of_size_reduction_chart_rust as chart


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("transfer_root", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--workbooks-only", action="store_true")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.size": 8, "axes.titlesize": 9})
    figure, axes = plt.subplots(1, 3, figsize=(7.1, 2.25), sharey=True)
    summaries = {}
    for axis, corpus in zip(axes, ("destatis", "enron", "datagov")):
        directory = args.transfer_root / f"{corpus}-v2"
        for entry in json.loads((directory / "inventory.json").read_text(encoding="utf-8-sig")):
            payload = (directory / entry["file"]).read_bytes()
            if len(payload) != entry["bytes"] or hashlib.sha256(payload).hexdigest() != entry["sha256"]:
                raise ValueError(f"Transfer verification failed for {corpus}/{entry['file']}")
        with (directory / "workbooks.csv").open(encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.DictReader(stream))
        processed = [row for row in rows if row["status"] == "processed"]
        ratios = sorted((100 * (1 - int(row["output_bytes"]) / int(row["input_bytes"])) for row in processed), reverse=True)
        percentiles = [100 * i / max(1, len(ratios) - 1) for i in range(len(ratios))]
        axis.fill_between(percentiles, ratios, color="#4477aa", alpha=.7)
        axis.axhline(0, color="black", linewidth=.5)
        label = "Data.gov" if corpus == "datagov" else corpus.capitalize()
        axis.set_title(f"{label} (n={len(processed):,})")
        axis.set_xlabel("Workbook percentile")
        axis.spines[["top", "right"]].set_visible(False)
        summaries[corpus] = dict(processed=len(processed), other_statuses=len(rows)-len(processed),
            unchanged=sum(int(row["output_bytes"]) == int(row["input_bytes"]) for row in processed),
            grew=sum(int(row["output_bytes"]) > int(row["input_bytes"]) for row in processed),
            aggregate_compressed_reduction_percent=100*(1-sum(int(row["output_bytes"]) for row in processed)/sum(int(row["input_bytes"]) for row in processed)))
        if args.workbooks_only:
            continue
        data = chart.read_rust_analysis(directory / "analysis.csv.gz")
        for unit, frame in (("sheet", chart.build_distribution_data(data)),
                            ("workbook", chart.build_file_distribution_data(data))):
            frame["Reduction Ratio"] = (frame["original_size"] - frame["cleaned_size"]) / frame["original_size"].where(frame["original_size"] > 0, 1)
            target = args.output / f"{corpus}-xml-by-{unit}.pdf"
            chart.plot_distribution(frame, target, target.with_suffix(".png"),
                title=f"{corpus.capitalize()} structural XML reduction by {unit}",
                x_axis_label="Rank, sorted independently within each panel")
        print(corpus, summaries[corpus], flush=True)
    axes[0].set_ylabel("Compressed XLSX reduction (%)")
    figure.tight_layout()
    for extension in ("pdf", "png"):
        figure.savefig(args.output / f"compressed-workbook-distributions.{extension}", dpi=180)
    (args.output / "summary.json").write_text(json.dumps(summaries, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
