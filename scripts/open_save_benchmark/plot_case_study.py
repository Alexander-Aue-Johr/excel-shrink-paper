"""Render compact case-study figures and native evidence from frozen measurements."""
import gzip
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import NullFormatter


ROOT = Path(__file__).resolve().parents[2]
PAPER = ROOT / "papers/sheetshrink"
EVIDENCE = PAPER / "evidence/open-save-case-study"
TOOLS = ["SheetShrink", "NXPowerLite", "XLStylesTool", "CheckPerformance", "Inquire"]
NAMES = dict(zip(TOOLS, ["SheetShrink", "NXPowerLite", "XLStylesTool", "Check Performance", "Inquire"]))
NAMES["Original"] = "Original"
COLORS = dict(zip(TOOLS + ["Original"], ["#0072B2", "#D55E00", "#009E73", "#E69F00", "#CC79A7", "#777777"]))
MARKERS = dict(zip(TOOLS + ["Original"], ["D", "s", "^", "v", "P", "o"]))


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def save_figure(figure, name):
    figure.savefig(PAPER / "figures" / (name + ".pdf"))
    plt.close(figure)


def draw_payload(native, cases):
    figure, axes = plt.subplots(1, 3, figsize=(7.15, 2.25), sharex=True, sharey=True)
    for axis, case_id, title in zip(axes, ["wanderungen", "qt-rr", "enron"],
                                    ["Wanderungen", "Data.gov QT/RR", "Enron"]):
        original = next(c for c in cases if c["case_id"] == case_id)["variants"]["Original"]
        for position, tool in enumerate(TOOLS):
            row = next(r for r in native if r["case_id"] == case_id and r["tool"] == tool)
            remaining = 100 * row["uncompressed_bytes"] / original["uncompressed_bytes"]
            axis.plot([remaining, 100], [position] * 2, color="#d7dce0", linewidth=1.2, zorder=1)
            axis.scatter(remaining, position, color=COLORS[tool], marker=MARKERS[tool], s=22, zorder=2)
            label = f"{remaining:.2f}".rstrip("0").rstrip(".") + "%"
            axis.annotate(label, (remaining, position), xytext=(0, 5),
                          textcoords="offset points", ha="center", fontsize=6.5)
        axis.axvline(100, color="#777777", linestyle="--", linewidth=0.6)
        axis.set_xscale("log")
        axis.set_xlim(0.5, 155)
        axis.set_ylim(4.5, -0.8)
        axis.set_title(title, fontsize=8)
        axis.set_xticks([1, 10, 100], ["1%", "10%", "100%"])
        axis.xaxis.set_minor_formatter(NullFormatter())
        axis.set_yticks(range(len(TOOLS)), [NAMES[t] for t in TOOLS])
        axis.tick_params(axis="y", length=0)
        axis.tick_params(axis="both", labelsize=7)
        axis.grid(axis="x", alpha=0.2)
        axis.set_axisbelow(True)
        for side in ["top", "right", "left"]:
            axis.spines[side].set_visible(False)
    figure.supxlabel("Remaining uncompressed payload (original = 100%)", fontsize=8)
    figure.subplots_adjust(left=0.18, right=0.98, top=0.87, bottom=0.22, wspace=0.25)
    save_figure(figure, "case-study-payload")


def read_library_resources(evidence_directory):
    path = PAPER / "evidence" / evidence_directory / "trials.jsonl.gz"
    with gzip.open(path, "rt", encoding="utf-8") as stream:
        trials = [json.loads(line) for line in stream]
    groups = sorted({(r["case_id"], r["library"], r["tool"]) for r in trials})
    summaries = []
    for case_id, library, tool in groups:
        if library not in {"openpyxl", "excel"}:
            continue
        selected = [r for r in trials if (r["case_id"], r["library"], r["tool"]) == (case_id, library, tool)]
        assert len(selected) == 3
        complete = [r for r in selected if r["status"] == "completed"]
        times = [r["worker_state"]["open_seconds"] + r["worker_state"]["save_seconds"] for r in complete]
        process_name = "python.exe" if library == "openpyxl" else "excel.exe"
        peaks = [max(p["working_set_bytes"] for p in r["resources"]["observed_process_peaks"].values()
                     if p["name"].lower() == process_name) / 2**20 for r in complete]
        summaries.append({"case": case_id, "library": library, "tool": tool,
                          "completed": len(complete), "attempts": 3,
                          "statuses": [r["status"] for r in selected],
                          "library_open_save_median_seconds": float(np.median(times)) if times else None,
                          "library_open_save_range_seconds": [min(times), max(times)] if times else None,
                          "sampled_worker_rss_max_mib": max(peaks) if peaks else None})
    return summaries


def draw_resources(resources):
    panels = [("wanderungen", "openpyxl", "Wanderungen · openpyxl"),
              ("qt-rr", "openpyxl", "Data.gov QT/RR · openpyxl"),
              ("enron", "openpyxl", "Enron · openpyxl"),
              ("wanderungen", "excel", "Wanderungen · Excel"),
              ("qt-rr", "excel", "Data.gov QT/RR · Excel")]
    figure, grid = plt.subplots(2, 3, figsize=(7.15, 3.9), sharex=True, sharey=True)
    axes = grid.flatten()
    for axis, (case_id, library, title) in zip(axes, panels):
        stopped = []
        for tool in ["Original"] + TOOLS:
            row = next(r for r in resources if r["case"] == case_id
                       and r["library"] == library and r["tool"] == tool)
            if row["completed"] == 0:
                assert set(row["statuses"]) == {"memory_limit"}
                stopped.append(NAMES[tool])
                continue
            assert row["completed"] == row["attempts"] == 3
            x = row["sampled_worker_rss_max_mib"] / 1024
            y = row["library_open_save_median_seconds"]
            assert 0.04 <= x <= 4.5 and 0.5 <= y <= 180
            axis.plot([x, x], row["library_open_save_range_seconds"], color=COLORS[tool], linewidth=0.8)
            axis.scatter(x, y, s=52 if tool == "Original" else 24,
                         facecolors="none" if tool == "Original" else COLORS[tool],
                         edgecolors=COLORS[tool], marker=MARKERS[tool],
                         linewidths=1, zorder=4 if tool == "Original" else 3)
        if stopped:
            axis.text(0.04, 0.96, "4 GiB commit stop\n" + ", ".join(stopped),
                      transform=axis.transAxes, va="top", fontsize=6, color="#994400")
        axis.set_title(title, fontsize=7.5)
        axis.set_xscale("log")
        axis.set_yscale("log")
        axis.set_xlim(0.04, 4.5)
        axis.set_ylim(0.5, 180)
        axis.set_xticks([0.05, 0.25, 1, 4], ["0.05", "0.25", "1", "4"])
        axis.set_yticks([1, 10, 100], ["1", "10", "100"])
        axis.xaxis.set_minor_formatter(NullFormatter())
        axis.yaxis.set_minor_formatter(NullFormatter())
        axis.tick_params(labelsize=6.5)
        axis.grid(which="major", alpha=0.2)
        axis.set_axisbelow(True)
        for side in ["top", "right"]:
            axis.spines[side].set_visible(False)
    # The upper-right panel has no lower neighbor with its own x-axis.
    axes[2].tick_params(labelbottom=True)
    axes[5].set_axis_off()
    handles = [plt.Line2D([], [], marker=MARKERS[t], linestyle="none",
                         markerfacecolor="none" if t == "Original" else COLORS[t],
                         markeredgecolor=COLORS[t], label=NAMES[t], markersize=5)
               for t in ["Original"] + TOOLS]
    axes[5].legend(handles=handles, loc="center left", frameon=False, fontsize=7)
    figure.supxlabel("Peak sampled library-process working set (GiB)", fontsize=8)
    figure.supylabel("Library opening + saving time (s)", fontsize=8)
    figure.subplots_adjust(left=0.09, right=0.99, bottom=0.13, top=0.93, hspace=0.38, wspace=0.25)
    save_figure(figure, "case-study-resources")


def write_native_table(native):
    lines = [r"\begin{table}[t]", r"\centering", r"\footnotesize",
             r"\setlength{\tabcolsep}{3pt}",
             r"\caption{Measured native pixel observations. Pairs gives measured/recorded regions; size changes counts unequal image dimensions.}",
             r"\label{tab:case-study-native}", r"\begin{tabular}{lrrr}", r"\hline",
             r"Case / tool & Pixels & Pairs & \shortstack{Size\\changes} \\", r"\hline"]
    for case_id, title in [("wanderungen", "Wanderungen"), ("enron", "Enron")]:
        lines.append(r"\multicolumn{4}{l}{\textbf{" + title + r"}} \\")
        for row in sorted(native, key=lambda r: TOOLS.index(r["tool"])):
            visual = row["visual"]
            if row["case_id"] != case_id or not visual.get("N"):
                continue
            name = NAMES[row["tool"]] + (r"\textsuperscript{*}" if case_id == "wanderungen" else "")
            lines.append(f"{name} & {visual['D']:,} & {visual['measured_regions']}/{visual['recorded_regions']} & "
                         + str(visual["different_dimensions_in_measured_regions"]) + r" \\")
        lines.append(r"\hline")
    lines += [r"\end{tabular}", r"\par\smallskip",
              r"\begin{minipage}{\columnwidth}\footnotesize \textsuperscript{*}Normalized baseline and separately regenerated SheetShrink control. "
              r"Counts use exact RGBA differences in matched-origin common rectangles. Overlapping captures can repeat pixels. "
              r"Different plans and incomplete scopes preclude cross-tool preservation ranking. Pixel denominators and identities remain in the artifact.\end{minipage}",
              r"\end{table}"]
    (EVIDENCE / "native-observations.tex").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8, "pdf.fonttype": 42})
    native = read_json(EVIDENCE / "matched-native-summary.json")
    cases = read_json(EVIDENCE / "sizes.json")
    resources = read_library_resources("bounded-open-save") + read_library_resources("enron-open-save-followup")
    (EVIDENCE / "plotted-library-resources.json").write_text(
        json.dumps(resources, indent=2) + "\n", encoding="utf-8")
    draw_payload(native, cases)
    draw_resources(resources)
    write_native_table(native)


if __name__ == "__main__":
    main()
