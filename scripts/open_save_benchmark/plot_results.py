"""Plot completed round trips and explicitly marked bounded failures."""
import argparse
import json
from pathlib import Path
import statistics

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", type=Path)
    args = parser.parse_args()
    records = [json.loads(line) for line in (args.root / "results.jsonl").read_text().splitlines()]
    records = [row for row in records if row["case_id"] != "probe"]
    cases = list(dict.fromkeys(row["case_id"] for row in records))
    libraries = list(dict.fromkeys(row["library"] for row in records))
    tools = ["Original", "SheetShrink", "NXPowerLite", "XLStylesTool", "CheckPerformance", "Inquire"]
    colors = ["#526477", "#00856a", "#4477aa", "#cc8833", "#aa4499", "#66aaaa"]
    figure, axes = plt.subplots(len(cases), len(libraries), figsize=(3.1 * len(libraries), 2.65 * len(cases)), squeeze=False)
    summary = []
    for row_index, case in enumerate(cases):
        for column, library in enumerate(libraries):
            axis = axes[row_index][column]
            for position, (tool, color) in enumerate(zip(tools, colors)):
                trials = [r for r in records if (r["case_id"], r["library"], r["tool"]) == (case, library, tool)]
                if not trials:
                    continue
                successes = [r for r in trials if r["status"] == "completed"]
                item = dict(case=case, library=library, tool=tool, trials=len(trials), completed=len(successes), statuses=[r["status"] for r in trials])
                if successes:
                    opened = statistics.median(r["worker_state"]["open_seconds"] for r in successes)
                    saved = statistics.median(r["worker_state"]["save_seconds"] for r in successes)
                    cleanup_times = [r["cleanup_seconds"] for r in successes if r.get("cleanup_seconds") is not None]
                    cleanup = statistics.median(cleanup_times) if cleanup_times else None
                    # Unknown cleanup cost is omitted and explicitly marked, never imputed.
                    left = cleanup if cleanup is not None else 0
                    if left:
                        axis.barh(position, left, color=color, alpha=.28)
                    axis.barh(position, opened, left=left, color=color)
                    axis.barh(position, saved, left=left + opened, color=color, hatch="////", edgecolor="white", linewidth=.3)
                    item.update(open_median_seconds=opened, save_median_seconds=saved, cleanup_seconds=cleanup,
                                total_seconds=None if cleanup is None else cleanup + opened + saved)
                    direct = [r["pipeline"]["pipeline_seconds"] for r in successes if r.get("pipeline", {}).get("pipeline_seconds") is not None]
                    if tool == "Original":
                        direct = [r["wall_seconds"] for r in successes]
                    if direct:
                        total = statistics.median(direct)
                        item["total_seconds"] = total
                        item["direct_end_to_end_median_seconds"] = total
                        overhead = max(0, total - left - opened - saved)
                        axis.barh(position, overhead, left=left + opened + saved,
                                  color="#bbbbbb", hatch="..", edgecolor="white", linewidth=.3)
                    repeats = [r["repeat_sheetshrink"] for r in successes if r.get("repeat_sheetshrink", {}).get("status") == "completed"]
                    if repeats:
                        item["repeat_sheetshrink_median_seconds"] = statistics.median(r["seconds"] for r in repeats)
                        item["repeat_sheetshrink_byte_identical"] = all(r.get("byte_identical", False) for r in repeats)
                    if cleanup is None:
                        axis.annotate("*", (opened + saved, position), xytext=(3, -3), textcoords="offset points", fontsize=9)
                failures = [r for r in trials if r["status"] != "completed"]
                if failures:
                    elapsed = statistics.median(r["wall_seconds"] for r in failures)
                    axis.scatter([elapsed], [position], marker="x", color="#b32424", s=25, zorder=4)
                    item["abort_median_seconds"] = elapsed
                item["peak_commit_mib"] = max(r["resources"]["peak_job_committed_bytes"] for r in trials) / 2**20
                summary.append(item)
            axis.set_yticks(range(len(tools)), tools if column == 0 else [])
            axis.invert_yaxis()
            axis.set_xscale("symlog", linthresh=1)
            panel = [item for item in summary if item["case"] == case and item["library"] == library]
            observed = [max(item.get("total_seconds") or (item.get("open_median_seconds", 0) + item.get("save_median_seconds", 0)), item.get("abort_median_seconds", 0)) for item in panel]
            axis.set_xlim(0, max(observed, default=1) * 1.2)
            axis.set_xlabel("Seconds (linear below 1, log above)", fontsize=8)
            axis.set_title(f"{case} / {library}", fontsize=10)
            axis.grid(axis="x", alpha=.18)
            axis.set_axisbelow(True)
            axis.spines[["top", "right"]].set_visible(False)
    figure.legend(handles=[Patch(facecolor="#526477", alpha=.28, label="Cleanup (when measured)"),
                           Patch(facecolor="#526477", label="Open"), Patch(facecolor="#526477", hatch="////", label="Save"),
                           Patch(facecolor="#bbbbbb", hatch="..", label="Startup / shutdown overhead")],
                  loc="upper center", ncol=4, frameon=False)
    limits = sorted({row["limit_bytes"] // 2**20 for row in records})
    figure.text(.5, .008, f"Red ×: bounded failure, not completion time. * Cleanup time unavailable. Job-commit cap: {limits} MiB; Timeouts (s): {sorted({row.get('timeout_seconds', 120) for row in records})}.", ha="center", fontsize=9)
    figure.tight_layout(rect=(0, .035, 1, .955))
    for extension in ("pdf", "png"):
        figure.savefig(args.root / f"open-save-comparison.{extension}", dpi=160)
    (args.root / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    memory_figure, memory_axes = plt.subplots(len(cases), len(libraries), figsize=(3.1 * len(libraries), 2.65 * len(cases)), squeeze=False)
    for row_index, case in enumerate(cases):
        for column, library in enumerate(libraries):
            axis = memory_axes[row_index][column]
            for position, (tool, color) in enumerate(zip(tools, colors)):
                matches = [item for item in summary if (item["case"], item["library"], item["tool"]) == (case, library, tool)]
                if matches:
                    item = matches[0]
                    axis.barh(position, item["peak_commit_mib"], color=color, alpha=.8)
                    if item["completed"] < item["trials"]:
                        axis.scatter([item["peak_commit_mib"]], [position], marker="x", color="#b32424", zorder=3)
            for limit in limits:
                axis.axvline(limit, color="#b32424", linestyle="--", linewidth=.8)
            axis.set_yticks(range(len(tools)), tools if column == 0 else [])
            axis.invert_yaxis()
            axis.set_title(f"{case} / {library}", fontsize=10)
            axis.set_xlabel("Peak job committed memory (MiB)", fontsize=8)
            axis.spines[["top", "right"]].set_visible(False)
    memory_figure.text(.5, .01, "Maximum observed across repetitions. Red × indicates at least one bounded failure. Dashed line marks the configured job limit.", ha="center", fontsize=9)
    memory_figure.tight_layout(rect=(0, .04, 1, 1))
    for extension in ("pdf", "png"):
        memory_figure.savefig(args.root / f"open-save-memory.{extension}", dpi=160)


if __name__ == "__main__":
    main()
