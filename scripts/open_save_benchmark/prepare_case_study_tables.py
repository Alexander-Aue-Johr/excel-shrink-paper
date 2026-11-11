"""Render manuscript tables from verified size inventories and retained trial receipts."""
import gzip
import json
import statistics
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PAPER = ROOT / "papers/sheetshrink"
EVIDENCE = PAPER / "evidence/open-save-case-study"
TOOLS = ["Original", "SheetShrink", "NXPowerLite", "XLStylesTool", "CheckPerformance", "Inquire"]
NAMES = {"Original": "Original", "SheetShrink": "Sheet Shrink", "NXPowerLite": "NXPowerLite", "XLStylesTool": "XLStylesTool", "CheckPerformance": "Check Performance", "Inquire": "Inquire"}
CASES = {"wanderungen": "Destatis Migration", "finance": "Destatis Municipal Finance", "qt-rr": "Data.gov QT/RR", "enron": "Enron b2d0dfef"}
LIBRARIES = {"openpyxl": "openpyxl", "pandas": "pandas to XLSX", "pandas-csv": "pandas to CSV", "excel": "Excel COM", "r-openxlsx": "R openxlsx", "r-readxl": "R readxl + writexl"}


def cell(value, state=None):
    if value is None:
        return r"\colorbox{casemissing}{\strut Pending}"
    if state:
        return rf"\colorbox{{{state}}}{{\strut {value}}}"
    return str(value)


def paired_bar(label, value, baseline, maximum, width=4.0):
    if value is None or baseline is None:
        return label
    scale = 1.1 / maximum
    original_length = scale
    output_length = scale * value / baseline
    color = "casebarblue" if value <= baseline else "casebarorange"
    return rf"\CasePairBars[{width}]{{{label}}}{{{original_length:.5f}}}{{{output_length:.5f}}}{{{color}}}"


def median(rows, getter):
    values = [value for row in rows if (value := getter(row)) is not None]
    return statistics.median(values) if values else None


def seconds(value):
    return cell(None) if value is None else f"{value:.2f}"


def sampled_rss(row, library_only=False):
    phases = row["resources"].get("sampled_phase_peaks", {})
    if library_only:
        phases = {name: value for name, value in phases.items() if name in {"importing", "opening", "saving", "completed"}}
    return max((phase["working_set_bytes"] for phase in phases.values()), default=0) / 2**20


def table_start(caption, columns, continued=False, label=None):
    lines = [r"\begin{table*}[p]", r"\centering", r"\footnotesize", r"\setlength{\tabcolsep}{2.5pt}", r"\renewcommand{\arraystretch}{1.08}"]
    if continued:
        lines.append(r"\addtocounter{table}{-1}")
    lines.append(r"\caption{" + caption + "}")
    if label:
        lines.append(r"\label{" + label + "}")
    lines.extend([r"\begin{tabular}{" + columns + "}", r"\hline"])
    return lines


def table_end(note):
    return [r"\hline", r"\end{tabular}", r"\par\smallskip", r"\begin{minipage}{0.98\textwidth}\footnotesize " + note + r"\end{minipage}", r"\end{table*}"]


def transformation_table(cases, visuals):
    lines = table_start("Opening/memory case-study inputs and native visual evidence. Sizes are decimal MB. Payload is the sum of all uncompressed ZIP-member sizes. Reduction columns compare each output with its original. Pending identifies unavailable evidence. Paired bars show original (gray) and output (blue/orange), with original = 100\\%.", "lrrrccc", label="tab:case-study-transformations")
    lines.append(r"Variant & ZIP (reduction) & Payload (reduction) & Pixels $D/N$ & Sheets & Geom. & State \\")
    for case in cases:
        original = case["variants"]["Original"]
        lines.extend([r"\hline", r"\multicolumn{7}{l}{\textbf{" + CASES[case["case_id"]] + r"}} \\"])
        zip_scale = max(v["zip_bytes"] / original["zip_bytes"] for v in case["variants"].values())
        payload_scale = max(v["uncompressed_bytes"] / original["uncompressed_bytes"] for v in case["variants"].values())
        for tool in TOOLS:
            size = case["variants"].get(tool)
            if tool == "SheetShrink":
                fresh = case.get("fresh_sheetshrink")
                size = {"zip_bytes": fresh["zip_bytes"][0], "uncompressed_bytes": None} if fresh else None
            visual = next((v for v in visuals if v["case_id"] == case["case_id"] and v["tool"] == tool), None)
            compressed = size["zip_bytes"] if size else None
            payload = size["uncompressed_bytes"] if size else None
            if tool == "Original":
                pixels, sheets, geometry, state = "Ref.", "Ref.", "Ref.", "Ref."
            elif visual and "different_pixel_observations" in visual:
                pixels = cell(f"{visual['different_pixel_observations']:,} / Pending", "caseadverse")
                sheets = f"{visual['changed_sheets']}/{visual['observed_sheets']}/{visual['total_sheets']}"
                geometry = str(visual["geometry_findings"])
                state = cell("L", "caseadverse")
            else:
                pixels, sheets, geometry = cell(None), cell(None), cell(None)
                state = cell("M" if visual and visual["status"] == "memory_limit" else "F" if visual and visual["status"] == "failed" else "Pending", "casemissing")
            compressed_label = cell(None) if compressed is None else f"{compressed / 1e6:.3f} ({100 * (1 - compressed / original['zip_bytes']):.2g}\\%)"
            payload_label = cell(None) if payload is None else f"{payload / 1e6:.2f} ({100 * (1 - payload / original['uncompressed_bytes']):.2g}\\%)"
            compressed_display = paired_bar(compressed_label, compressed, original["zip_bytes"], zip_scale)
            payload_display = paired_bar(payload_label, payload, original["uncompressed_bytes"], payload_scale)
            values = [NAMES[tool], compressed_display, payload_display, pixels, sheets, geometry, state]
            lines.append(" & ".join(values) + r" \\")
    note = r"Sheet Shrink outputs are freshly generated in each six-worker trial. Their compressed size is recorded consistently across repetitions, while their uncompressed payload and matched native visual evidence remain pending. NXPowerLite uses the cell-formatting-focused profile. XLStylesTool retains the tested settings. Ref. identifies the original baseline. L denotes completed targeted captures with coverage limitations, M a memory-limit stop, F a failed attempt and Pending unavailable evidence. Sheets gives changed/observed/total worksheets. Geometry counts flagged capture findings. Pixel counts aggregate retained diagnostic observations under policy v12 after paired native rebuild. The total comparable-pixel denominator N remains pending. Different capture scopes and unresolved attribution constrain comparisons. Gray bars denote the original. Blue denotes a value at or below the original, and orange a larger value. Orange cell backgrounds mark observed visual differences. Each pair has a common zero. Original = 100\%, with a shared linear scale within each workbook and size column. Absolute MB and reduction percentages accompany the bars."
    lines.extend(table_end(note))
    return "\n".join(lines) + "\n"


def resource_tables(cases, records):
    output = []
    summaries = []
    for index, case in enumerate(cases[:3]):
        caption = f"Opening/saving resources on the Ryzen 5 4500U laptop, 8 GiB installed RAM, six logical processors. {CASES[case['case_id']]}. Part {index + 1} of 3. Each row has three attempts under a 4 GiB job-commit cap and 900 s timeout."
        lines = table_start(caption, "lrrrrrrrc", continued=index > 0, label="tab:case-study-resources" if index == 0 else None)
        lines.append(r"Variant & Cleanup & Open & Save & End-to-end & RSS lib. & RSS all & Commit & Outcome \\")
        lines.append(r" & \multicolumn{4}{c}{seconds, median of completed trials} & \multicolumn{3}{c}{MiB, maximum across attempts} & \\")
        for library, title in LIBRARIES.items():
            group = [r for r in records if r["case_id"] == case["case_id"] and r["library"] == library]
            if not group:
                continue
            lines.extend([r"\hline", r"\multicolumn{9}{l}{\textit{" + title + r"}} \\"])
            original_complete = [r for r in group if r["tool"] == "Original" and r["status"] == "completed"]
            original_total = median(original_complete, lambda r: r["wall_seconds"])
            original_rss = max((sampled_rss(r, True) for r in original_complete), default=None) if len(original_complete) == 3 else None
            completed_totals = []
            completed_memory = []
            for candidate_tool in TOOLS:
                completed_rows = [r for r in group if r["tool"] == candidate_tool and r["status"] == "completed"]
                if len(completed_rows) == 3:
                    elapsed = median(completed_rows, lambda r: r.get("pipeline", {}).get("pipeline_seconds") if candidate_tool == "SheetShrink" else r["wall_seconds"])
                    completed_totals.append(elapsed)
                    completed_memory.append(max(sampled_rss(r, True) for r in completed_rows))
            time_scale = max([1] + [t / original_total for t in completed_totals]) if original_total else 1
            memory_scale = max([1] + [m / original_rss for m in completed_memory]) if original_rss else 1
            for tool in TOOLS:
                trials = [r for r in group if r["tool"] == tool]
                completed = [r for r in trials if r["status"] == "completed"]
                cleanup = median(completed, lambda r: r.get("cleanup_seconds")) if tool == "SheetShrink" else None
                opened = median(completed, lambda r: r["worker_state"].get("open_seconds"))
                saved = median(completed, lambda r: r["worker_state"].get("save_seconds"))
                total = median(completed, lambda r: r.get("pipeline", {}).get("pipeline_seconds") if tool == "SheetShrink" else r["wall_seconds"])
                rss_library = max(sampled_rss(r, True) for r in trials)
                rss_all = max(sampled_rss(r) for r in trials)
                commit = max(r["resources"]["peak_job_committed_bytes"] for r in trials) / 2**20
                cpu = median(trials, lambda r: r["resources"]["cpu_user_seconds"] + r["resources"]["cpu_kernel_seconds"])
                outcome = f"{len(completed)}/{len(trials)}"
                display_total = seconds(total)
                if not completed:
                    stopped = median(trials, lambda r: r["wall_seconds"])
                    display_total = cell(f"M@{stopped:.2f}", "caseadverse")
                    outcome = cell(outcome + " M", "caseadverse")
                elif tool == "SheetShrink" and original_total is not None:
                    display_total = cell(f"{total:.2f}", "casegood" if total < original_total else "caseadverse")
                if len(completed) == 3 and len(original_complete) == 3:
                    display_total = paired_bar(seconds(total), total, original_total, time_scale, width=2.2)
                    display_rss = paired_bar(f"{rss_library:.0f}", rss_library, original_rss, memory_scale, width=2.2)
                else:
                    display_rss = f"{rss_library:.0f}"
                cleanup_display = seconds(cleanup) if tool == "SheetShrink" else "--" if tool == "Original" else "prep."
                lines.append(" & ".join([NAMES[tool], cleanup_display, seconds(opened), seconds(saved), display_total, display_rss, f"{rss_all:.0f}", f"{commit:.0f}", outcome]) + r" \\")
                summaries.append(dict(case=case["case_id"], library=library, tool=tool, completed=len(completed), attempts=len(trials), cleanup_median_seconds=cleanup, open_median_seconds=opened, save_median_seconds=saved, end_to_end_median_seconds=total, sampled_library_rss_max_mib=rss_library, sampled_pipeline_rss_max_mib=rss_all, peak_job_commit_max_mib=commit, cpu_median_seconds=cpu, wall_time_range_seconds=[min(r["wall_seconds"] for r in trials), max(r["wall_seconds"] for r in trials)], statuses=[r["status"] for r in trials]))
        if index == 2:
            lines.extend([r"\hline", r"\multicolumn{9}{l}{Enron fourth-case measurements use a separate host lane in Section~\ref{sec:enron-open-save-followup}.} \\"])
        note = r"Sheet Shrink runs fresh cleanup immediately before the library. Its end-to-end clock spans cleanup launch through library exit. Original and prepared NXPowerLite, XLStylesTool, Check Performance and Inquire rows report library process lifetime. prep. denotes preparation outside this trial. Open and save measure the retained operation endpoints. End-to-end includes startup/shutdown and is measured directly. RSS lib. is the maximum sampled process-tree working set in importing, opening, saving and completed phases. RSS all includes cleanup and other recorded phases. Sampling is nominally 20 ms. Commit is the Windows job peak. The columns have different counter semantics. Outcome gives completed/attempted. M@t is elapsed time until the memory-limit stop, with missing completion phases marked Pending. Blue/orange bars indicate values at or below/above the original on the same host and endpoint. Paired bars in end-to-end and library RSS cells use original = 100\%, a common zero and a shared linear scale per library/workbook/host. Gray denotes the original and blue/orange the compared value. Bars require three completed attempts in both conditions. Other hosts, dispersion and CPU fields remain explicit in the machine-readable artifact."
        lines.extend(table_end(note))
        output.append("\n".join(lines))
    (EVIDENCE / "resource-summary.json").write_text(json.dumps(summaries, indent=2) + "\n")
    return "\n\n".join(output) + "\n"


def main():
    cases = json.loads((EVIDENCE / "sizes.json").read_text())
    visuals = json.loads((EVIDENCE / "visual-summary.json").read_text())
    with gzip.open(PAPER / "evidence/bounded-open-save/trials.jsonl.gz", "rt") as stream:
        records = [json.loads(line) for line in stream]
    assert len(records) == 288
    (EVIDENCE / "transformations.tex").write_text(transformation_table(cases, visuals))
    (EVIDENCE / "resources.tex").write_text(resource_tables(cases, records))


if __name__ == "__main__":
    main()
