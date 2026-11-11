"""Render manuscript tables from verified size inventories and retained trial receipts."""
import gzip
import json
import statistics
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PAPER = ROOT / "papers/sheetshrink"
EVIDENCE = PAPER / "evidence/open-save-case-study"
TOOLS = ["Original", "SheetShrink", "NXPowerLite", "XLStylesTool", "CheckPerformance", "Inquire"]
NAMES = {"Original": "Original", "SheetShrink": "SheetShrink", "NXPowerLite": "NXPowerLite", "XLStylesTool": "XLStylesTool", "CheckPerformance": "Check Performance", "Inquire": "Inquire"}
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


def native_evidence_cell(visual):
    status = visual["status"]
    if "N" in visual and visual["N"] > 0:
        lines = [f"{visual['D']:,} / {visual['N']:,}",
                 f"{visual['measured_regions']}/{visual['recorded_regions']} image pairs"]
        if visual["different_dimensions_in_measured_regions"]:
            lines.append(f"{visual['different_dimensions_in_measured_regions']} size differences")
        if visual["baseline_scope"] == "documented_normalized_baseline":
            lines.append("Normalized baseline")
        return r"\begin{tabular}[c]{@{}l@{}}" + r"\\".join(lines) + r"\end{tabular}"
    if "recorded_regions" in visual:
        if visual["recorded_regions"] == 0:
            return "No capture regions"
        return f"No retained pairs (0/{visual['recorded_regions']})"
    return {"memory_limit": "Memory limit", "failed": "Failed",
            "running": "No completed report", "no_matched_report": "No matched report"}.get(
                status, "No matched report")


def transformation_table(cases, native_rows):
    lines = table_start(
        "Selected-case uncompressed payload reduction and retained native pixel evidence. "
        "Payload is the sum of all ZIP-member sizes in decimal MB. "
        "Image-pair counts state measured/recorded coverage.",
        "lrl", label="tab:case-study-transformations")
    lines[0] = r"\begin{table*}[t]"
    lines.append(r"Variant & Payload MB (reduction) & Changed / compared pixels; coverage \\")
    for case in cases:
        if case["case_id"] == "finance":
            continue
        original = case["variants"]["Original"]
        selected = [r for r in native_rows if r["case_id"] == case["case_id"]]
        scale = max([1] + [r["uncompressed_bytes"] / original["uncompressed_bytes"] for r in selected])
        lines.extend([r"\hline", r"\multicolumn{3}{l}{\textbf{" + CASES[case["case_id"]] + r"}} \\"])
        for tool in TOOLS:
            if tool == "Original":
                payload = original["uncompressed_bytes"]
                visual = "Reference"
                name = NAMES[tool]
            else:
                row = next(r for r in selected if r["tool"] == tool)
                payload = row["uncompressed_bytes"]
                visual = native_evidence_cell(row["visual"])
                name = NAMES[tool] + (r"\textsuperscript{*}" if row["representative_control"] else "")
            reduction = 100 * (1 - payload / original["uncompressed_bytes"])
            label = f"{payload / 1e6:.2f} ({reduction:.2g}\\%)"
            display = paired_bar(label, payload, original["uncompressed_bytes"], scale, width=4.4)
            lines.append(" & ".join([name, display, visual]) + r" \\")
    note = (
        r"\textsuperscript{*}SheetShrink uses separately regenerated controls whose uncompressed members match the native candidates; "
        "full per-trial member inventories of the timed pipelines were not retained. "
        "Wanderungen uses a normalized baseline. Counts compare exact RGBA values in common rectangles with matched origins; "
        "overlapping captures can count pixels repeatedly. Image-size differences remain separate from common-area pixel counts. "
        "Different plans and incomplete scopes preclude a cross-tool preservation ranking. "
        "No capture regions or retained pairs provide no pixel-equality observation. "
        "Gray bars show the original; blue/orange show payload at or below/above the original. "
        "Input/output hashes, scope details and failures remain in the artifact.")
    lines.extend(table_end(note))
    return "\n".join(lines) + "\n"


def resource_summaries(cases, records):
    summaries = []
    for case in cases[:3]:
        for library in LIBRARIES:
            group = [r for r in records if r["case_id"] == case["case_id"] and r["library"] == library]
            if not group:
                continue
            for tool in TOOLS:
                trials = [r for r in group if r["tool"] == tool]
                completed = [r for r in trials if r["status"] == "completed"]
                summaries.append(dict(
                    case=case["case_id"],
                    library=library,
                    tool=tool,
                    completed=len(completed),
                    attempts=len(trials),
                    cleanup_median_seconds=median(completed, lambda r: r.get("cleanup_seconds")) if tool == "SheetShrink" else None,
                    open_median_seconds=median(completed, lambda r: r["worker_state"].get("open_seconds")),
                    save_median_seconds=median(completed, lambda r: r["worker_state"].get("save_seconds")),
                    end_to_end_median_seconds=median(completed, lambda r: r.get("pipeline", {}).get("pipeline_seconds") if tool == "SheetShrink" else r["wall_seconds"]),
                    sampled_library_rss_max_mib=max(sampled_rss(r, True) for r in trials),
                    sampled_pipeline_rss_max_mib=max(sampled_rss(r) for r in trials),
                    peak_job_commit_max_mib=max(r["resources"]["peak_job_committed_bytes"] for r in trials) / 2**20,
                    cpu_median_seconds=median(trials, lambda r: r["resources"]["cpu_user_seconds"] + r["resources"]["cpu_kernel_seconds"]),
                    wall_time_range_seconds=[min(r["wall_seconds"] for r in trials), max(r["wall_seconds"] for r in trials)],
                    statuses=[r["status"] for r in trials],
                ))
    return summaries


def resource_tables(cases, records):
    summaries = resource_summaries(cases, records)
    (EVIDENCE / "resource-summary.json").write_text(json.dumps(summaries, indent=2) + "\n")
    caption = "Opening/saving observations on two hosts. Cells show median total seconds above maximum memory across three attempts. Host blocks specify commit caps and memory metrics. The timeout is 900 s."
    lines = table_start(caption, "lrrrrrr", label="tab:case-study-resources")
    lines.append(r"Endpoint & Original & SheetShrink & NXPowerLite & XLStylesTool & \shortstack{Check\\Performance} & Inquire \\")
    lines.append(r"\multicolumn{7}{l}{\textbf{Ryzen 5 4500U laptop: 4 GiB commit cap; library working set (MiB)}} \\")
    for case in cases[:3]:
        if case["case_id"] == "finance":
            continue
        lines.extend([r"\hline", r"\multicolumn{7}{l}{\textbf{" + CASES[case["case_id"]] + r"}} \\"])
        for library, title in LIBRARIES.items():
            group = [r for r in summaries if r["case"] == case["case_id"] and r["library"] == library]
            if not group:
                continue
            original = group[0]
            completed = [r for r in group if r["completed"] == r["attempts"]]
            time_baseline = original["end_to_end_median_seconds"]
            rss_baseline = original["sampled_library_rss_max_mib"]
            time_scale = max([1] + [r["end_to_end_median_seconds"] / time_baseline for r in completed]) if time_baseline else 1
            rss_scale = max([1] + [r["sampled_library_rss_max_mib"] / rss_baseline for r in completed])
            values = [title]
            for row in group:
                total = row["end_to_end_median_seconds"]
                rss = row["sampled_library_rss_max_mib"]
                time_display = seconds(total)
                rss_display = f"{rss:.0f}"
                if not row["completed"]:
                    assert set(row["statuses"]) == {"memory_limit"}
                    time_display = cell("M", "caseadverse")
                elif row["completed"] == row["attempts"] and original["completed"] == original["attempts"]:
                    time_display = paired_bar(time_display, total, time_baseline, time_scale, width=2.2)
                    rss_display = paired_bar(rss_display, rss, rss_baseline, rss_scale, width=2.2)
                elif row["completed"] != row["attempts"]:
                    time_display += f" ({row['completed']}/{row['attempts']})"
                values.append(r"\shortstack[r]{" + time_display + r"\\" + rss_display + "}")
            lines.append(" & ".join(values) + r" \\")
    note = r"SheetShrink includes fresh cleanup through library exit. Other columns measure processing of prepared outputs, with their cleanup excluded. Completed cells have three successful attempts. M means all three attempts stopped at the memory limit, with reached memory retained below. Bars compare original (gray) and measured value (blue at or below, orange above) on a common linear scale per endpoint and metric, requiring completed originals. RSS covers library phases, with nominal 20 ms sampling. Pipeline working sets, OS job-commit peaks, cleanup/open/save phases, stopped elapsed times, CPU, ranges and all 288 receipts remain in the artifact. The larger-host block reports aggregate process-tree commit in GiB, including cleanup. Dashes indicate unmeasured conditions."
    lines.extend((EVIDENCE / "larger-host-resources.tex").read_text().splitlines())
    lines.extend(table_end(note))
    return "\n".join(lines) + "\n"


def main():
    cases = json.loads((EVIDENCE / "sizes.json").read_text())
    native_rows = json.loads((EVIDENCE / "matched-native-summary.json").read_text())
    with gzip.open(PAPER / "evidence/bounded-open-save/trials.jsonl.gz", "rt") as stream:
        records = [json.loads(line) for line in stream]
    assert len(records) == 288
    (EVIDENCE / "transformations.tex").write_text(transformation_table(cases, native_rows))
    (EVIDENCE / "resources.tex").write_text(resource_tables(cases, records))


if __name__ == "__main__":
    main()
