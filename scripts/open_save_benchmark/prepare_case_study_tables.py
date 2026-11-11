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
        if case["case_id"] == "finance":
            continue
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
    caption = "Opening/saving summary on the Ryzen 5 4500U laptop with 8 GiB RAM. Each cell shows median total seconds above maximum sampled library working set in MiB. Three attempts per condition, 4 GiB job-commit cap, 900 s timeout."
    lines = table_start(caption, "lrrrrrr", label="tab:case-study-resources")
    lines.append(r"Endpoint & Original & Sheet Shrink & NXPowerLite & XLStylesTool & \shortstack{Check\\Performance} & Inquire \\")
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
    note = r"Sheet Shrink includes fresh cleanup through library exit. Other columns measure processing of prepared outputs, with their cleanup excluded. Completed cells have three successful attempts. M means all three attempts stopped at the memory limit, with reached library RSS retained below. Bars compare original (gray) and measured value (blue at or below, orange above) on a common linear scale per endpoint and metric, requiring completed originals. RSS covers library phases, with nominal 20 ms sampling. Pipeline working sets, OS job-commit peaks, cleanup/open/save phases, stopped elapsed times, CPU, ranges and all 288 receipts remain in the artifact. Separate host results follow below."
    lines.extend(table_end(note))
    return "\n".join(lines) + "\n"


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
