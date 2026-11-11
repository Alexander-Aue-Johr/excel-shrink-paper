"""Render manuscript tables from verified size inventories and retained trial receipts."""
import gzip
import json
import statistics
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PAPER = ROOT / "papers/sheetshrink"
EVIDENCE = PAPER / "evidence/open-save-case-study"
TOOLS = ["Original", "SheetShrink", "NXPowerLite", "XLStylesTool", "CheckPerformance", "Inquire"]
NAMES = {"Original": "Original", "SheetShrink": "SS-F", "NXPowerLite": "NX", "XLStylesTool": "XL", "CheckPerformance": "CP", "Inquire": "IQ"}
CASES = {"wanderungen": "Destatis Migration", "finance": "Destatis Municipal Finance", "qt-rr": "Data.gov QT/RR", "enron": "Enron b2d0dfef"}
LIBRARIES = {"openpyxl": "openpyxl", "pandas": "pandas to XLSX", "pandas-csv": "pandas to CSV", "excel": "Excel COM", "r-openxlsx": "R openxlsx", "r-readxl": "R readxl + writexl"}


def cell(value, state=None):
    if value is None:
        return r"\colorbox{casemissing}{\strut P}"
    if state:
        return rf"\colorbox{{{state}}}{{\strut {value}}}"
    return str(value)


def reduction(size, original):
    if size is None:
        return cell(None)
    percentage = 100 * (1 - size / original)
    color = "casegood" if percentage > 0 else "caseadverse" if percentage < 0 else None
    formatted = f"{percentage:.1f}"
    if 0 < abs(percentage) < 0.1:
        formatted = f"{percentage:.2g}" if abs(percentage) < 0.001 else f"{percentage:.3f}"
    return cell(formatted, color)


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
    lines = [r"\begin{table*}[p]", r"\centering", r"\footnotesize", r"\setlength{\tabcolsep}{4pt}", r"\renewcommand{\arraystretch}{1.08}"]
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
    lines = table_start("Opening/memory case-study inputs and native visual evidence. Sizes are decimal MB. Payload is the sum of all uncompressed ZIP-member sizes. Reduction columns compare each output with its original. Pending evidence is marked P.", "lrrrrrccc", label="tab:case-study-transformations")
    lines.append(r"Variant & ZIP & Red. (\%) & Payload & Red. (\%) & Pixels $D/N$ & Sheets & Geom. & State \\")
    for case in cases:
        original = case["variants"]["Original"]
        lines.extend([r"\hline", r"\multicolumn{9}{l}{\textbf{" + CASES[case["case_id"]] + r"}} \\"])
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
                pixels = cell(f"{visual['different_pixel_observations']:,} / P", "caseadverse")
                sheets = f"{visual['changed_sheets']}/{visual['observed_sheets']}/{visual['total_sheets']}"
                geometry = str(visual["geometry_findings"])
                state = cell("L", "caseadverse")
            else:
                pixels, sheets, geometry = cell(None), cell(None), cell(None)
                state = cell("M" if visual and visual["status"] == "memory_limit" else "F" if visual and visual["status"] == "failed" else "P", "casemissing")
            values = [NAMES[tool], cell(None) if compressed is None else f"{compressed / 1e6:.3f}", reduction(compressed, original["zip_bytes"]), cell(None) if payload is None else f"{payload / 1e6:.2f}", reduction(payload, original["uncompressed_bytes"]), pixels, sheets, geometry, state]
            lines.append(" & ".join(values) + r" \\")
    note = r"SS-F denotes fresh six-worker SheetShrink trial outputs. Their compressed size is recorded consistently across repetitions, while their uncompressed payload and matched native visual evidence remain pending. NX uses the cell-formatting-focused profile. XL retains the tested settings. CP and IQ denote Check Performance and Inquire. Ref. identifies the original baseline. L denotes completed targeted captures with coverage limitations, M a memory-limit stop, F a failed attempt and P pending or unavailable evidence. Sheets gives changed/observed/total worksheets. Geometry counts flagged capture findings. Pixel counts aggregate retained diagnostic observations under policy v12 after paired native rebuild. The total comparable-pixel denominator N remains pending. Different capture scopes and unresolved attribution constrain comparisons. Blue tint accompanies positive size reduction. Orange marks observed differences. All numerical values remain explicit."
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
                cleanup_display = seconds(cleanup) if tool == "SheetShrink" else "--" if tool == "Original" else "prep."
                lines.append(" & ".join([NAMES[tool], cleanup_display, seconds(opened), seconds(saved), display_total, f"{rss_library:.0f}", f"{rss_all:.0f}", f"{commit:.0f}", outcome]) + r" \\")
                summaries.append(dict(case=case["case_id"], library=library, tool=tool, completed=len(completed), attempts=len(trials), cleanup_median_seconds=cleanup, open_median_seconds=opened, save_median_seconds=saved, end_to_end_median_seconds=total, sampled_library_rss_max_mib=rss_library, sampled_pipeline_rss_max_mib=rss_all, peak_job_commit_max_mib=commit, cpu_median_seconds=cpu, wall_time_range_seconds=[min(r["wall_seconds"] for r in trials), max(r["wall_seconds"] for r in trials)], statuses=[r["status"] for r in trials]))
        if index == 2:
            lines.extend([r"\hline", r"\multicolumn{9}{l}{Enron b2d0dfef, Gotham and Windows-VM measurements: P.} \\"])
        note = r"SS-F runs fresh cleanup immediately before the library. Its end-to-end clock spans cleanup launch through library exit. Original and prepared NX/XL/CP/IQ rows report library process lifetime. prep. denotes preparation outside this trial. Open and save measure the retained operation endpoints. End-to-end includes startup/shutdown and is measured directly. RSS lib. is the maximum sampled process-tree working set in importing, opening, saving and completed phases. RSS all includes cleanup and other recorded phases. Sampling is nominally 20 ms. Commit is the Windows job peak. The columns have different counter semantics. Outcome gives completed/attempted. M@t is elapsed time until the memory-limit stop, with missing completion phases marked P. Blue/orange highlights faster/slower SS-F end-to-end time against a fully completed original on the same host and endpoint. Other hosts, missing Enron observations, dispersion and CPU fields remain explicit in the machine-readable artifact."
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
