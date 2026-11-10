"""Join frozen comparator receipts to repeated SheetShrink observations by SHA-256."""
import argparse
import csv
import json
from pathlib import Path
import statistics


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sheetshrink", type=Path, required=True)
    parser.add_argument("--comparator", type=Path, required=True)
    parser.add_argument("--trials", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--latex", type=Path)
    args = parser.parse_args()
    repeated = json.loads(args.sheetshrink.read_text(encoding="utf-8-sig"))
    comparator = json.loads(args.comparator.read_text(encoding="utf-8-sig"))
    pairs = pair_observations(repeated, comparator)
    add_resource_summaries(pairs, args.trials)
    summaries = write_pairs(pairs, args.output)
    if args.latex:
        write_resource_table(summaries, args.latex)


def pair_observations(repeated, comparator):
    if repeated["protocol"]["host_id"] != comparator["host_id"]:
        raise ValueError("Comparator and SheetShrink measurements must use the same host")
    baseline = {(row["corpus"], row["sha256"]): row for row in repeated["records"]}
    pairs = []
    for row in comparator["records"]:
        if not row["eligible_for_formatting_comparison"]:
            continue
        match = baseline[(row["corpus"], row["input_sha"])]
        if match["successful_trials"] != repeated["protocol"]["measured_rounds"]:
            raise ValueError(f"Incomplete SheetShrink repetitions for {row['file']}")
        primary = row["attempt_history"][0]
        if primary["state"] not in ("converted", "no_change"):
            raise ValueError("A successful retry cannot replace the primary timing observation")
        resources = json.loads(primary["evidence"])["resources"]
        pairs.append({"host_id": repeated["protocol"]["host_id"],
                      "comparator": comparator.get("tool_label", "NXPowerLite"), "corpus": row["corpus"], "rank": row["rank"],
                      "file": row["file"], "input_sha256": row["input_sha"],
                      "comparator_seconds": resources["wall_seconds"],
                      "sheetshrink_seconds": match["median_seconds"],
                      "sheetshrink_q1_seconds": match["q1_seconds"],
                      "sheetshrink_q3_seconds": match["q3_seconds"],
                      "sheetshrink_repetitions": match["successful_trials"],
                      "comparator_repetitions": 1,
                      "comparator_endpoint": "GUI orchestration through output confirmation and initial verification",
                      "sheetshrink_endpoint": "fresh CLI process creation through exit",
                      "comparator_cpu_seconds": resources["cpu_user_seconds"] + resources["cpu_kernel_seconds"],
                      "comparator_sampled_peak_working_set_bytes": resources["sampled_peak_working_set_bytes"],
                      "comparator_sampled_peak_private_bytes": resources["sampled_peak_private_bytes"],
                      "sheetshrink_output_sha256": match["output_sha256"],
                      "comparator_output_sha256": row["output_sha"],
                      "visual_validation": "pending"})
    return pairs


def add_resource_summaries(pairs, path):
    trials = [json.loads(line) for line in path.read_text().splitlines()]
    for pair in pairs:
        selected = [row["measurement"]["resources"] for row in trials
                    if row["corpus"] == pair["corpus"] and row["sha256"] == pair["input_sha256"]
                    and not row["warmup"] and row["status"] == "completed"]
        if len(selected) != pair["sheetshrink_repetitions"]:
            raise ValueError("Raw trial count differs from the summarized repetitions")
        pair["sheetshrink_cpu_seconds"] = statistics.median(
            row["cpu_user_seconds"] + row["cpu_kernel_seconds"] for row in selected)
        for field in ("observed_os_peak_working_set_bytes", "sampled_peak_private_bytes"):
            values = [row[field] for row in selected if row.get(field) is not None]
            pair[f"sheetshrink_{field}"] = statistics.median(values) if len(values) == len(selected) else None


def write_pairs(pairs, output):
    if not pairs:
        raise ValueError("No eligible paired observations remain")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(pairs[0]))
        writer.writeheader()
        writer.writerows(pairs)
    summaries = []
    for corpus in ("destatis", "enron", "datagov"):
        rows = [row for row in pairs if row["corpus"] == corpus]
        if not rows:
            continue
        summaries.append({"corpus": corpus, "paired_workbooks": len(rows),
                          "comparator_median_seconds": statistics.median(row["comparator_seconds"] for row in rows),
                          "sheetshrink_median_seconds": statistics.median(row["sheetshrink_seconds"] for row in rows),
                          "median_paired_ratio": statistics.median(row["comparator_seconds"] / row["sheetshrink_seconds"] for row in rows),
                          "sheetshrink_cpu_seconds": statistics.median(row["sheetshrink_cpu_seconds"] for row in rows),
                          "comparator_cpu_seconds": statistics.median(row["comparator_cpu_seconds"] for row in rows),
                          "sheetshrink_peak_working_set_mib": statistics.median(row["sheetshrink_observed_os_peak_working_set_bytes"] for row in rows) / 2**20,
                          "comparator_sampled_working_set_mib": statistics.median(row["comparator_sampled_peak_working_set_bytes"] for row in rows) / 2**20})
    output.with_suffix(".json").write_text(json.dumps(summaries, indent=2), encoding="utf-8")
    print(json.dumps(summaries, indent=2))
    return summaries


def write_resource_table(summaries, output):
    names = {"destatis": "Destatis", "enron": "Enron", "datagov": "Data.gov"}
    rows = []
    for row in summaries:
        rows.append(f"{names[row['corpus']]} & {row['paired_workbooks']} & "
                    f"{row['sheetshrink_cpu_seconds']:.2f} & {row['comparator_cpu_seconds']:.2f} & "
                    f"{row['sheetshrink_peak_working_set_mib']:.0f} & {row['comparator_sampled_working_set_mib']:.0f} "
                    + r"\\")
    text = r"""\begin{table}[t]
\caption{Median resource observations across paired workbooks on the laptop.
SS denotes SheetShrink, NX denotes NXPowerLite. CPU is process time in seconds.
WS is working set in MiB, using OS peak counters for SS and sampled peaks for NX.}
\label{tab:top49-resources}
\centering
\setlength{\tabcolsep}{3pt}
\begin{tabular}{lrrrrr}
\hline
Corpus & Pairs & CPU SS & CPU NX & WS SS & WS NX \\
\hline
""" + "\n".join(rows) + r"""
\hline
\end{tabular}
\end{table}
"""
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
