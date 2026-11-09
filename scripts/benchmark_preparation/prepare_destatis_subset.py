"""Prepare a historical stress selection and run SheetShrink on verified copies."""

import argparse
import collections
import csv
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import time


WANDERUNGEN = "wanderungen-2010120217005.xlsx"


def main():
    arguments = parse_arguments()
    if (arguments.output_root / "results.jsonl").exists():
        raise FileExistsError("Choose a fresh output root for a new preparation attempt")
    candidates = select_historical_candidates(arguments)
    manifest = stage_verified_inputs(arguments, candidates)
    results = clean_working_copies(arguments, manifest)
    write_rankings(arguments.output_root, results)


def parse_arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus-root", type=Path, required=True)
    parser.add_argument("--source-profile", type=Path, required=True)
    parser.add_argument("--historical-csv", type=Path, required=True)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--source-commit", required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    return parser.parse_args()


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(4 * 1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def select_historical_candidates(arguments):
    totals = collections.defaultdict(lambda: {"original": 0, "cleaned": 0})
    with arguments.historical_csv.open(encoding="utf-8-sig", newline="") as stream:
        for row in csv.DictReader(stream):
            if row["Sheet / Workbook"] == "Workbook":
                continue
            total = totals[row["File"]]
            total["original"] += int(row["Original Size"])
            total["cleaned"] += int(row["Cleaned Size"])
    available = {
        name: values for name, values in totals.items()
        if values["original"] > 0 and (arguments.corpus_root / name).is_file()
    }
    absolute = sorted(available, key=lambda name: (
        -(available[name]["original"] - available[name]["cleaned"]), name))
    relative = sorted(available, key=lambda name: (
        -(1 - available[name]["cleaned"] / available[name]["original"]), name))
    selected = {}
    for absolute_name, relative_name in zip(absolute, relative):
        for name in (absolute_name, relative_name):
            if len(selected) < 100:
                selected.setdefault(name, available[name])
        if len(selected) == 100:
            break
    if len(selected) != 100:
        raise ValueError("A complete selection requires 100 available historical inputs")
    if WANDERUNGEN not in available:
        raise ValueError("Required Wanderungen source is missing")
    selected.setdefault(WANDERUNGEN, available[WANDERUNGEN])
    return [{"name": name, "historical_worksheet_xml": values}
            for name, values in selected.items()]


def stage_verified_inputs(arguments, candidates):
    profile = json.loads(arguments.source_profile.read_text(encoding="utf-8-sig"))
    identities = {item["name"]: item for item in profile["workbooks"]}
    input_root = arguments.output_root / "inputs"
    input_root.mkdir(parents=True, exist_ok=True)
    for candidate in candidates:
        name = candidate["name"]
        if Path(name).name != name:
            raise ValueError("Corpus filenames must be flat")
        identity = identities[name]
        source = arguments.corpus_root / name
        if source.stat().st_size != identity["bytes"] or sha256(source) != identity["sha256"]:
            raise ValueError(f"Source identity mismatch: {name}")
        target = input_root / name
        if not target.exists():
            shutil.copyfile(source, target)
        if sha256(target) != identity["sha256"]:
            raise ValueError(f"Working-copy identity mismatch: {name}")
        candidate.update(input_bytes=identity["bytes"], input_sha256=identity["sha256"])
    manifest = {
        "selection": "Alternating historical absolute and relative worksheet-XML reduction ranks; 100 distinct inputs plus Wanderungen if absent",
        "population_role": "Purposive historical stress selection",
        "historical_csv_sha256": sha256(arguments.historical_csv),
        "source_profile_sha256": sha256(arguments.source_profile),
        "source_archive": profile["archive"],
        "software_commit": arguments.source_commit,
        "binary_sha256": sha256(arguments.binary),
        "configuration": {"rayon_threads": 1, "writer_workers": 1, "max_active_workbooks": 1,
                          "cleanup_policy": "ordinary production defaults; analysis CSV adds statistics"},
        "inputs": candidates,
    }
    (arguments.output_root / "selection.json").write_text(
        json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    return manifest


def clean_working_copies(arguments, manifest):
    output_root = arguments.output_root / "sheetshrink"
    evidence_root = arguments.output_root / "evidence"
    output_root.mkdir(exist_ok=True)
    evidence_root.mkdir(exist_ok=True)
    results = []
    with (arguments.output_root / "results.jsonl").open("x", encoding="utf-8") as journal:
        for index, candidate in enumerate(manifest["inputs"], 1):
            name = candidate["name"]
            evidence = evidence_root / f"{index:03d}"
            evidence.mkdir(exist_ok=True)
            command = [str(arguments.binary), str(arguments.output_root / "inputs" / name),
                       str(output_root), "--rayon-threads", "1", "--writer-workers", "1",
                       "--max-active-workbooks", "1", "--analysis-csv", str(evidence / "sheets.csv"),
                       "--batch-report", str(evidence / "batch.json"),
                       "--shrink-evidence-json", str(evidence / "package.json")]
            started = time.perf_counter()
            with (evidence / "console.log").open("w", encoding="utf-8") as log:
                completed = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT, check=False)
            output = output_root / name
            result = {"index": index, "file": name, "input_bytes": candidate["input_bytes"],
                      "input_sha256": candidate["input_sha256"], "exit_code": completed.returncode,
                      "elapsed_seconds": time.perf_counter() - started, "command": command}
            if completed.returncode == 0 and output.is_file():
                result.update(output_bytes=output.stat().st_size, output_sha256=sha256(output))
                result["saved_bytes"] = result["input_bytes"] - result["output_bytes"]
                result["reduction_percent"] = 100 * result["saved_bytes"] / result["input_bytes"]
            results.append(result)
            journal.write(json.dumps(result) + "\n")
            journal.flush()
            print(json.dumps({key: result.get(key) for key in
                              ("index", "file", "exit_code", "reduction_percent")}), flush=True)
    return results


def write_rankings(root, results):
    completed = [result for result in results if "output_bytes" in result]
    fields = ("index", "file", "input_bytes", "output_bytes", "saved_bytes",
              "reduction_percent", "elapsed_seconds", "input_sha256", "output_sha256")
    for metric in ("saved_bytes", "reduction_percent"):
        ranking = sorted(completed, key=lambda row: (-row[metric], row["file"]))
        with (root / f"ranked-by-{metric}.csv").open("w", encoding="utf-8", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=fields, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(ranking)
    print(json.dumps({"attempted": len(results), "outputs": len(completed)}), flush=True)


if __name__ == "__main__":
    main()
