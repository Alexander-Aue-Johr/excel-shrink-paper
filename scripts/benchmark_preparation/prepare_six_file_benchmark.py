"""Stage the frozen six-file panel and measure the pinned SheetShrink executable."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile


def parse_arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archive", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--engine-source", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, default=Path(__file__).resolve().parents[2] /
                        "papers/research-audit/comparator-benchmark-20261109/panel.json")
    return parser.parse_args()


def sha256(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def stage_originals(arguments, panel):
    if arguments.archive.stat().st_size != panel["archive"]["bytes"] or sha256(arguments.archive) != panel["archive"]["sha256"]:
        raise ValueError("Source archive identity mismatch")
    arguments.output_root.mkdir(parents=True, exist_ok=False)
    inputs = arguments.output_root / "inputs"
    inputs.mkdir()
    with ZipFile(arguments.archive) as archive:
        for workbook in panel["workbooks"]:
            name = workbook["file"]
            if Path(name).name != name:
                raise ValueError("Panel filenames must be flat")
            target = inputs / name
            with archive.open(workbook["archive_member"]) as source, target.open("xb") as destination:
                shutil.copyfileobj(source, destination)
            if target.stat().st_size != workbook["input_bytes"] or sha256(target) != workbook["input_sha256"]:
                raise ValueError(f"Extracted input identity mismatch: {name}")
            for tool in ("nxpowerlite", "check-performance", "xlstyles"):
                directory = arguments.output_root / "comparators" / tool / "inputs"
                directory.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(target, directory / name)
    results = arguments.output_root / "results"
    results.mkdir()
    (results / "provisional-panel.json").write_text(json.dumps({"destatis_candidates": panel["workbooks"]}, indent=2) + "\n")


def measure_sheetshrink(arguments, panel):
    support = arguments.engine_source / "docs/visual-validation/experiments"
    sys.path.insert(0, str(support))
    from process_metrics import run_measured

    output = arguments.output_root / "comparators/sheetshrink/outputs"
    output.mkdir(parents=True)
    receipt = {"source_commit": panel["source_commit"], "binary_sha256": sha256(arguments.binary),
               "metrics_helper_sha256": sha256(support / "process_metrics.py"),
               "measurement_scope": "one fresh direct SheetShrink process per workbook; includes requested statistics",
               "trials": []}
    for workbook in panel["workbooks"]:
        trial = arguments.output_root / "resources/sheetshrink" / workbook["id"]
        trial.mkdir(parents=True)
        command = [str(arguments.binary.resolve()), str(arguments.output_root / "inputs" / workbook["file"]), str(output),
                   "--rayon-threads", "1", "--writer-workers", "1", "--max-active-workbooks", "1",
                   "--analysis-csv", str(trial / "sheets.csv"), "--batch-report", str(trial / "batch.json"),
                   "--shrink-evidence-json", str(trial / "package.json")]
        metrics = run_measured(command, trial, 600)
        result = output / workbook["file"]
        with ZipFile(result) as package:
            if package.testzip() is not None:
                raise ValueError(f"Invalid output package: {result}")
        if sha256(arguments.output_root / "inputs" / workbook["file"]) != workbook["input_sha256"]:
            raise ValueError("Working input changed during processing")
        record = {"id": workbook["id"], "file": workbook["file"], "command": command,
                  "metrics": metrics, "output_bytes": result.stat().st_size, "output_sha256": sha256(result),
                  "matches_frozen_output": sha256(result) == workbook["output_sha256"]}
        receipt["trials"].append(record)
        (arguments.output_root / "resources/sheetshrink.json").write_text(json.dumps(receipt, indent=2) + "\n")
        print(json.dumps({"id": workbook["id"], "seconds": metrics["seconds"],
                          "matches_frozen_output": record["matches_frozen_output"]}), flush=True)


def main():
    arguments = parse_arguments()
    arguments.output_root = arguments.output_root.resolve()
    panel = json.loads(arguments.manifest.read_text(encoding="utf-8-sig"))
    stage_originals(arguments, panel)
    measure_sheetshrink(arguments, panel)


if __name__ == "__main__":
    main()
