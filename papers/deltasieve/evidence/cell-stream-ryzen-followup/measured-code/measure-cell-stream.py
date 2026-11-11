"""Measure Figure 5 through its opt-in Rust test, one fresh process per sample.

Input preparation, ZIP decompression and independent XML checks are untimed.
Runs use normal filesystem caching. This is not a cold-storage benchmark.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import platform
import random
import re
import statistics
import subprocess
import time
import xml.etree.ElementTree as ET
import zipfile


TEST = "regression::targeted_comparison::cell_stream::verification::measure_figure5_stream_comparison"


def file_digest(path):
    with path.open("rb") as input_file:
        return hashlib.file_digest(input_file, "sha256").hexdigest()


def generate_sheet(path, cells, variant):
    """Use 100 columns so all scales remain within Excel's coordinate limits."""
    with path.open("wb", buffering=1024 * 1024) as output:
        output.write(b"<worksheet><sheetData>")
        for index in range(cells):
            row, column = divmod(index, 100)
            if column == 0:
                output.write(f'<row r="{row + 1}">'.encode())
            if variant != "empty":
                letters = ""
                number = column + 1
                while number:
                    number, digit = divmod(number - 1, 26)
                    letters = chr(65 + digit) + letters
                value = index
                if variant == "dense" or (variant == "sparse" and index == cells // 2):
                    value += 1
                output.write(f'<c r="{letters}{row + 1}"><v>{value}</v></c>'.encode())
            if column == 99 or index == cells - 1:
                output.write(b"</row>")
        output.write(b"</sheetData></worksheet>")


def real_sheet_oracle(path):
    """Count cells and hash their attributes/text with an independent XML parser.

    This oracle checks our single-value mutation. It does not implement the
    production fingerprint's style or XML-structure equivalence rules.
    """
    digest = hashlib.sha256()
    count = 0
    for event, element in ET.iterparse(path, events=("end",)):
        local = element.tag.rsplit("}", 1)[-1]
        if local == "c":
            digest.update(repr((element.attrib, tuple(element.itertext()))).encode("utf-8"))
            element.clear()
            count += 1
        elif local == "row":
            element.clear()
    return {"cells": count, "cell_value_digest": digest.hexdigest()}


def copy_with_one_value_change(source, target):
    """Change exactly one numeric value byte, preserving XML structure."""
    changed = False
    offset = 0
    with source.open("rb") as input_file, target.open("wb") as output:
        pending = b""
        while chunk := input_file.read(1024 * 1024):
            pending += chunk
            if not changed:
                match = re.search(rb"<v>([0-9])", pending)
                if match:
                    position = match.start(1)
                    original = pending[position]
                    replacement = 48 + (original - 48 + 1) % 10
                    pending = pending[:position] + bytes([replacement]) + pending[position + 1:]
                    mutation = {"byte_offset": offset + position, "before": chr(original), "after": chr(replacement)}
                    changed = True
            flushed = pending[:-128]
            output.write(flushed)
            offset += len(flushed)
            pending = pending[-128:]
        output.write(pending)
    if not changed:
        raise ValueError(f"No numeric cell value in {source}")
    return mutation


def prepare_inputs(directory, sizes, enron_directory):
    inputs = directory / "inputs"
    inputs.mkdir()
    cases = []
    for cells in sizes:
        before = inputs / f"synthetic-{cells}-before.xml"
        generate_sheet(before, cells, "before")
        for variant in ("identical", "sparse", "dense", "empty"):
            after = before if variant == "identical" else inputs / f"synthetic-{cells}-{variant}.xml"
            if after != before:
                generate_sheet(after, cells, variant)
            expected = {"added": 0, "removed": cells if variant == "empty" else 0,
                        "changed": cells if variant == "dense" else int(variant == "sparse")}
            case = {"name": f"synthetic-{cells}-{variant}", "before": str(before),
                    "after": str(after), "cells_before": cells, "expected": expected,
                    "collect": False}
            cases.append(case)
            if variant == "dense":
                cases.append({**case, "name": f"synthetic-{cells}-dense-collected", "collect": True})
    if enron_directory:
        workbooks = sorted(enron_directory.glob("*.xlsx"), key=lambda path: path.stat().st_size, reverse=True)[:3]
        for workbook in workbooks:
            with zipfile.ZipFile(workbook) as archive:
                sheets = [entry for entry in archive.infolist()
                          if re.fullmatch(r"xl/worksheets/sheet[0-9]+\.xml", entry.filename)]
                sheet = max(sheets, key=lambda entry: entry.file_size)
                before = inputs / f"enron-{workbook.stem}.xml"
                with archive.open(sheet) as input_file, before.open("wb") as output:
                    while chunk := input_file.read(1024 * 1024):
                        output.write(chunk)
                styles = inputs / f"enron-{workbook.stem}-styles.xml"
                styles.write_bytes(archive.read("xl/styles.xml"))
                provenance = {"workbook": str(workbook), "sheet": sheet.filename,
                              "sheet_crc32": sheet.CRC, "sheet_xml_bytes": sheet.file_size,
                              "workbook_sha256": file_digest(workbook)}
            oracle = real_sheet_oracle(before)
            after = inputs / f"enron-{workbook.stem}-one-change.xml"
            mutation = copy_with_one_value_change(before, after)
            after_oracle = real_sheet_oracle(after)
            assert oracle["cells"] == after_oracle["cells"]
            assert oracle["cell_value_digest"] != after_oracle["cell_value_digest"]
            for variant in ("identical", "one-change"):
                cases.append({"name": f"enron-{workbook.stem[:12]}-{variant}", "before": str(before),
                              "after": str(before if variant == "identical" else after),
                              "styles": str(styles), "cells_before": oracle["cells"],
                              "collect": False, "provenance": provenance, "oracle": oracle,
                              "mutation": mutation, "expected": {"added": 0, "removed": 0,
                                                                    "changed": int(variant == "one-change")}})
    return cases


def measure(binary, case, scratch):
    environment = os.environ.copy()
    environment.update(CELL_STREAM_BEFORE=case["before"], CELL_STREAM_AFTER=case["after"],
                       CELL_STREAM_SCRATCH=str(scratch), CELL_STREAM_COLLECT=str(int(case["collect"])))
    environment.pop("CELL_STREAM_STYLES", None)
    if "styles" in case:
        environment["CELL_STREAM_STYLES"] = case["styles"]
    started = time.perf_counter()
    result = subprocess.run([str(binary), "--exact", TEST, "--ignored", "--nocapture", "--test-threads=1"],
                            env=environment, capture_output=True, text=True, timeout=300)
    if result.returncode:
        raise RuntimeError(result.stdout + result.stderr)
    line = next(line for line in result.stdout.splitlines() if "FIGURE5_RESULT " in line)
    metrics = json.loads(line.split("FIGURE5_RESULT ", 1)[1])
    for key, expected in case["expected"].items():
        assert metrics[key] == expected, (case["name"], key, metrics[key], expected)
    assert not list(scratch.iterdir()), "Fingerprint files were not released"
    return {"case": case["name"], "collect": case["collect"],
            "process_wall_ms": (time.perf_counter() - started) * 1000, **metrics}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sizes", type=int, nargs="+", default=[10_000, 100_000, 1_000_000])
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--enron-directory", type=Path)
    args = parser.parse_args()
    if os.name != "nt":
        parser.error("Process memory measurements currently require Windows")
    if args.repeats < 1 or any(size < 1 or size > 104_857_600 for size in args.sizes):
        parser.error("Use positive repeats and cell counts within the generated Excel grid")
    args.output.mkdir(parents=True, exist_ok=False)
    scratch = args.output / "scratch"
    scratch.mkdir()
    cases = prepare_inputs(args.output, args.sizes, args.enron_directory)
    (args.output / "manifest.json").write_text(json.dumps({
        "cases": cases, "repeats": args.repeats, "cache_policy": "normal filesystem caching; inputs generated before measurement",
        "platform": platform.platform(), "processor": platform.processor(), "logical_cpus": os.cpu_count(),
        "rustc": subprocess.check_output(["rustc", "-Vv"], text=True),
        "revision": subprocess.check_output(["git", "rev-parse", "HEAD"], text=True).strip(),
        "binary_sha256": file_digest(args.binary),
        "rustflags": os.environ.get("RUSTFLAGS", ""),
        "source_diff_sha256": hashlib.sha256(subprocess.check_output(["git", "diff"])).hexdigest(),
        "source_files_sha256": {str(path): file_digest(path) for path in (
            Path("excel-shrink-cli/src/regression/targeted_comparison/cell_stream.rs"),
            Path("excel-shrink-cli/src/regression/targeted_comparison/cell_stream/verification.rs"),
            Path("excel-shrink-cli/src/regression/streaming_xml.rs"), Path("Cargo.lock"),
        )},
        "processor_details": subprocess.check_output([
            "powershell", "-NoProfile", "-Command",
            "Get-CimInstance Win32_Processor | Select-Object Name,NumberOfLogicalProcessors | ConvertTo-Json -Compress"
        ], text=True).strip(),
    }, indent=2), encoding="utf-8")
    jobs = [(case, repeat) for case in cases for repeat in range(args.repeats)]
    random.Random(5).shuffle(jobs)
    results = []
    with (args.output / "samples.jsonl").open("w", encoding="utf-8") as output:
        for case, repeat in jobs:
            sample = {"repeat": repeat, **measure(args.binary, case, scratch)}
            results.append(sample)
            output.write(json.dumps(sample) + "\n")
            output.flush()
            print(f'{len(results)}/{len(jobs)} {case["name"]}: {sample["before_ms"] + sample["after_ms"]:.1f} ms writing, {sample["merge_ms"]:.1f} ms merging', flush=True)
    summary = []
    for case in cases:
        samples = [sample for sample in results if sample["case"] == case["name"]]
        summary.append({"case": case["name"], "cells_before": case["cells_before"],
                        **{key: statistics.median(sample[key] for sample in samples)
                           for key in ("before_ms", "after_ms", "merge_ms", "scratch_bytes", "retained_capacity_bytes")},
                        "peak_working_set_bytes": max(sample["merge_memory"]["peak_working_set_bytes"] for sample in samples),
                        "peak_private_commit_bytes": max(sample["merge_memory"]["peak_pagefile_bytes"] for sample in samples)})
    (args.output / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
