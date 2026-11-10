"""Repeat the normal SheetShrink CLI on a frozen, hash-verified selection."""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import random
import statistics
import subprocess
import time

import psutil


def main():
    args = parse_arguments()
    inputs = read_verified_inputs(args.selection)
    measure = load_measurement_helper(args.metrics_helper)
    protocol = prepare_protocol(args)
    records = read_records(args.output)
    run_repetitions(args, inputs, measure, records)
    publish_summary(args.output, inputs, records, protocol)


def parse_arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--selection", type=Path, required=True)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--metrics-helper", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--repetitions", type=int, default=9)
    parser.add_argument("--timeout", type=int, default=600)
    args = parser.parse_args()
    if args.repetitions < 3:
        parser.error("At least three measured repetitions are required")
    return args


def digest(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def read_verified_inputs(selection):
    inputs = json.loads(selection.read_text(encoding="utf-8-sig"))
    for item in inputs:
        if digest(item["input_path"]) != item["sha256"]:
            raise ValueError(f"Original identity mismatch: {item['file']}")
    return inputs


def load_measurement_helper(path):
    spec = importlib.util.spec_from_file_location("frozen_process_metrics", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.run_measured


def prepare_protocol(args):
    args.output.mkdir(parents=True, exist_ok=True)
    protocol = {"schema_version": 1, "selection_sha256": digest(args.selection),
                "binary_sha256": digest(args.binary), "binary": str(args.binary),
                "metrics_helper_sha256": digest(args.metrics_helper),
                "warmup_rounds": 1, "measured_rounds": args.repetitions,
                "order": "independently shuffled rounds with seed 20261110 + round",
                "cache": "warm filesystem cache, no cache eviction",
                "wall_scope": "fresh CLI process creation through exit, output hashing excluded",
                "parallelism": "one workbook, one Rayon worker, one writer",
                "cleanup_policy": "default", "timeout_seconds": args.timeout,
                "memory_sampling_seconds": 0.01, "host_id": os.environ["COMPUTERNAME"]}
    target = args.output / "protocol.json"
    if target.exists() and json.loads(target.read_text()) != protocol:
        raise ValueError("Existing experiment has a different protocol")
    target.write_text(json.dumps(protocol, indent=2), encoding="utf-8")
    return protocol


def read_records(output):
    path = output / "trials.jsonl"
    return [json.loads(line) for line in path.read_text().splitlines()] if path.exists() else []


def concurrent_work():
    names = {"excel.exe", "deltasieve.exe", "nxpowerlite.exe", "xlstylestooluwp.exe",
             "excel-check-performance.exe", "cargo.exe", "rustc.exe"}
    return [item.info for item in psutil.process_iter(["pid", "name"])
            if (item.info["name"] or "").lower() in names]


def run_repetitions(args, inputs, measure, records):
    finished = {(row["corpus"], row["rank"], row["round"]) for row in records}
    for number in range(args.repetitions + 1):
        order = list(inputs)
        random.Random(20261110 + number).shuffle(order)
        for item in order:
            identity = (item["corpus"], item["rank"], number)
            if identity in finished:
                continue
            if psutil.virtual_memory().available < 512 * 1024 * 1024:
                raise RuntimeError("Less than 512 MiB physical RAM is available")
            record = run_trial(args, item, number, measure)
            with (args.output / "trials.jsonl").open("a", encoding="utf-8") as stream:
                stream.write(json.dumps(record) + "\n")
                stream.flush()
                os.fsync(stream.fileno())
            records.append(record)
        print(json.dumps({"completed_round": number, "total_trials": len(records)}), flush=True)


def run_trial(args, item, number, measure):
    directory = args.output / "trials" / str(number) / item["corpus"] / str(item["rank"])
    directory.mkdir(parents=True, exist_ok=True)
    destination = args.output / "variants" / item["corpus"] / str(item["rank"])
    destination.mkdir(parents=True, exist_ok=True)
    os.environ["EXCEL_SHRINK_LOG_DIR"] = str(directory / "logs")
    command = [str(args.binary), item["input_path"], str(destination),
               "--force-overwrite", "--disable-multiprocessing", "--rayon-threads", "1",
               "--writer-workers", "1", "--max-active-workbooks", "1",
               "--console-log-level", "error", "--file-log-level", "error"]
    record = dict(item, round=number, warmup=number == 0, command=command,
                  started_utc=time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                  concurrent_work=concurrent_work())
    try:
        record["measurement"] = measure(command, directory, args.timeout)
        output = destination / item["file"]
        record.update(status="completed", output_path=str(output),
                      output_sha256=digest(output), output_bytes=output.stat().st_size)
    except (subprocess.SubprocessError, OSError) as error:
        record.update(status="failed", error=str(error))
        receipt = directory / "process-metrics.json"
        if receipt.exists():
            record["measurement"] = json.loads(receipt.read_text())
    return record


def publish_summary(output, inputs, records, protocol):
    summaries = []
    for item in inputs:
        trials = [row for row in records if row["corpus"] == item["corpus"] and
                  row["rank"] == item["rank"] and not row["warmup"]]
        completed = [row for row in trials if row["status"] == "completed"]
        summary = dict(item, measured_trials=len(trials), successful_trials=len(completed))
        if completed:
            times = [row["measurement"]["seconds"] for row in completed]
            quartiles = statistics.quantiles(times, n=4, method="inclusive") if len(times) > 1 else times * 3
            summary.update(median_seconds=statistics.median(times), q1_seconds=quartiles[0],
                           q3_seconds=quartiles[2], min_seconds=min(times), max_seconds=max(times),
                           output_path=completed[-1]["output_path"],
                           output_sha256=completed[-1]["output_sha256"],
                           output_bytes=completed[-1]["output_bytes"],
                           distinct_output_hashes=len({row["output_sha256"] for row in completed}))
        summaries.append(summary)
    result = {"protocol": protocol, "records": summaries}
    (output / "SheetShrink-Top49-Medians.json").write_text(json.dumps(result, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
