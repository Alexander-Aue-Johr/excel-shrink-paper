import json
import platform
import subprocess
import time
from pathlib import Path

from .common import RunLock, long_path, read_json, sha256, utc_now, write_json
from .release import accepted_outputs


MAXIMUM = [
    "--remove-textless-default-font-representatives",
    "--remove-textless-style-representatives",
    "--remove-textless-column-style-representatives",
    "--remove-unreferenced-whitespace-cells",
]
PROFILES = {"default": [], "maximum": MAXIMUM, "too-aggressive": MAXIMUM + ["--cell-formula-references-only"]}


def _validate_batch_report(report, root, run, inputs):
    expected = {(root / relative).resolve(): (run / relative).resolve() for relative in inputs}
    seen = set()
    counts = {"total": len(expected), "processed": 0, "preserved": 0, "skipped": 0, "failed": 0}
    if report.get("schema_version") != 1:
        raise RuntimeError("The batch report has an unsupported schema.")
    for file in report["files"]:
        source, output = Path(file["input"]).resolve(), Path(file["output"]).resolve()
        if source in seen or expected.get(source) != output:
            raise RuntimeError("The batch report has an unexpected input/output mapping: " + str(source))
        seen.add(source)
        if file["status"] not in ("processed", "preserved"):
            raise RuntimeError("A benchmark input was skipped or failed: " + str(source))
        counts[file["status"]] += 1
    if seen != set(expected) or report.get("summary") != counts:
        raise RuntimeError("The batch report does not account for every input.")
    actual = {
        (run / path.relative_to(long_path(run))).resolve()
        for path in long_path(run / "xlsx").rglob("*") if path.is_file()
    }
    if actual != set(expected.values()):
        raise RuntimeError("The benchmark output directory differs from its batch report.")


def benchmark(root, executable, destination, repeats=3):
    root = Path(root).resolve()
    executable, destination = Path(executable).resolve(), Path(destination).resolve()
    with RunLock(root):
        if destination.exists():
            raise ValueError("Use a new benchmark directory to avoid skipped or overwritten workbooks.")
        inputs = accepted_outputs(root)
        binary_hash = sha256(executable)
        destination.mkdir(parents=True)
        write_json(destination / "run.json", {
            "started_utc": utc_now(), "executable": str(executable), "binary_sha256": binary_hash,
            "input_manifest_sha256": sha256(root / "conversions.jsonl"),
            "input_files": len(inputs), "input_bytes": sum(long_path(root / p).stat().st_size for p in inputs),
            "profiles": PROFILES, "repeats": repeats, "machine": platform.platform(),
            "cache_policy": "No OS cache reset; profile order rotates between repetitions.",
            "timing": "One prebuilt CLI process per corpus and profile. Input hashing and output hashing are outside the wall time.",
        })
        measurements = []
        profiles = list(PROFILES)
        for repeat in range(repeats):
            order = profiles[repeat % len(profiles):] + profiles[:repeat % len(profiles)]
            for profile in order:
                run = destination / ("{:02d}-{}".format(repeat + 1, profile))
                run.mkdir()
                report = run / "batch.json"
                command = [str(executable), str(root / "xlsx"), str(run / "xlsx"), "--recursive", "--batch-report", str(report), *PROFILES[profile]]
                print("Benchmark {}: repetition {}".format(profile, repeat + 1), flush=True)
                with (run / "stdout.log").open("wb") as stdout, (run / "stderr.log").open("wb") as stderr:
                    started = time.perf_counter()
                    process = subprocess.run(command, cwd=run, stdout=stdout, stderr=stderr, check=False)
                    elapsed = time.perf_counter() - started
                measurement = {"profile": profile, "repetition": repeat + 1, "command": command,
                               "wall_seconds": elapsed, "exit_code": process.returncode}
                write_json(run / "measurement.json", measurement)
                measurements.append(measurement)
                write_json(destination / "measurements.json", measurements)
                if process.returncode != 0 or not report.exists():
                    raise RuntimeError("A batch failed; see " + str(run))
                result = read_json(report)
                _validate_batch_report(result, root, run, inputs)
                measurement["pipeline_seconds"] = result["elapsed_ms"] / 1000
                measurement["outcomes"] = result["summary"]
                write_json(run / "measurement.json", measurement)
                write_json(destination / "measurements.json", measurements)
                variants = []
                for relative, digest in sorted(inputs.items()):
                    output = run / relative
                    variants.append({"input_path": relative, "input_sha256": digest,
                                     "output_path": relative, "output_sha256": sha256(output), "output_bytes": long_path(output).stat().st_size})
                with (run / "variants.jsonl").open("w", encoding="utf-8") as stream:
                    for variant in variants:
                        stream.write(json.dumps(variant) + "\n")
                if sha256(executable) != binary_hash:
                    raise RuntimeError("The benchmark executable changed during the experiment.")
        print("Benchmark complete: " + str(destination))
