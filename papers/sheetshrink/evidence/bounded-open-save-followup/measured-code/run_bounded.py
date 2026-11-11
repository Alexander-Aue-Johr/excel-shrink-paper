"""Run resumable open/save trials inside a hard-limited Windows job."""
import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys
import time

from memory_job import MemoryJob


def main():
    args = parse_arguments()
    sys.path.insert(0, str(args.root / "software/packages"))
    global psutil
    import psutil
    manifest = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    args.root.mkdir(parents=True, exist_ok=True)
    acquire_run_lock(args.root)
    verify_inputs(manifest, args.sheetshrink_binary is not None)
    existing = read_records(args.root)
    completed = {row["trial_id"] for row in existing}
    for repetition in range(args.repetitions):
        for case in manifest["cases"]:
            for variant in case["variants"]:
                for library in args.libraries:
                    identity = f"{case['case_id']}/{variant['tool']}/{library}/{repetition}"
                    if identity in completed:
                        continue
                    wait_for_headroom(args)
                    result = run_trial(args, identity, case, variant, library, repetition)
                    append_record(args.root, result)
                    print(json.dumps({"trial_id": identity, "status": result["status"],
                                      "seconds": result["wall_seconds"],
                                      "peak_commit_mib": result["resources"]["peak_job_committed_bytes"] / 2**20}), flush=True)


def acquire_run_lock(root):
    import msvcrt
    global run_lock
    run_lock = (root / "runner.lock").open("a+b")
    run_lock.write(b"0")
    run_lock.flush()
    run_lock.seek(0)
    try:
        msvcrt.locking(run_lock.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError as error:
        raise RuntimeError("Another benchmark is already using this result directory") from error


def verify_inputs(manifest, fresh_sheetshrink=False):
    for case in manifest["cases"]:
        for variant in case["variants"]:
            if fresh_sheetshrink and variant["tool"] == "SheetShrink":
                continue
            with Path(variant["path"]).open("rb") as stream:
                actual = hashlib.file_digest(stream, "sha256").hexdigest()
            if actual != variant["sha256"]:
                raise ValueError(f"Input hash mismatch: {variant['path']}")


def parse_arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--helper-directory", type=Path, required=True)
    parser.add_argument("--libraries", nargs="+", default=["openpyxl", "pandas", "excel", "r-openxlsx", "r-readxl"])
    parser.add_argument("--repetitions", type=int, default=3)
    parser.add_argument("--limit-mib", type=int, default=4096)
    parser.add_argument("--reserve-mib", type=int, default=0)
    parser.add_argument("--timeout", type=int, default=900)
    parser.add_argument("--sheetshrink-binary", type=Path)
    return parser.parse_args()


def read_records(root):
    path = root / "results.jsonl"
    records = [json.loads(line) for line in path.read_text().splitlines()] if path.exists() else []
    identities = {record["trial_id"] for record in records}
    for receipt in sorted(root.glob("trials/*/*/*/*/receipt.json")):
        record = json.loads(receipt.read_text(encoding="utf-8"))
        if record["trial_id"] not in identities:
            append_record(root, record)
            records.append(record)
            identities.add(record["trial_id"])
    return records


def append_record(root, result):
    with (root / "results.jsonl").open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(result) + "\n")
        stream.flush()
        os.fsync(stream.fileno())


def wait_for_headroom(args):
    if args.reserve_mib == 0:
        return
    started = time.perf_counter()
    while min(psutil.virtual_memory().available, available_commit_bytes()) < (args.limit_mib + args.reserve_mib) * 2**20:
        if time.perf_counter() - started > 120:
            raise RuntimeError("Insufficient RAM or commit headroom to start the next bounded trial")
        time.sleep(1)


def read_state(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def command_for(library, source, destination, state, root):
    if library.startswith("r-"):
        return [str(root / "software/R/bin/Rscript.exe"), "--vanilla", str(Path(__file__).with_name("worker.R")),
                library, source, str(state), str(destination)]
    return [sys.executable, str(Path(__file__).with_name("worker.py")), library, source, str(state), str(destination)]


def process_memory(job, process_peaks):
    private = resident = 0
    for pid in job.process_ids():
        try:
            process = psutil.Process(pid)
            memory = process.memory_info()
            peak = process_peaks.setdefault(str(pid), {"name": process.name(), "private_bytes": 0, "working_set_bytes": 0})
            peak["private_bytes"] = max(peak["private_bytes"], memory.private)
            peak["working_set_bytes"] = max(peak["working_set_bytes"], memory.rss)
            private += memory.private
            resident += memory.rss
        except psutil.Error:
            continue
    return private, resident


def bind_excel(job, state, state_path, existing_excel, started):
    pid = state["excel_pid"]
    process = psutil.Process(pid)
    if pid in existing_excel or process.name().lower() != "excel.exe" or process.create_time() < started:
        raise RuntimeError("Excel process ownership could not be confirmed")
    try:
        job.attach(pid)
    except OSError:
        process.kill()
        process.wait(timeout=10)
        raise
    state_path.with_suffix(".permit").write_text(str(pid), encoding="ascii")




class MemoryStatus(ctypes.Structure):
    _fields_ = [("length", wintypes.DWORD), ("load", wintypes.DWORD)] + [
        (name, ctypes.c_ulonglong) for name in (
            "total_phys", "avail_phys", "total_page", "avail_page",
            "total_virtual", "avail_virtual", "extended")
    ]


def available_commit_bytes():
    memory = MemoryStatus()
    memory.length = ctypes.sizeof(memory)
    if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory)):
        raise ctypes.WinError()
    return memory.avail_page


def monitor(job, process, state_path, args, existing_excel, birth):
    started = time.perf_counter()
    peaks = {}
    process_peaks = {}
    last_private = last_resident = 0
    memory_events = []
    status = None
    excel_bound = False
    minimum_ram = minimum_commit = 2**64
    while process.poll() is None:
        state = read_state(state_path)
        phase = state.get("phase", "starting")
        free_ram = psutil.virtual_memory().available
        free_commit = available_commit_bytes()
        minimum_ram = min(minimum_ram, free_ram)
        minimum_commit = min(minimum_commit, free_commit)
        private, resident = process_memory(job, process_peaks)
        last_private, last_resident = private, resident
        peak = peaks.setdefault(phase, {"private_bytes": 0, "working_set_bytes": 0})
        peak["private_bytes"] = max(peak["private_bytes"], private)
        peak["working_set_bytes"] = max(peak["working_set_bytes"], resident)
        if phase == "excel_binding" and not excel_bound:
            bind_excel(job, state, state_path, existing_excel, birth)
            excel_bound = True
        events = job.memory_events()
        memory_events.extend(events)
        if events:
            status = "memory_limit"
        elif free_ram < args.reserve_mib * 2**20:
            status = "system_ram_reserve"
        elif free_commit < args.reserve_mib * 2**20:
            status = "system_commit_reserve"
        elif time.perf_counter() - started > args.timeout:
            status = "timeout"
        if status:
            job.terminate()
            break
        time.sleep(0.02)
    process.wait(timeout=10)
    memory_events.extend(job.memory_events())
    state = read_state(state_path)
    if status is None:
        status = "memory_limit" if memory_events else state.get("phase", "worker_exit")
        if status == "failed" and (state.get("error_type") == "MemoryError" or "cannot allocate" in state.get("error", "").lower()):
            status = "allocation_failed_under_limit"
    counters = job.counters()
    counters.update(sampled_phase_peaks=peaks, observed_process_peaks=process_peaks, last_private_bytes=last_private,
                    last_working_set_bytes=last_resident, memory_events=memory_events,
                    nominal_sampling_seconds=0.02, minimum_free_ram_bytes=minimum_ram,
                    minimum_commit_headroom_bytes=minimum_commit)
    remaining = job.process_ids()
    if remaining:
        job.terminate()
    return status, state, counters


def run_trial(args, identity, case, variant, library, repetition):
    directory = args.root / "trials" / identity
    directory.mkdir(parents=True, exist_ok=True)
    temporary_directory = directory / "temp"
    temporary_directory.mkdir(exist_ok=True)
    state_path = directory / "state.json"
    state_path.unlink(missing_ok=True)
    state_path.with_suffix(".permit").unlink(missing_ok=True)
    destination = directory / ("csv-output" if library == "pandas-csv" else "saved.xlsx")
    environment = dict(os.environ, BENCHMARK_PACKAGES=str(args.root / "software/packages"),
                       BENCHMARK_ROOT=str(args.root),
                       R_LIBS_USER=str(args.root / "software/R-library"),
                       TEMP=str(temporary_directory), TMP=str(temporary_directory))
    existing_excel = {process.pid for process in psutil.process_iter(["name"])
                      if (process.info["name"] or "").lower() == "excel.exe"}
    command = command_for(library, variant["path"], destination, state_path, args.root)
    fresh_pipeline = variant["tool"] == "SheetShrink" and args.sheetshrink_binary is not None
    if fresh_pipeline:
        original = next(item for item in case["variants"] if item["tool"] == "Original")
        command = [sys.executable, str(Path(__file__).with_name("pipeline.py")), library,
                   original["path"], str(state_path), str(destination), str(args.sheetshrink_binary)]
        state_path.with_name("pipeline.json").unlink(missing_ok=True)
    system_memory = psutil.virtual_memory()
    controller_memory = psutil.Process().memory_info()
    birth, started = time.time(), time.perf_counter()
    with MemoryJob(args.helper_directory, args.limit_mib * 2**20) as job:
        with (directory / "stdout.log").open("wb") as stdout, (directory / "stderr.log").open("wb") as stderr:
            process = job.spawn(command, stdout=stdout, stderr=stderr, env=environment)
            try:
                status, state, resources = monitor(job, process, state_path, args, existing_excel, birth)
            finally:
                if process.poll() is None:
                    job.terminate()
                    process.wait(timeout=10)
    result = {"trial_id": identity, "case_id": case["case_id"], "file": case["file"],
              "tool": variant["tool"], "input_sha256": variant["sha256"], "library": library,
              "repetition": repetition, "status": status, "worker_state": state,
              "wall_seconds": time.perf_counter() - started, "exit_code": process.returncode,
              "limit_bytes": args.limit_mib * 2**20, "reserve_bytes": args.reserve_mib * 2**20, "timeout_seconds": args.timeout,
              "requested_allocation_bytes": None, "resources": resources,
              "host_context": {"available_ram_before_bytes": system_memory.available,
                               "total_ram_bytes": system_memory.total,
                               "controller_private_before_bytes": controller_memory.private,
                               "controller_rss_before_bytes": controller_memory.rss},
              "cleanup_seconds": variant.get("cleanup_seconds"),
              "cleanup_provenance": variant.get("cleanup_provenance"), "command": command}
    if fresh_pipeline:
        pipeline = read_state(state_path.with_name("pipeline.json"))
        result.update(pipeline=pipeline, cleanup_seconds=pipeline.get("cleanup_seconds"),
                      cleanup_provenance="Fresh sequential CLI execution in this bounded trial",
                      original_sha256=original["sha256"], input_sha256=None)
        fresh_output = directory / "fresh-sheetshrink" / Path(original["path"]).name
        if fresh_output.is_file():
            with fresh_output.open("rb") as stream:
                result["input_sha256"] = hashlib.file_digest(stream, "sha256").hexdigest()
            if pipeline.get("cleanup_exit_code") == 0:
                result["repeat_sheetshrink"] = measure_repeat(args, directory, fresh_output, pipeline, environment)
    if destination.is_file():
        if status == "completed":
            result["saved_bytes"] = destination.stat().st_size
            with destination.open("rb") as stream:
                result["saved_sha256"] = hashlib.file_digest(stream, "sha256").hexdigest()
        else:
            result["partial_output_bytes"] = destination.stat().st_size
    elif destination.is_dir():
        result["saved_bytes"] = sum(path.stat().st_size for path in destination.glob("*.csv"))
    receipt = directory / "receipt.json"
    receipt.write_text(json.dumps(result, indent=2), encoding="utf-8")
    result["cleanup_pending"] = []
    for artifact in (destination, temporary_directory, directory / "fresh-sheetshrink", directory / "repeat-sheetshrink"):
        error = remove_artifact(artifact, args.root)
        if error:
            result["cleanup_pending"].append({"path": str(artifact), "error": error})
    receipt.write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


def measure_repeat(args, directory, source, pipeline, environment):
    command = list(pipeline["cleanup_command"])
    output = directory / "repeat-sheetshrink"
    output.mkdir(exist_ok=True)
    command[1:3] = [str(source), str(output)]
    state = directory / "repeat-state.json"
    state.write_text(json.dumps({"phase": "sheetshrink-repeat"}), encoding="utf-8")
    started = time.perf_counter()
    with MemoryJob(args.helper_directory, args.limit_mib * 2**20) as job:
        with (directory / "repeat-stdout.log").open("wb") as stdout, (directory / "repeat-stderr.log").open("wb") as stderr:
            process = job.spawn(command, stdout=stdout, stderr=stderr, env=environment)
            status, _, resources = monitor(job, process, state, args, set(), time.time())
    elapsed = time.perf_counter() - started
    if status == "sheetshrink-repeat":
        status = "completed" if process.returncode == 0 else "failed"
    result = dict(status=status, seconds=elapsed, resources=resources, command=command)
    repeated = output / source.name
    if repeated.is_file():
        with repeated.open("rb") as stream:
            result["output_sha256"] = hashlib.file_digest(stream, "sha256").hexdigest()
        with source.open("rb") as stream:
            result["input_sha256"] = hashlib.file_digest(stream, "sha256").hexdigest()
        result["byte_identical"] = result["output_sha256"] == result["input_sha256"]
        result["input_bytes"] = source.stat().st_size
        result["output_bytes"] = repeated.stat().st_size
    return result


def remove_artifact(path, root):
    resolved = path.resolve()
    if resolved == root.resolve() or not resolved.is_relative_to(root.resolve()):
        raise ValueError("Cleanup target is outside the trial directory")
    for attempt in range(3):
        try:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink(missing_ok=True)
            return None
        except OSError as error:
            if attempt == 2:
                return str(error)
            time.sleep(0.25)


if __name__ == "__main__":
    main()
