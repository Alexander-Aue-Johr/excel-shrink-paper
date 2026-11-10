"""Observe a launched command or an explicitly selected GUI process and its descendants."""
import argparse
import json
from pathlib import Path
import subprocess
import sys
import time
from zipfile import ZipFile, BadZipFile

import psutil


def parse_arguments():
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--pid", type=int)
    mode.add_argument("--command-file", type=Path)
    parser.add_argument("--engine-source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--events", type=Path)
    parser.add_argument("--target", type=Path)
    parser.add_argument("--timeout", type=float, default=900)
    arguments = parser.parse_args()
    if arguments.pid and (arguments.events is None or arguments.target is None):
        parser.error("GUI observations require an events file and a target XLSX")
    if arguments.timeout <= 0:
        parser.error("Timeout must be positive")
    return arguments


def read_events(path):
    if path is None or not path.exists():
        return []
    complete_lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    return [json.loads(line) for line in complete_lines if line.endswith("\n")]


def target_stamp(path):
    if path is None:
        return None
    try:
        stat = path.stat()
    except FileNotFoundError:
        return None
    return stat.st_size, stat.st_mtime_ns


def readable_package(path):
    try:
        with ZipFile(path) as package:
            return package.testzip() is None
    except (OSError, BadZipFile):
        return False


def sample_tree(root, known, probe_type):
    try:
        descendants = root.children(recursive=True)
    except psutil.NoSuchProcess:
        descendants = []
    for process in [root, *descendants]:
        try:
            identity = process.pid, process.create_time()
            if identity not in known:
                known[identity] = {"process": process, "probe": probe_type(process.pid), "name": process.name()}
        except (psutil.Error, OSError):
            continue
    working_set = private = 0
    for entry in known.values():
        entry["probe"].sample()
        try:
            if not entry["process"].is_running():
                continue
            memory = entry["process"].memory_info()
            working_set += memory.rss
            private += memory.private
        except psutil.Error:
            pass
    return {"at": time.time(), "tree_working_set_bytes": working_set, "tree_private_bytes": private}


def observe(arguments, root, child, probe_type):
    known = {}
    initial = target_stamp(arguments.target)
    previous_target = None
    stable_samples = 0
    peak_working_set = peak_private = 0
    started = time.time()
    status = "running"
    with (arguments.output / "samples.jsonl").open("w") as journal:
        while status == "running":
            sample = sample_tree(root, known, probe_type)
            peak_working_set = max(peak_working_set, sample["tree_working_set_bytes"])
            peak_private = max(peak_private, sample["tree_private_bytes"])
            journal.write(json.dumps(sample) + "\n")
            events = read_events(arguments.events)
            if child is not None and child.poll() is not None:
                status = "completed" if child.returncode == 0 else "command_failed"
            if child is None and any(event["phase"] == "optimize_start" for event in events):
                current = target_stamp(arguments.target)
                stable_samples = stable_samples + 1 if current == previous_target and current != initial else 0
                previous_target = current
                if current is not None and current != initial and stable_samples >= 2 and readable_package(arguments.target):
                    status = "output_package_verified"
            if time.time() - started > arguments.timeout:
                status = "timed_out"
                if child is not None:
                    for entry in reversed(list(known.values())):
                        try:
                            entry["process"].kill()
                        except psutil.Error:
                            pass
                    child.wait(timeout=15)
            if status == "running":
                time.sleep(0.1)
    finished = time.time()
    processes = []
    for (pid, birth), entry in known.items():
        try:
            metrics = entry["probe"].finish()
            metrics["sampling_interval_seconds"] = 0.1
            metrics["scope"] = "one observed process in the selected subtree"
            processes.append({"pid": pid, "created_at": birth, "name": entry["name"], **metrics})
        finally:
            entry["probe"].close()
    return {"status": status, "started_at": started, "finished_at": finished,
            "observation_seconds": finished - started, "events": read_events(arguments.events),
            "sampling_interval_seconds": 0.1, "sampled_tree_peak_working_set_bytes": peak_working_set,
            "sampled_tree_peak_private_bytes": peak_private, "processes": processes,
            "cpu_user_seconds": sum(process["cpu_user_seconds"] for process in processes),
            "cpu_kernel_seconds": sum(process["cpu_kernel_seconds"] for process in processes),
            "exit_code": None if child is None else child.returncode,
            "limitations": ["Tree peaks are sampled. Short-lived undiscovered descendants may be missed.",
                            "Summed working sets can count shared resident pages more than once.",
                            "Per-process OS peaks cover that process's lifetime, including GUI setup.",
                            "GUI completion is the first stable CRC-valid output package observed after optimize_start."]}


def main():
    arguments = parse_arguments()
    arguments.output.mkdir(parents=True, exist_ok=False)
    support = arguments.engine_source / "docs/visual-validation/experiments"
    sys.path.insert(0, str(support))
    from process_metrics import WindowsProbe

    child = None
    started = time.time()
    if arguments.command_file:
        command = json.loads(arguments.command_file.read_text(encoding="utf-8-sig"))
        with (arguments.output / "stdout.log").open("wb") as stdout, (arguments.output / "stderr.log").open("wb") as stderr:
            child = subprocess.Popen(command, stdout=stdout, stderr=stderr)
    root = psutil.Process(arguments.pid or child.pid)
    root_birth = root.create_time()
    if arguments.pid and started - root_birth > 120:
        raise ValueError("A GUI resource trial requires a fresh process launched within 120 seconds")
    result = observe(arguments, root, child, WindowsProbe)
    result.update(root_pid=root.pid, root_created_at=root_birth, psutil_version=psutil.__version__,
                  command_started_at=started, command=None if child is None else command)
    result["command_wall_seconds"] = result["finished_at"] - started
    (arguments.output / "resources.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps({"status": result["status"], "seconds": result["observation_seconds"],
                      "peak_tree_working_set_bytes": result["sampled_tree_peak_working_set_bytes"]}), flush=True)


if __name__ == "__main__":
    main()
