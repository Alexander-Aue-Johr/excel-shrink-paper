"""Clean one original and immediately process its fresh output."""
import json
import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import subprocess
import sys
import time

from run_bounded import command_for


class ProcessMemory(ctypes.Structure):
    _fields_ = [("cb", wintypes.DWORD), ("faults", wintypes.DWORD)] + [
        (name, ctypes.c_size_t) for name in ("peak_working_set", "working_set", "peak_paged_pool",
        "paged_pool", "peak_nonpaged_pool", "nonpaged_pool", "pagefile", "peak_pagefile", "private")]


def process_memory(process):
    counters = ProcessMemory()
    counters.cb = ctypes.sizeof(counters)
    query = ctypes.WinDLL("psapi", use_last_error=True).GetProcessMemoryInfo
    query.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessMemory), wintypes.DWORD]
    if not query(int(process._handle), ctypes.byref(counters), counters.cb):
        return {"error": str(ctypes.WinError(ctypes.get_last_error()))}
    return {"peak_working_set_bytes": counters.peak_working_set,
            "peak_committed_bytes": counters.peak_pagefile, "private_at_exit_bytes": counters.private}


def main():
    library, source, state, destination, binary = sys.argv[1:]
    source, state, destination = Path(source), Path(state), Path(destination)
    output = destination.parent / "fresh-sheetshrink"
    output.mkdir(exist_ok=True)
    state.write_text(json.dumps({"phase": "sheetshrink", "monotonic": time.perf_counter()}), encoding="utf-8")
    workers = str(os.cpu_count() or 1)
    command = [binary, str(source), str(output), "--force-overwrite",
               "--rayon-threads", workers, "--writer-workers", workers, "--max-active-workbooks", "1",
               "--console-log-level", "error", "--file-log-level", "error"]
    started = time.perf_counter()
    cleaned = subprocess.Popen(command)
    cleaned.wait()
    metadata = {"cleanup_seconds": time.perf_counter() - started,
                "cleanup_exit_code": cleaned.returncode, "cleanup_command": command,
                "original_path": str(source), "fresh_output_path": str(output / source.name)}
    metadata["cleanup_resources"] = process_memory(cleaned)
    metadata_path = state.with_name("pipeline.json")
    metadata_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    if cleaned.returncode != 0 or not (output / source.name).is_file():
        state.write_text(json.dumps({"phase": "failed", "failed_phase": "sheetshrink",
                                    "error": "SheetShrink did not produce a successful output"}), encoding="utf-8")
        return 1
    command = command_for(library, str(output / source.name), destination, state,
                          Path(os.environ["BENCHMARK_ROOT"]))
    processed = subprocess.run(command, check=False)
    metadata["pipeline_seconds"] = time.perf_counter() - started
    metadata["library_exit_code"] = processed.returncode
    metadata_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return processed.returncode


if __name__ == "__main__":
    sys.exit(main())
