"""Run the remaining Enron and EUSES preparation steps without overlapping Excel jobs."""

import argparse
import ctypes
from datetime import datetime, timezone
import os
from pathlib import Path
import re
import subprocess
import time

from .enron.common import RunLock, read_json, sha256, utc_now, write_json


def started_ms(value):
    normalized = re.sub(r"\.(\d+)(?=[+-])", lambda match: "." + match[1][:6].ljust(6, "0"),
                        value.replace("Z", "+00:00"))
    instant = datetime.fromisoformat(normalized)
    delta = instant - datetime(1970, 1, 1, tzinfo=timezone.utc)
    return delta.days * 86400000 + delta.seconds * 1000 + delta.microseconds // 1000


class OwnedProcess:
    """Keep a Windows process handle after checking its creation time."""

    def __init__(self, pid, expected_start):
        from ctypes import wintypes
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        self.kernel.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
        self.kernel.OpenProcess.restype = wintypes.HANDLE
        self.kernel.GetProcessTimes.argtypes = (wintypes.HANDLE,) + (ctypes.POINTER(wintypes.FILETIME),) * 4
        self.kernel.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
        self.kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
        self.handle = self.kernel.OpenProcess(0x1000 | 0x100000, False, pid)
        if not self.handle:
            error = ctypes.get_last_error()
            if error in (87, 1168):
                return
            raise OSError(error, "Could not inspect the preparation runner.")
        times = [wintypes.FILETIME() for _ in range(4)]
        if not self.kernel.GetProcessTimes(self.handle, *(ctypes.byref(value) for value in times)):
            self.close()
            raise OSError(ctypes.get_last_error(), "Could not verify the runner creation time.")
        ticks = (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime
        if (ticks - 116444736000000000) // 10000 != expected_start:
            self.close()

    def ended(self):
        if not self.handle:
            return True
        result = self.kernel.WaitForSingleObject(self.handle, 0)
        if result not in (0, 258):
            raise OSError(ctypes.get_last_error(), "Could not inspect the runner state.")
        return result == 0

    def close(self):
        if self.handle:
            self.kernel.CloseHandle(self.handle)
            self.handle = None


def output_idle(output):
    path = Path(output) / ".run.lock"
    if not path.exists():
        return True
    with path.open("r+b") as stream:
        if os.name == "nt":
            import msvcrt
            try:
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                return False
            msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            try:
                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                return False
            fcntl.flock(stream.fileno(), fcntl.LOCK_UN)
    return True


def completion(output, receipt, corpus, partition):
    """A retained skip can produce exit 1; require the runner's final snapshots."""
    room = Path(output) / "control-room"
    queue, progress = read_json(room / "queue.json"), read_json(room / "progress.json")
    if queue.get("runner_pid") != receipt["Pid"] or queue.get("runner_started_unix_ms") != started_ms(receipt["StartedUtc"]):
        raise RuntimeError("The output belongs to another preparation runner.")
    if queue.get("running") is not False or queue.get("phase") != "finished":
        raise RuntimeError("The preparation runner ended before its final checkpoint.")
    if any(group.get("active") for group in queue.get("groups", [])):
        raise RuntimeError("The final checkpoint still lists active workbooks.")
    if progress.get("run_id") != queue.get("run_id"):
        raise RuntimeError("The preparation progress belongs to another run.")
    scope = progress["scope"]
    expected_part = None if partition == 0 else {"part": partition, "parts": 4}
    if scope.get("corpora") != [corpus] or scope.get("size_part") != expected_part:
        raise RuntimeError("The preparation checkpoint has a different selection.")
    counts = progress["counts"]
    if counts.get("in_progress") != 0 or counts.get("preparation_remaining") != 0:
        raise RuntimeError("Some selected workbooks have not reached a preparation outcome.")
    if not output_idle(output):
        raise RuntimeError("The preparation runner has not released its output lock.")
    return {"run_id": queue["run_id"], "counts": counts, "finished_utc": utc_now()}


def wait_run(receipt, corpus, partition, report, stop):
    owner = OwnedProcess(receipt["Pid"], started_ms(receipt["StartedUtc"]))
    try:
        while not owner.ended():
            if stop():
                raise RuntimeError("The queue was stopped; the current runner was left running.")
            progress_path = Path(receipt["OutputRoot"]) / "control-room/progress.json"
            try:
                progress = read_json(progress_path)
                report({"runner_pid": receipt["Pid"], "progress": progress.get("counts")})
            except (OSError, ValueError):
                pass
            time.sleep(15)
        return completion(receipt["OutputRoot"], receipt, corpus, partition)
    finally:
        owner.close()


def check_existing_preparations(output, corpus_id, fingerprint, profiles, builds, executable_sha256):
    """Block compatibility misses rather than silently converting valid work again."""
    checked = 0
    for path in (Path(output) / "records").glob("*.json"):
        record = read_json(path)
        if record.get("corpus_id", record.get("identity", "").split(":", 1)[0]) != corpus_id:
            continue
        if not record.get("prepared") or (record.get("complete") and record.get("status") != "skipped"):
            continue
        if record.get("source_fingerprint") != fingerprint:
            raise RuntimeError("A saved preparation uses a different source catalog: " + path.name)
        if record.get("prepared_by") not in builds and record.get("preparation_accepted_by") != executable_sha256:
            raise RuntimeError("A saved preparation has an unreviewed producer: " + path.name)
        rows = record.get("results", [])
        actual = {row.get("profile"): row.get("config_hash") for row in rows}
        if len(rows) != len(profiles) or len(actual) != len(profiles) or any(actual.get(name) != value for name, value in profiles.items()):
            raise RuntimeError("A saved preparation has different profile settings: " + path.name)
        if not (Path(output) / "prepared" / (path.stem + ".zip")).is_file():
            raise RuntimeError("A saved preparation is missing its bundle: " + path.name)
        checked += 1
    return checked


def launch_job(config, root, name, corpus, output, port, builds, *, archive=None):
    if (root / "stop.request").exists():
        raise RuntimeError("The queue was stopped before starting step " + name)
    job_root = root / "jobs" / name
    job_root.mkdir(parents=True, exist_ok=True)
    receipt_path = job_root / "receipt.json"
    if receipt_path.is_file():
        return read_json(receipt_path)
    if sha256(config["executable"]) != config["executable_sha256"]:
        raise RuntimeError("The pinned preparation executable has changed.")
    job = {
        "rust_repo": config["rust_repo"], "base_configuration": config["base_configuration"],
        "configuration_path": str(job_root / "config.json"), "run_directory": str(job_root / "run"),
        "executable": config["executable"], "executable_sha256": config["executable_sha256"],
        "receipt_path": str(receipt_path), "corpus": corpus, "output_root": str(output),
        "port": port, "approved_builds": sorted(builds), "retry_skipped": name == "B",
    }
    if archive:
        job.update(archive=archive["path"], corpus_id=archive["corpus_id"], total_workbooks=archive["workbooks"])
    job_path = job_root / "job.json"
    write_json(job_path, job)
    script = Path(__file__).resolve().parents[2] / "scripts/run_preparation_job.ps1"
    with (job_root / "launch.stdout.log").open("w", encoding="utf-8") as stdout, (job_root / "launch.stderr.log").open("w", encoding="utf-8") as stderr:
        result = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                                 str(script), "-JobFile", str(job_path)], stdout=stdout, stderr=stderr, timeout=300)
    if result.returncode or not receipt_path.is_file():
        raise RuntimeError("Could not start preparation step " + name + "; see " + str(job_root / "launch.stderr.log"))
    return read_json(receipt_path)


def wait_download(root, report, stop):
    receipt_path = Path(root) / "download-process.json"
    if not receipt_path.is_file():
        return
    receipt = read_json(receipt_path)
    owner = OwnedProcess(receipt["pid"], started_ms(receipt["started_utc"]))
    try:
        while not owner.ended():
            if stop():
                raise RuntimeError("The queue was stopped while the EUSES download continued.")
            path = Path(root) / "download-status.json"
            report({"message": "Waiting for the parallel EUSES download.",
                    "download": read_json(path) if path.is_file() else {}})
            time.sleep(15)
    finally:
        owner.close()
    if stop():
        raise RuntimeError("The queue was stopped before EUSES preparation.")


def run_queue(config_path):
    config = read_json(config_path)
    root = Path(config["queue_root"])
    root.mkdir(parents=True, exist_ok=True)
    checkpoint = root / "checkpoints.json"
    done = read_json(checkpoint) if checkpoint.exists() else {}
    stage = "A"

    def stop():
        return (root / "stop.request").exists()

    def report(extra=None):
        write_json(root / "status.json", {"state": "running", "stage": stage,
                   "pid": os.getpid(), "updated_utc": utc_now(), **(extra or {})})

    def save(name, value):
        done[name] = value
        write_json(checkpoint, done)

    with RunLock(root):
        try:
            if "A" not in done:
                report({"message": "Waiting for Enron partition 1/4 to finish."})
                save("A", wait_run(config["initial_run"], "Enron", 1, report, stop))
            stage = "copy"
            report({"message": "Waiting for the verified local transfer copy."})
            while True:
                if stop():
                    raise RuntimeError("The queue was stopped before import.")
                status_path = Path(config["transfer_status"])
                status = read_json(status_path) if status_path.is_file() else {}
                if status.get("state") == "staged":
                    break
                if status.get("state") == "failed":
                    raise RuntimeError("The transfer copy failed: " + status.get("error", "see transfer status"))
                report({"message": "Waiting for the transfer copy.", "copy": status})
                time.sleep(15)
            if "import" not in done:
                from .preparation_transfer import merge_transfer
                stage = "import"
                report({"message": "Importing the verified Enron partition 4/4."})
                merged = merge_transfer(config["staged_transfer"], config["enron_output"],
                                        expected_source=config["source"], expected_profiles=config["profiles"],
                                        on_progress=lambda value: report({"transfer": value}))
                if merged.get("conflicts"):
                    raise RuntimeError("The preparation transfer has conflicts; review its import report.")
                save("import", merged)
            if "B" not in done:
                stage = "B"
                builds = set(config["approved_builds"] + [config["executable_sha256"]])
                builds.update(read_json(Path(config["staged_transfer"]) / "manifest.json")["preparation_builds"])
                report({"message": "Checking saved preparations before retrying failed Enron files."})
                checked = check_existing_preparations(config["enron_output"], config["source"]["id"], config["source"]["source_fingerprint"],
                                                       config["profiles"], builds, config["executable_sha256"])
                receipt = launch_job(config, root, "B", "Enron", config["enron_output"], 8771, builds)
                report({"message": "Retrying failed Enron preparations.", "reuse_checked": checked})
                save("B", wait_run(receipt, "Enron", 0, report, stop))
            if "C" not in done:
                from .euses import download_euses, prepare_runner_archive
                stage = "C"
                report({"message": "Checking the parallel EUSES download."})
                wait_download(config["euses_root"], report, stop)
                download_euses(config["euses_root"], on_progress=lambda value: report({"download": value}))
                save("C", prepare_runner_archive(config["euses_root"]))
            if "D" not in done:
                stage = "D"
                report({"message": "Preparing EUSES workbooks."})
                receipt = launch_job(config, root, "D", "EUSES", config["euses_output"], 8772,
                                     {config["executable_sha256"]}, archive=done["C"])
                save("D", wait_run(receipt, "EUSES", 0, report, stop))
            write_json(root / "status.json", {"state": "complete", "updated_utc": utc_now(), "steps": done})
            print("Enron retries and EUSES preparation have finished. Review the recorded exceptions.", flush=True)
            return 0
        except Exception as error:
            write_json(root / "status.json", {"state": "failed", "stage": stage,
                       "error": str(error), "updated_utc": utc_now()})
            print("Preparation queue stopped: " + str(error), flush=True)
            return 1


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("config", type=Path)
    return run_queue(parser.parse_args(argv).config)


if __name__ == "__main__":
    raise SystemExit(main())
