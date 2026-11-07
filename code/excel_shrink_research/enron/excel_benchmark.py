import json
import time
from pathlib import Path

from .common import RunLock, long_path, sha256, utc_now, write_json
from .excel import convert_file
from .release import accepted_outputs


def benchmark_excel(root, destination, stage_timeout=120, total_timeout=420):
    """Measure Excel SaveAs on the same frozen XLSX inputs used by Rust."""
    root, destination = Path(root).resolve(), Path(destination).resolve()
    with RunLock(root):
        if destination.exists():
            raise ValueError("Use a new Excel benchmark directory.")
        inputs = accepted_outputs(root)
        destination.mkdir(parents=True)
        write_json(destination / "run.json", {
            "started_utc": utc_now(), "input_manifest_sha256": sha256(root / "conversions.jsonl"),
            "input_files": len(inputs), "workers": 1,
            "policy": "Fresh Excel instance per input; manual calculation, no link updates. Reopen validation is timed separately.",
        })
        counters = {"accepted": 0, "review": 0}
        times = {"open_seconds": 0, "save_seconds": 0, "convert_seconds": 0, "verify_seconds": 0}
        started = time.perf_counter()
        with (destination / "results.jsonl").open("w", encoding="utf-8") as stream:
            for index, (relative, digest) in enumerate(sorted(inputs.items())):
                attempt = destination / "_work" / ("{:06d}".format(index))
                result = convert_file(long_path(root / relative), long_path(destination / relative), attempt,
                                      stage_timeout=stage_timeout, total_timeout=total_timeout, force_save=True)
                counters[result["status"]] += 1
                convert = result.get("phases", {}).get("convert", {})
                verify = result.get("phases", {}).get("verify", {})
                times["open_seconds"] += convert.get("timings", {}).get("open", 0)
                times["save_seconds"] += convert.get("timings", {}).get("save", 0)
                times["convert_seconds"] += convert.get("elapsed_seconds", 0)
                times["verify_seconds"] += verify.get("elapsed_seconds", 0)
                stream.write(json.dumps({"input_path": relative, "input_sha256": digest, "result": result}) + "\n")
                stream.flush()
                if (index + 1) % 100 == 0:
                    print("Excel SaveAs: {}/{} files".format(index + 1, len(inputs)), flush=True)
        write_json(destination / "summary.json", {
            **counters, **times, "wall_seconds": time.perf_counter() - started,
            "timing_note": "Wall time includes worker startup, staging, and reopen validation; use the separate phase totals when reporting COM Open and SaveAs.",
        })
        if counters["review"]:
            raise RuntimeError("Some Excel SaveAs inputs need review; see results.jsonl.")
        print("Excel benchmark complete: " + str(destination))
