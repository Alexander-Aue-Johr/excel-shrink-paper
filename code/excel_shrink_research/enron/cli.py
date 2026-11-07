import argparse
import json
import sys
from pathlib import Path

from .common import RunLock, read_json, utc_now, write_json


def main(argv=None):
    parser = argparse.ArgumentParser(description="Prepare the Enron spreadsheet corpus and run corpus experiments.")
    parser.add_argument("--root", type=Path, default=Path("F:/corpora/enron-v2"), help="Data directory (outside the Git checkout).")
    commands = parser.add_subparsers(dest="command", required=True)
    prepare = commands.add_parser("prepare", help="Download, inventory, and convert; resume existing checkpoints.")
    prepare.add_argument("--config", type=Path, default=Path(__file__).resolve().parents[3] / "experiments/enron/config.json")
    prepare.add_argument("--code-record", type=Path)
    prepare.add_argument("--retry-review", action="store_true", help="Retry recorded conversion failures with the same settings.")
    prepare.add_argument("--limit", type=int, help="Limit new conversion attempts for a pilot.")
    compare = commands.add_parser("compare", help="Compare original spreadsheet contents with an older corpus archive.")
    compare.add_argument("old_archive", type=Path)
    compare.add_argument("--destination", type=Path, help="Report directory (default: ROOT/comparison).")
    commands.add_parser("status", help="Print the latest progress and whether the process holds its lock.")
    commands.add_parser("stop", help="Request a stop at the next checkpoint.")
    package = commands.add_parser("package", help="Build a local XLSX release archive with source mappings and exceptions.")
    package.add_argument("destination", type=Path)
    benchmark = commands.add_parser("benchmark", help="Run all three shrink profiles with a prebuilt executable.")
    benchmark.add_argument("executable", type=Path)
    benchmark.add_argument("destination", type=Path)
    benchmark.add_argument("--repeats", type=int, default=3)
    excel_benchmark = commands.add_parser("excel-benchmark", help="Measure Excel opening and saving the frozen XLSX inputs.")
    excel_benchmark.add_argument("destination", type=Path)
    export = commands.add_parser("export-study", help="Preserve prepared study workbooks for reuse on this or another machine.")
    export.add_argument("study_root", type=Path)
    export.add_argument("destination", type=Path)
    export.add_argument("--watch", action="store_true", help="Keep exporting newly prepared bundles until the stop file exists.")
    export.add_argument("--stop-file", type=Path)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        if args.command == "prepare":
            from .workflow import prepare
            config = read_json(args.config)
            if config["conversion_workers"] != 1:
                raise ValueError("This first conversion run uses one Excel worker. Measure a pilot before increasing concurrency.")
            if min(config["stage_timeout_seconds"], config["file_timeout_seconds"]) <= 0:
                raise ValueError("Timeouts must be positive.")
            if args.limit is not None and args.limit < 1:
                raise ValueError("The pilot limit must be positive.")
            return prepare(root, config, args.retry_review, args.limit, args.code_record)
        if args.command == "compare":
            from .compare import compare_corpora, index_old_archive
            destination = (args.destination or root / "comparison").resolve()
            def report_progress(event):
                if event.get("event") == "legacy_index":
                    write_json(destination / "status.json", {
                        "state": "running", "updated_utc": utc_now(), **event})
                elif event.get("event") == "comparison" and "new_only_unique_sha256" in event:
                    print("Compared {}/{} archives: {}".format(
                        event["completed"], event["total"], event["archive"]), flush=True)
            with RunLock(destination / ".command"):
                write_json(destination / "status.json", {
                    "state": "running", "event": "legacy_index", "updated_utc": utc_now()})
                print("Checking the existing corpus index.", flush=True)
                try:
                    index = index_old_archive(args.old_archive.resolve(), destination, on_progress=report_progress)
                    summary = compare_corpora(root, index, destination, on_progress=report_progress)
                except Exception as error:
                    write_json(destination / "status.json", {
                        "state": "failed", "error": str(error), "updated_utc": utc_now()})
                    raise
            print(json.dumps(summary, indent=2), flush=True)
            return 0
        if args.command == "status":
            if not root.exists():
                print("No run has started in " + str(root))
                return 0
            status = read_json(root / "status.json") if (root / "status.json").exists() else {"state": "not_started"}
            try:
                with RunLock(root):
                    status["process_active"] = False
            except RuntimeError:
                status["process_active"] = True
            print(json.dumps(status, indent=2))
        elif args.command == "stop":
            if not root.exists():
                raise ValueError("No run has started in " + str(root))
            (root / "stop.request").touch()
            print("Stop requested. The current Excel call remains bounded by its timeout.")
        elif args.command == "package":
            from .release import package
            package(root, args.destination)
        elif args.command == "benchmark":
            from .benchmark import benchmark
            if args.repeats < 1:
                raise ValueError("At least one repetition is required.")
            benchmark(root, args.executable, args.destination, args.repeats)
        elif args.command == "excel-benchmark":
            from .excel_benchmark import benchmark_excel
            benchmark_excel(root, args.destination)
        elif args.command == "export-study":
            from .study_export import preserve_once, watch_study
            if args.watch:
                if args.stop_file is None:
                    raise ValueError("--watch requires --stop-file.")
                watch_study(args.study_root, args.destination, args.stop_file)
            else:
                result = preserve_once(args.study_root, args.destination)
                print(json.dumps(result, indent=2))
                if result["errors"] or result["snapshot"]["errors"]:
                    return 1
        return 0
    except (OSError, ValueError, RuntimeError) as exc:
        print("Error: " + str(exc), file=sys.stderr, flush=True)
        return 1
