import hashlib
import json
import os
import shutil
import time
from collections import Counter
from pathlib import Path

from .common import RunLock, contained, long_path, read_json, rows, sha256, utc_now, write_json
from .excel import convert_file, preflight
from .ingest import download_sources, freeze_sources, inventory_sources
from .reuse import reuse_file


class StopRequested(Exception):
    pass


class Progress:
    def __init__(self, root):
        self.root = root
        self.last = 0
        self.last_print = 0
        self.stage = None

    def check_stop(self):
        if (self.root / "stop.request").exists():
            raise StopRequested()

    def __call__(self, event):
        self.check_stop()
        now = time.monotonic()
        stage = event.get("event", event.get("stage", "working"))
        if now - self.last < 2 and stage == self.stage:
            return
        write_json(self.root / "status.json", {
            "state": "running", "pid": os.getpid(), "updated_utc": utc_now(), **event,
        })
        if stage != self.stage or now - self.last_print > 30:
            message = event.get("message", stage.replace("_", " "))
            archive = event.get("archive", "")
            completed, total = event.get("completed"), event.get("total")
            count = " ({}/{})".format(completed, total) if completed is not None else ""
            print("{}  {}{} {}".format(utc_now(), message, count, archive).rstrip(), flush=True)
            self.last_print = now
        self.last, self.stage = now, stage


def copy_for_review(root, row, result):
    source_path = row.get("extracted_path")
    if source_path:
        source = contained(root, source_path)
        relative = source.relative_to(root / "extracted")
        target = root / "review" / "files" / relative
        long_path(target.parent).mkdir(parents=True, exist_ok=True)
        if not long_path(target).exists() or sha256(target) != row["sha256"]:
            shutil.copyfile(long_path(source), long_path(target))
        review_path = target.relative_to(root).as_posix()
    else:
        review_path = None
    write_json(root / "review" / "issues" / (row["id"] + ".json"), {
        "source": row, "review_path": review_path, "result": result,
    })


def prepare(root, config, retry_review=False, limit=None, code_record=None):
    root = Path(root).resolve()
    root.mkdir(parents=True, exist_ok=True)
    progress = Progress(root)
    with RunLock(root):
        (root / "stop.request").unlink(missing_ok=True)
        started = time.monotonic()
        record = {
            "started_utc": utc_now(), "pid": os.getpid(), "config": config,
            "code": read_json(code_record) if code_record else None,
        }
        run_path = root / "runs" / (time.strftime("%Y%m%dT%H%M%SZ", time.gmtime()) + "-" + str(os.getpid()))
        write_json(run_path / "run.json", record)
        try:
            record["host"] = preflight()
            write_json(run_path / "run.json", record)
            progress({"event": "sources", "message": "Checking the source manifest"})
            manifest = freeze_sources(root)
            if manifest.get("item", manifest.get("identifier", config["source_item"])) != config["source_item"]:
                raise ValueError("The source manifest and experiment configuration disagree.")
            if len(manifest["archives"]) != config["expected_archives"]:
                raise ValueError("Unexpected archive count; review the source manifest.")
            download_sources(root, manifest, on_progress=progress)
            inventory = inventory_sources(root, manifest, on_progress=progress)
            progress({"event": "preparation", "message": "Preparing workbooks from existing results first"})
            summary = convert_inventory(root, inventory, config, progress, retry_review, limit)
            if not summary["limited"]:
                summary.update(inventory_sha256=sha256(inventory),
                               conversion_manifest_sha256=sha256(root / "conversions.jsonl"),
                               source_manifest_sha256=sha256(root / "sources.json"))
            summary.update(elapsed_seconds=time.monotonic() - started, finished_utc=utc_now())
            write_json(root / "summary.json", summary)
            write_json(run_path / "summary.json", summary)
            write_json(root / "status.json", {"state": "prepared" if not summary["limited"] else "pilot_complete", **summary})
            print("Preparation complete. Accepted workbooks: {}. Reused: {}. Source entries for review: {}.".format(
                summary["unique_accepted"], summary["unique_reused"], summary["review_occurrences"]), flush=True)
            return 0
        except StopRequested:
            write_json(root / "status.json", {"state": "stopped", "updated_utc": utc_now()})
            print("Stopped at a checkpoint. Run the same command to resume.", flush=True)
            return 0
        except Exception as exc:
            failure = {"state": "failed", "error": str(exc), "updated_utc": utc_now()}
            write_json(root / "status.json", failure)
            write_json(run_path / "failure.json", failure)
            raise


def convert_inventory(root, inventory, config, progress, retry_review=False, limit=None):
    counts = Counter()
    unique = {}
    attempts = 0
    limited = False
    cache_roots = [root / "reuse"]
    reuse_sources = root / "reuse-sources.json"
    if reuse_sources.exists():
        cache_roots.extend(Path(path).resolve() for path in read_json(reuse_sources)["caches"])
    code_files = {p.name: sha256(p) for p in sorted(Path(__file__).parent.iterdir())
                  if p.is_file() and p.suffix in (".py", ".ps1", ".cs")}
    code_digest = hashlib.sha256(json.dumps(code_files, sort_keys=True).encode()).hexdigest()
    write_json(root / "state" / "code" / (code_digest + ".json"), code_files)
    result_path = root / "conversions.jsonl"
    pending = result_path.with_suffix(".jsonl.partial")
    total = sum(1 for row in rows(inventory) if row["classification"] in ("candidate", "ambiguous"))
    with pending.open("w", encoding="utf-8", newline="\n") as output:
        for row in rows(inventory):
            counts["inventory_entries"] += 1
            if row["classification"] not in ("candidate", "ambiguous"):
                continue
            progress.check_stop()
            digest = row.get("sha256")
            source = contained(root, row["extracted_path"]) if row.get("extracted_path") else None
            result = unique.get(digest) if digest else None
            if result is None:
                state = root / "state" / "conversions" / ((digest or row["id"]) + ".json")
                if state.exists():
                    cached = read_json(state)
                    if cached.get("config") != config:
                        raise ValueError("Conversion settings changed. Use a new data root to keep results comparable.")
                    result = cached["result"]
                    if result["status"] == "accepted":
                        target = contained(root, result["output_path"])
                        if not long_path(target).is_file() or sha256(target) != result["output_sha256"]:
                            raise ValueError("A converted output changed or is missing: " + str(target))
                        for artifact in list(result.get("variants", {}).values()) + ([result["conversion_artifact"]] if result.get("conversion_artifact") else []):
                            target = contained(root, artifact["path"])
                            target.relative_to(root / "variants")
                            if sha256(target) != artifact["sha256"]:
                                raise ValueError("A saved study variant changed: " + artifact["path"])
                    elif retry_review:
                        result = None
                if result is None:
                    if source:
                        if sha256(source) != digest:
                            raise ValueError("An extracted input changed: " + str(source))
                        relative = source.relative_to(root / "extracted")
                        destination = root / "xlsx" / (relative.as_posix() + ".xlsx")
                        attempt = root / "_work" / (digest[:16] + "-" + str(time.time_ns()))
                        attempt.mkdir(parents=True)
                        result = reuse_file(digest, destination, root / "variants", attempt,
                                            cache_roots, relative_path=relative.as_posix() + ".xlsx")
                        if result is not None:
                            result["attempt_path"] = attempt.relative_to(root).as_posix()
                    if result is None and source and row.get("attempt_excel", row["classification"] == "candidate"):
                        if limit is not None and attempts >= limit:
                            limited = True
                            break
                        result = convert_file(long_path(source), long_path(destination), attempt,
                            stage_timeout=config["stage_timeout_seconds"],
                            total_timeout=config["file_timeout_seconds"])
                        attempts += 1
                        result["attempt_path"] = attempt.relative_to(root).as_posix()
                    if result is None:
                        result = {"status": "review", "reason": "format_review", "evidence": row.get("evidence", row.get("kind"))}
                    if result["status"] == "accepted":
                        result["output_path"] = destination.relative_to(root).as_posix()
                        result["output_sha256"] = sha256(destination)
                        for variant in result.get("variants", {}).values():
                            variant["path"] = Path(variant["path"]).relative_to(root).as_posix()
                        if result.get("conversion_artifact"):
                            converted = result["conversion_artifact"]
                            converted["path"] = Path(converted["path"]).relative_to(root).as_posix()
                    result["converter_code_sha256"] = code_digest
                    write_json(state, {"source_sha256": digest, "config": config, "result": result})
                if digest:
                    unique[digest] = result
            counts["candidate_occurrences"] += 1
            counts[result["status"] + "_occurrences"] += 1
            if result.get("method") == "reuse_study_preparation":
                counts["reused_occurrences"] += 1
            if result["status"] != "accepted":
                copy_for_review(root, row, result)
            else:
                # Retain old attempt evidence; current issues list only unresolved occurrences.
                (root / "review" / "issues" / (row["id"] + ".json")).unlink(missing_ok=True)
            output.write(json.dumps({"source": row, "result": result}, ensure_ascii=True) + "\n")
            output.flush()
            progress({"event": "preparation", "completed": counts["candidate_occurrences"], "total": total,
                      "accepted": counts["accepted_occurrences"], "reused": counts["reused_occurrences"],
                      "review": counts["review_occurrences"]})
    if not limited:
        os.replace(pending, result_path)
    return {
        **dict(counts), "review_occurrences": counts["review_occurrences"],
        "unique_accepted": sum(x["status"] == "accepted" for x in unique.values()),
        "unique_review": sum(x["status"] != "accepted" for x in unique.values()),
        "unique_reused": sum(x.get("method") == "reuse_study_preparation" and x["status"] == "accepted" for x in unique.values()),
        "conversion_attempts": attempts, "limited": limited,
        "scope": "Spreadsheet candidates in the frozen EDRM v2 XML archives; unresolved formats remain listed for review.",
    }
