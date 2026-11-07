"""Keep completed study bundles for reuse, independently of visual comparisons."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import time
from typing import Callable
import uuid
import zipfile


_HASH = re.compile(r"^[0-9a-f]{64}$")
_REQUIRED = (
    "original.xlsx",
    "default.xlsx",
    "maximum-cleaning.xlsx",
    "too-aggressive-cleaning.xlsx",
)


def _io(path: Path) -> Path:
    value = str(path.absolute())
    if os.name == "nt" and not value.startswith("\\\\?\\"):
        value = "\\\\?\\UNC\\" + value[2:] if value.startswith("\\\\") else "\\\\?\\" + value
    return Path(value)


def _stat(path: Path) -> dict:
    value = _io(path).stat()
    return {"size": value.st_size, "mtime_ns": value.st_mtime_ns}


def _read_json(path: Path) -> dict:
    with _io(path).open(encoding="utf-8-sig") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise ValueError("Expected a JSON object")
    return value


def _write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    pending = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with pending.open("w", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, ensure_ascii=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(pending, path)
    finally:
        pending.unlink(missing_ok=True)


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _snapshot_json(path: Path, value: dict) -> None:
    """Write snapshot metadata without a disk flush for every small record."""
    pending = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        _io(pending).write_text(json.dumps(value, ensure_ascii=False) + "\n", encoding="utf-8")
        os.replace(_io(pending), _io(path))
    finally:
        _io(pending).unlink(missing_ok=True)


def snapshot_study(study_root: str | Path, destination: str | Path) -> dict:
    """Retain stable bundles quickly, before doing any hashing or ZIP inspection.

    Prepared bundles are published by replacement. A hard link therefore keeps
    the completed version even when the study later removes its own filename.
    Filesystems without hard links use a copy instead.
    """
    study_root = Path(study_root).resolve()
    destination = Path(destination).resolve()
    target_bundles = destination / "prepared"
    target_records = destination / "records"
    target_bundles.mkdir(parents=True, exist_ok=True)
    target_records.mkdir(parents=True, exist_ok=True)
    summary = {
        "snapshot_root": str(destination), "captured": 0, "existing": 0,
        "pending": 0, "errors": [],
    }
    snapshot_manifest = destination / "snapshot.json"
    if not snapshot_manifest.exists():
        _snapshot_json(snapshot_manifest, {"schema_version": 1, "study_root": str(study_root)})
    prepared = study_root / "prepared"
    if not prepared.is_dir():
        return summary
    for source in prepared.glob("*.zip"):
        pending = None
        try:
            before = _stat(source)
            version = f"{source.stem}-{before['size']}-{before['mtime_ns']}"
            target_bundle = target_bundles / (version + ".zip")
            target_record = target_records / (version + ".json")
            if target_bundle.is_file() and target_record.is_file():
                summary["existing"] += 1
                continue
            record_path = study_root / "records" / (source.stem + ".json")
            record_stat = _stat(record_path)
            record = _read_json(record_path)
            source_sha = record.get("source_sha256", "")
            if (
                not record.get("prepared")
                or not isinstance(source_sha, str)
                or not _HASH.fullmatch(source_sha)
                or _stat(record_path) != record_stat
            ):
                summary["pending"] += 1
                continue
            pending = target_bundles / (version + "." + uuid.uuid4().hex + ".tmp")
            try:
                os.link(source, pending)
            except OSError:
                shutil.copy2(source, pending)
            if _stat(source) != before or _stat(pending) != before:
                summary["pending"] += 1
                continue
            # Replacements during capture must not pair a new bundle with an
            # older preparation record. Comparison-only record changes are safe.
            if _stat(record_path) != record_stat:
                latest = _read_json(record_path)
                fields = ("source_sha256", "baseline_sha256", "prepared_by")
                if not latest.get("prepared") or any(
                    latest.get(field) != record.get(field) for field in fields
                ):
                    summary["pending"] += 1
                    continue
            os.replace(pending, target_bundle)
            pending = None
            _snapshot_json(target_record, record)
            summary["captured"] += 1
        except (FileNotFoundError, PermissionError, json.JSONDecodeError):
            summary["pending"] += 1
        except (OSError, ValueError) as error:
            if len(summary["errors"]) < 20:
                summary["errors"].append({"bundle": source.name, "error": str(error)})
        finally:
            if pending is not None:
                pending.unlink(missing_ok=True)
    for category in ("runs", "machines"):
        metadata = study_root / "metadata" / category
        if not metadata.is_dir():
            continue
        output = destination / "metadata" / category
        output.mkdir(parents=True, exist_ok=True)
        for source in metadata.glob("*.json"):
            target = output / source.name
            if not target.exists():
                try:
                    _snapshot_json(target, _read_json(source))
                except (OSError, ValueError):
                    continue
    return summary


def _inspect_bundle(path: Path, record: dict) -> list[dict]:
    files = []
    with zipfile.ZipFile(path) as bundle:
        names = bundle.namelist()
        for name in _REQUIRED:
            if names.count(name) != 1:
                raise ValueError(f"Bundle must contain exactly one {name}")
        for info in bundle.infolist():
            if info.is_dir():
                continue
            digest = hashlib.sha256()
            with bundle.open(info) as stream:
                for block in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(block)
            if info.filename in (*_REQUIRED, "converted.xlsx", "preparation.json"):
                files.append({
                    "name": info.filename,
                    "sha256": digest.hexdigest(),
                    "bytes": info.file_size,
                    "size": info.file_size,
                })
    original = next(item for item in files if item["name"] == "original.xlsx")
    expected = record.get("baseline_sha256")
    if expected and original["sha256"] != expected:
        raise ValueError("Prepared baseline does not match the study record")
    return files


def _provenance(study_root: Path, record: dict) -> dict:
    result = {"study_root": str(study_root), "record": record.get("provenance", {})}
    try:
        result["study_root"] = _read_json(study_root / "snapshot.json")["study_root"]
    except (OSError, ValueError, KeyError):
        pass
    references = set()
    preparation = record.get("provenance", {}).get("preparation", {})
    for value in (preparation, record.get("last_execution", {})):
        if not isinstance(value, dict):
            continue
        if value.get("run_id"):
            references.add(f"metadata/runs/{value['run_id']}.json")
        if value.get("machine_manifest"):
            references.add(value["machine_manifest"])
    manifests = {}
    for relative in sorted(references):
        candidate = (study_root / relative).resolve()
        if not candidate.is_relative_to(study_root):
            continue
        try:
            manifests[relative] = _read_json(candidate)
        except (OSError, ValueError):
            continue
    result["manifests"] = manifests
    return result


def _cached(receipt: dict, destination: Path, bundle_stat: dict, source_sha: str) -> bool:
    if receipt.get("bundle_stat") != bundle_stat or receipt.get("source_sha256") != source_sha:
        return False
    for field, stat_field in (("bundle_path", "cached_bundle_stat"), ("entry_path", "entry_stat")):
        relative = receipt.get(field)
        if not isinstance(relative, str):
            return False
        candidate = (destination / relative).resolve()
        if not candidate.is_relative_to(destination):
            return False
        try:
            if _stat(candidate) != receipt.get(stat_field):
                return False
        except OSError:
            return False
    return True


def export_study(
    study_root: str | Path,
    destination: str | Path,
    on_progress: Callable[[dict], None] | None = None,
) -> dict:
    """Copy stable prepared bundles without changing the active study.

    Receipts avoid reading unchanged payloads on later calls. Each source hash
    can retain several bundles; different preparations are never overwritten.
    """
    study_root = Path(study_root).resolve()
    destination = Path(destination).resolve()
    summary = {"available": 0, "exported": 0, "existing": 0, "pending": 0, "errors": []}
    prepared = study_root / "prepared"
    if not prepared.is_dir():
        return summary
    study_key = hashlib.sha256(str(study_root).encode("utf-8")).hexdigest()
    receipts = destination / ".receipts" / study_key
    pending_root = destination / ".pending"
    for source in sorted(prepared.glob("*.zip")):
        summary["available"] += 1
        pending = None
        try:
            record_path = study_root / "records" / (source.stem + ".json")
            record_before = _stat(record_path)
            record = _read_json(record_path)
            if _stat(record_path) != record_before or not record.get("prepared"):
                summary["pending"] += 1
                continue
            source_sha = record.get("source_sha256", "")
            if not isinstance(source_sha, str) or not _HASH.fullmatch(source_sha):
                raise ValueError("Study record has no valid original source SHA-256")
            bundle_before = _stat(source)
            receipt_path = receipts / (source.stem + ".json")
            try:
                receipt = _read_json(receipt_path)
            except (OSError, ValueError):
                receipt = {}
            if _cached(receipt, destination, bundle_before, source_sha):
                summary["existing"] += 1
                continue

            pending_root.mkdir(parents=True, exist_ok=True)
            pending = pending_root / (uuid.uuid4().hex + ".zip")
            shutil.copyfile(source, pending)
            if _stat(source) != bundle_before:
                summary["pending"] += 1
                continue
            latest_record = _read_json(record_path)
            identity_fields = ("source_sha256", "baseline_sha256", "prepared_by")
            if not latest_record.get("prepared") or any(
                latest_record.get(field) != record.get(field) for field in identity_fields
            ):
                summary["pending"] += 1
                continue
            bundle_sha = _hash_file(pending)
            files = _inspect_bundle(pending, record)
            bundle_relative = f"bundles/{bundle_sha}.zip"
            cached_bundle = destination / bundle_relative
            cached_bundle.parent.mkdir(parents=True, exist_ok=True)
            # A missing or damaged copy can be replaced from this verified input.
            if not cached_bundle.exists() or _hash_file(cached_bundle) != bundle_sha:
                os.replace(pending, cached_bundle)
                pending = None
            entry_relative = f"entries/{source_sha}/{bundle_sha}.json"
            entry_path = destination / entry_relative
            entry = {
                "schema_version": 1,
                "source_sha256": source_sha,
                "bundle_sha256": bundle_sha,
                "bundle_path": bundle_relative,
                "baseline_repaired": record.get("baseline_repaired"),
                "record": record,
                "provenance": _provenance(study_root, record),
                "files": files,
            }
            # Metadata is published only after the payload is complete.
            if not entry_path.exists():
                _write_json(entry_path, entry)
            _write_json(receipt_path, {
                "source_sha256": source_sha,
                "bundle_stat": bundle_before,
                "bundle_path": bundle_relative,
                "cached_bundle_stat": _stat(cached_bundle),
                "entry_path": entry_relative,
                "entry_stat": _stat(entry_path),
            })
            summary["exported"] += 1
            if on_progress:
                on_progress(dict(summary))
        except (FileNotFoundError, PermissionError, json.JSONDecodeError):
            # The study can replace its record or remove a bundle between scans.
            summary["pending"] += 1
        except (OSError, ValueError, zipfile.BadZipFile, RuntimeError) as error:
            if len(summary["errors"]) < 20:
                summary["errors"].append({"bundle": source.name, "error": str(error)})
        finally:
            if pending is not None:
                pending.unlink(missing_ok=True)
    return summary


def preserve_once(study_root: str | Path, destination: str | Path) -> dict:
    study_key = hashlib.sha256(str(Path(study_root).resolve()).encode("utf-8")).hexdigest()
    snapshot = snapshot_study(study_root, Path(destination) / ".studies" / study_key)
    result = export_study(snapshot["snapshot_root"], destination)
    result["snapshot"] = snapshot
    return result


def watch_study(
    study_root: str | Path,
    destination: str | Path,
    stop_path: str | Path,
    interval: float = 10,
) -> None:
    """Keep new bundles while the study runs; a stop file ends the watcher."""
    if interval <= 0:
        raise ValueError("Watch interval must be positive")
    stop_path = Path(stop_path)
    previous_errors = None
    study_key = hashlib.sha256(str(Path(study_root).resolve()).encode("utf-8")).hexdigest()
    snapshot_root = Path(destination) / ".studies" / study_key
    last_snapshot = 0.0

    def capture(_progress=None):
        nonlocal last_snapshot
        if stop_path.exists() or time.monotonic() - last_snapshot < interval:
            return
        snapshot = snapshot_study(study_root, snapshot_root)
        last_snapshot = time.monotonic()
        if snapshot["captured"] or snapshot["errors"]:
            print(json.dumps({"snapshot": snapshot}), flush=True)

    print(f"Preserving study bundles in {Path(destination).resolve()}", flush=True)
    while not stop_path.exists():
        # Capture every current bundle before starting the slower content index.
        capture()
        summary = export_study(snapshot_root, destination, on_progress=capture)
        errors = summary["errors"]
        if summary["exported"] or errors != previous_errors:
            print(json.dumps(summary), flush=True)
        previous_errors = errors
        time.sleep(interval)
    print("Study bundle export stopped", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("study_root", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--watch", action="store_true")
    parser.add_argument("--stop-file", type=Path)
    parser.add_argument("--interval", type=float, default=10)
    args = parser.parse_args()
    if args.watch:
        if args.stop_file is None:
            parser.error("--watch requires --stop-file")
        watch_study(args.study_root, args.destination, args.stop_file, args.interval)
    else:
        print(json.dumps(preserve_once(args.study_root, args.destination)))


if __name__ == "__main__":
    main()
