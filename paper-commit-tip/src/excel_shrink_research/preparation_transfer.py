"""Copy preparation transfers, then merge them only into an idle study."""

import argparse
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import tempfile
import uuid
import zipfile

from .enron.common import long_path, read_json, sha256, utc_now, write_json


PROFILES = ("default", "maximum-cleaning", "too-aggressive-cleaning")
WORKBOOKS = ("original.xlsx",) + tuple(name + ".xlsx" for name in PROFILES)
_HASH = re.compile(r"[0-9a-f]{64}\Z")
_KEY = re.compile(r"[0-9a-f]{20}\Z")


class StudyBusy(RuntimeError):
    pass


@contextmanager
def _lock(path, message):
    path.parent.mkdir(parents=True, exist_ok=True)
    stream = long_path(path).open("a+b")
    try:
        stream.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            raise StudyBusy(message) from None
        yield
    finally:
        stream.close()


def _relative(value):
    if not isinstance(value, str) or not value or "\\" in value:
        raise ValueError("A transfer path must be a relative POSIX path.")
    parts = PurePosixPath(value).parts
    if not parts or value != "/".join(parts) or any(part in (".", "..", "/") for part in parts):
        raise ValueError("Unsafe transfer path: " + value)
    if any(any(ord(c) < 32 or c in '<>:"|?*' for c in part) or part.endswith((".", " ")) for part in parts):
        raise ValueError("Unsupported transfer path: " + value)
    return Path(*parts)


def _path(root, relative):
    root = Path(root).resolve()
    path = root / _relative(relative)
    path.resolve().relative_to(root)
    if path.is_symlink():
        raise ValueError("Transfer symlinks are not followed: " + relative)
    return long_path(path)


def _hash(value, label):
    if not isinstance(value, str) or not _HASH.fullmatch(value):
        raise ValueError("Invalid SHA-256 for " + label)
    return value


def _load_manifest(root):
    root = Path(root).resolve()
    if (root / "TRANSFER_INCOMPLETE.txt").exists() or not (root / "COPY_VERIFIED.txt").is_file():
        raise ValueError("The transfer is not marked as completely copied.")
    expected = (root / "manifest.sha256").read_text(encoding="utf-8-sig").split()
    if not expected:
        raise ValueError("The manifest checksum is missing.")
    digest = sha256(root / "manifest.json")
    if digest != _hash(expected[0].lower(), "manifest.json"):
        raise ValueError("The transfer manifest checksum does not match.")
    manifest = read_json(root / "manifest.json")
    if (manifest.get("format") != "excel-shrink-preparation-transfer" or
            manifest.get("schema_version") != 1 or manifest.get("complete") is not True):
        raise ValueError("Unsupported or incomplete preparation transfer.")
    if manifest.get("contains_completed_comparison_claims") is not False:
        raise ValueError("A preparation transfer must not assert completed visual comparisons.")
    corpus = manifest.get("corpus", {})
    for field in ("id", "source_fingerprint", "catalog_sha256"):
        _hash(corpus.get(field), "corpus " + field)
    _relative(corpus.get("archive_path"))
    if type(corpus.get("total_workbooks")) is not int or corpus["total_workbooks"] < 1:
        raise ValueError("The original corpus workbook count is missing.")
    selected = manifest.get("size_part", {})
    if (type(selected.get("part")) is not int or type(selected.get("parts")) is not int or
            not 1 <= selected["part"] <= selected["parts"]):
        raise ValueError("The transferred size partition is invalid.")
    begin = corpus["total_workbooks"] * (selected["part"] - 1) // selected["parts"]
    end = corpus["total_workbooks"] * selected["part"] // selected["parts"]
    if manifest.get("counts", {}).get("records") != end - begin:
        raise ValueError("The transfer does not contain every record in its selected partition.")
    if set(manifest.get("profiles", {})) != set(PROFILES):
        raise ValueError("The transfer must identify all three profile configurations.")
    builds = manifest.get("preparation_builds", [])
    if not isinstance(builds, list) or len(set(builds)) != len(builds):
        raise ValueError("The producer build list is invalid.")
    for build in builds:
        _hash(build, "producer build")
    files = manifest.get("files")
    if not isinstance(files, list) or not files:
        raise ValueError("The transfer has no payload manifest.")
    seen, total = set(), 0
    for row in files:
        name = row.get("path")
        _path(root, name)
        folded = name.casefold()
        if folded in seen:
            raise ValueError("Duplicate transfer path: " + name)
        seen.add(folded)
        if type(row.get("bytes")) is not int or row["bytes"] < 0:
            raise ValueError("Invalid payload size: " + name)
        _hash(row.get("sha256"), name)
        total += row["bytes"]
    if manifest.get("total_payload_bytes") != total:
        raise ValueError("The payload size total does not match the manifest.")
    return manifest, digest


def _verify_file(path, row):
    if not path.is_file() or path.stat().st_size != row["bytes"] or sha256(path) != row["sha256"]:
        raise ValueError("Payload differs from its manifest: " + row["path"])


def _stamp(path):
    value = path.stat()
    return {"size": value.st_size, "mtime_ns": value.st_mtime_ns}


def _copy_verified(source, target, row):
    if target.exists():
        _verify_file(target, row)
        return False
    target.parent.mkdir(parents=True, exist_ok=True)
    pending = target.with_name(target.name + "." + uuid.uuid4().hex + ".partial")
    digest, copied = hashlib.sha256(), 0
    try:
        with source.open("rb") as input_stream, pending.open("xb") as output:
            for block in iter(lambda: input_stream.read(1024 * 1024), b""):
                copied += len(block)
                if copied > row["bytes"]:
                    raise ValueError("Payload exceeds its recorded size: " + row["path"])
                digest.update(block)
                output.write(block)
            output.flush()
            os.fsync(output.fileno())
        if copied != row["bytes"] or digest.hexdigest() != row["sha256"]:
            raise ValueError("Payload differs from its manifest: " + row["path"])
        _verify_file(pending, row)
        os.replace(pending, target)
        return True
    finally:
        pending.unlink(missing_ok=True)


def stage_transfer(source, destination, on_progress=None):
    """Copy every payload with SHA-256 checks, without changing a running study."""
    source, destination = Path(source).resolve(), Path(destination).resolve()
    if source == destination or destination.is_relative_to(source) or source.is_relative_to(destination):
        raise ValueError("Source and staged transfer directories must be separate.")
    manifest, digest = _load_manifest(source)
    destination.mkdir(parents=True, exist_ok=True)
    summary = {"state": "copying", "manifest_sha256": digest, "source": str(source),
               "destination": str(destination), "files_total": len(manifest["files"]),
               "bytes_total": manifest["total_payload_bytes"], "files_copied": 0,
               "files_existing": 0, "files_verified": 0, "bytes_verified": 0}
    status = destination / "stage-status.json"
    with _lock(destination / ".transfer.lock", "Another process is copying this transfer."):
        try:
            if (destination / "manifest.json").exists() and sha256(destination / "manifest.json") != digest:
                raise ValueError("The staging directory belongs to a different transfer.")
            verified_files = {}
            write_json(status, summary)
            for row in manifest["files"]:
                changed = _copy_verified(_path(source, row["path"]), _path(destination, row["path"]), row)
                verified_files[row["path"]] = _stamp(_path(destination, row["path"]))
                summary["files_copied" if changed else "files_existing"] += 1
                summary["files_verified"] += 1
                summary["bytes_verified"] += row["bytes"]
                if summary["files_verified"] % 25 == 0 or summary["files_verified"] == summary["files_total"]:
                    summary.update(updated_utc=utc_now(), current_file=row["path"])
                    write_json(status, summary)
                    if on_progress:
                        on_progress({"event": "transfer_copy", **summary})
            for name in ("manifest.json", "manifest.sha256", "AGENT_README.txt", "COPY_VERIFIED.txt"):
                source_file = source / name
                if source_file.is_file():
                    row = {"path": name, "bytes": source_file.stat().st_size, "sha256": sha256(source_file)}
                    _copy_verified(long_path(source_file), _path(destination, name), row)
            _load_manifest(destination)
            receipt = destination / "verified-files.json"
            write_json(receipt, {"manifest_sha256": digest, "files": verified_files})
            summary.update(state="staged", verified_files_sha256=sha256(receipt), updated_utc=utc_now())
            write_json(status, summary)
            return summary
        except Exception as error:
            summary.update(state="failed", error=str(error), updated_utc=utc_now())
            write_json(status, summary)
            raise


def _check_record(record, key, manifest, prepared):
    identity = record.get("identity")
    corpus = manifest["corpus"]
    if not isinstance(identity, str) or hashlib.sha256(identity.encode()).hexdigest()[:20] != key:
        raise ValueError("The record filename and identity disagree: " + key)
    pieces = identity.split(":", 2)
    selection = record.get("run_selection", {})
    if (len(pieces) != 3 or pieces[0] != corpus["id"] or not pieces[1].isdigit() or
            record.get("corpus_id") != corpus["id"] or record.get("source_fingerprint") != corpus["source_fingerprint"] or
            selection.get("archive_index") != int(pieces[1]) or selection.get("corpus_id") != corpus["id"] or
            selection.get("catalog_sha256") != corpus["catalog_sha256"] or
            selection.get("source_fingerprint") != corpus["source_fingerprint"] or
            selection.get("size_part") != manifest["size_part"]):
        raise ValueError("The record does not match the transferred corpus selection: " + key)
    if prepared:
        if (record.get("prepared") is not True or record.get("complete") is not False or record.get("status") != "prepared" or
                record.get("prepared_by") not in manifest["preparation_builds"]):
            raise ValueError("The incoming record is not a reusable preparation: " + key)
        _hash(record.get("source_sha256"), key)
        if record.get("baseline_sha256") is not None:
            _hash(record["baseline_sha256"], key)
        results = record.get("results", [])
        if len(results) != len(PROFILES) or {row.get("profile") for row in results} != set(PROFILES):
            raise ValueError("The prepared record does not contain all three profiles: " + key)
        for row in results:
            if (row.get("config_hash") != manifest["profiles"][row["profile"]] or
                    row.get("checks_complete") is not False or row.get("status") != "not compared"):
                raise ValueError("The incoming preparation has inconsistent profile or comparison claims: " + key)
    elif record.get("prepared") is not False or record.get("status") != "skipped" or record.get("complete") is not True:
        raise ValueError("The incoming failure record is not a preparation exception: " + key)
    return int(pieces[1]), pieces[2]


def _check_xlsx(stream):
    stream.seek(0)
    with zipfile.ZipFile(stream) as workbook:
        names = workbook.namelist()
        if names.count("[Content_Types].xml") != 1 or names.count("xl/workbook.xml") != 1:
            raise ValueError("A prepared file is not an XLSX workbook package.")
        declaration = workbook.getinfo("[Content_Types].xml")
        if declaration.file_size > 8 * 1024 * 1024 or b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" not in workbook.read(declaration):
            raise ValueError("A prepared file does not declare an XLSX workbook.")
        if any(entry.flag_bits & 1 for entry in workbook.infolist()):
            raise ValueError("A prepared XLSX package contains encrypted ZIP entries.")


def _check_bundle(path, record, manifest, work):
    work.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as bundle:
        entries = bundle.infolist()
        names = [entry.filename for entry in entries]
        for name in WORKBOOKS + ("measurements.json",):
            if names.count(name) != 1:
                raise ValueError("A preparation bundle is missing or duplicates " + name)
        metadata = {}
        for entry in entries:
            _relative(entry.filename)
            if entry.is_dir() or (entry.external_attr >> 16) & 0o170000 == stat.S_IFLNK or entry.flag_bits & 1:
                raise ValueError("Preparation bundles cannot contain directories, links or encrypted entries.")
            digest = hashlib.sha256()
            if entry.filename.endswith(".xlsx"):
                with tempfile.TemporaryFile(dir=work) as staged:
                    with bundle.open(entry) as stream:
                        for block in iter(lambda: stream.read(1024 * 1024), b""):
                            digest.update(block)
                            staged.write(block)
                    _check_xlsx(staged)
            else:
                with bundle.open(entry) as stream:
                    for block in iter(lambda: stream.read(1024 * 1024), b""):
                        digest.update(block)
            metadata[entry.filename] = {"bytes": entry.file_size, "sha256": digest.hexdigest()}
        if record.get("baseline_sha256") and metadata["original.xlsx"]["sha256"] != record["baseline_sha256"]:
            raise ValueError("The prepared baseline differs from the record.")
        if bundle.getinfo("measurements.json").file_size > 8 * 1024 * 1024:
            raise ValueError("The preparation measurements are unexpectedly large.")
        measurements = json.loads(bundle.read("measurements.json").decode("utf-8-sig"))
        if not isinstance(measurements, list) or len(measurements) != len(PROFILES):
            raise ValueError("The bundle does not contain measurements for all profiles.")
        actual = {item.get("profile"): item for item in measurements}
        if set(actual) != set(PROFILES):
            raise ValueError("The bundle profile measurements are duplicated or incomplete.")
        for result in record["results"]:
            name = result["profile"]
            measured = actual[name]
            if (measured.get("config_hash") != manifest["profiles"][name] or
                    measured.get("shrink_seconds") != result.get("shrink_seconds") or
                    measured.get("shrunk_bytes") != metadata[name + ".xlsx"]["bytes"] or
                    measured.get("original_bytes") != metadata["original.xlsx"]["bytes"]):
                raise ValueError("The record, workbook sizes and preparation measurements disagree.")


def _references(record):
    value = record.get("provenance", {}).get("preparation", {})
    if value.get("machine_manifest"):
        yield "study/" + value["machine_manifest"]
    if value.get("run_id"):
        yield "study/metadata/runs/" + value["run_id"] + ".json"


def _validate_payload(root, manifest, progress=None):
    files = {row["path"]: row for row in manifest["files"]}
    verified = {}
    try:
        stage = read_json(root / "stage-status.json")
        receipt_path = root / "verified-files.json"
        receipt = read_json(receipt_path)
        if (stage.get("state") == "staged" and stage.get("manifest_sha256") == sha256(root / "manifest.json") == receipt.get("manifest_sha256")
                and stage.get("verified_files_sha256") == sha256(receipt_path)):
            verified = receipt["files"]
    except (OSError, ValueError, KeyError):
        pass
    for number, row in enumerate(files.values(), 1):
        path = _path(root, row["path"])
        if not path.is_file() or verified.get(row["path"]) != _stamp(path):
            _verify_file(path, row)
        if progress and number % 100 == 0:
            progress({"stage": "verifying_payload", "completed": number, "total": len(files)})
    records = []
    seen, observed_builds = set(), set()
    source_path = manifest["corpus"]["archive_path"]
    if source_path not in files:
        raise ValueError("The unchanged source archive is missing from the payload manifest.")
    with zipfile.ZipFile(_path(root, source_path)) as source:
        source_entries = source.infolist()
        ranked = sorted((info.file_size, info.orig_filename, index)
                        for index, info in enumerate(source_entries)
                        if not info.is_dir() and info.orig_filename.lower().endswith((".xls", ".xlsx")))
        if len(ranked) != manifest["corpus"]["total_workbooks"]:
            raise ValueError("The unchanged source ZIP has a different workbook count.")
        ranks = {entry[2]: number + 1 for number, entry in enumerate(ranked)}
        selection = manifest["size_part"]
        first = len(ranked) * (selection["part"] - 1) // selection["parts"] + 1
        last = len(ranked) * selection["part"] // selection["parts"]
        for name, metadata in files.items():
            if not name.startswith("study/records/"):
                continue
            key = Path(name).stem
            if not _KEY.fullmatch(key) or name != "study/records/" + key + ".json":
                raise ValueError("Unexpected study record path: " + name)
            record = read_json(_path(root, name))
            prepared = record.get("prepared") is True
            index, member_name = _check_record(record, key, manifest, prepared)
            if index in seen or index >= len(source_entries) or source_entries[index].orig_filename != member_name:
                raise ValueError("The preparation identity does not match the unchanged source archive.")
            rank = ranks.get(index)
            recorded = record["run_selection"]
            if (rank is None or not first <= rank <= last or recorded.get("rank") != rank or
                    recorded.get("bytes") != source_entries[index].file_size):
                raise ValueError("The preparation does not belong to its claimed size partition.")
            seen.add(index)
            original_hash = hashlib.sha256()
            with source.open(source_entries[index]) as stream:
                for block in iter(lambda: stream.read(1024 * 1024), b""):
                    original_hash.update(block)
            if record.get("source_sha256") and original_hash.hexdigest() != record["source_sha256"]:
                raise ValueError("The original source bytes do not match their preparation: " + key)
            for reference in _references(record):
                if reference not in files or not reference.startswith("study/metadata/"):
                    raise ValueError("A preparation provenance reference is missing: " + reference)
            bundle_name = "study/prepared/" + key + ".zip"
            if prepared:
                if bundle_name not in files:
                    raise ValueError("The prepared workbook bundle is not in the manifest: " + key)
                _check_bundle(_path(root, bundle_name), record, manifest, root / ".validation")
                observed_builds.add(record["prepared_by"])
            elif bundle_name in files:
                raise ValueError("A failed preparation unexpectedly has a prepared bundle: " + key)
            records.append((key, record, prepared))
            if progress and len(records) % 25 == 0:
                progress({"stage": "checking_preparations", "completed": len(records), "total": manifest["counts"]["records"]})
    counts = manifest["counts"]
    if (len(records) != counts["records"] or sum(item[2] for item in records) != counts["prepared"] or
            sum(not item[2] for item in records) != counts["failed_preparation"] or
            observed_builds != set(manifest["preparation_builds"])):
        raise ValueError("Preparation counts or producer build identifiers disagree with the manifest.")
    return records, files


def _copy_atomic(source, target):
    target.parent.mkdir(parents=True, exist_ok=True)
    pending = target.with_name(target.name + "." + uuid.uuid4().hex + ".pending")
    try:
        shutil.copyfile(source, pending)
        os.replace(pending, target)
    finally:
        pending.unlink(missing_ok=True)


def merge_transfer(staged, study, on_progress=None, *, expected_source=None, expected_profiles=None):
    """Merge verified preparations under the Rust study lock, retaining better local results."""
    staged, study = Path(staged).resolve(), Path(study).resolve()
    manifest, digest = _load_manifest(staged)
    corpus = manifest["corpus"]
    if expected_source is None:
        expected_source = read_json(study / "source-catalogs" / (corpus["id"] + ".json"))
        expected_source["id"] = expected_source.get("corpus_id")
    for field in ("id", "source_fingerprint", "catalog_sha256"):
        if expected_source.get(field) != corpus.get(field):
            raise ValueError("The transferred corpus does not match the local " + field)
    if expected_profiles is None or expected_profiles != manifest.get("profiles"):
        raise ValueError("Current profile configuration hashes must match before preparation reuse.")
    archive_row = next((row for row in manifest["files"] if row["path"] == corpus["archive_path"]), {})
    if expected_source.get("archive_sha256") and expected_source["archive_sha256"] != archive_row.get("sha256"):
        raise ValueError("The transferred source archive differs from the local source archive.")
    summary = {"state": "checking", "manifest_sha256": digest, "imported_prepared": 0,
               "imported_failed": 0, "retained_existing": 0, "conflicts": [], "metadata_copied": 0,
               "approved_preparation_builds": [], "updated_utc": utc_now()}
    def progress(event):
        summary.update(event, updated_utc=utc_now())
        write_json(staged / "merge-status.json", summary)
        if on_progress:
            on_progress({"event": "transfer_merge", **summary})
    with _lock(study / ".run.lock", "The study is active; preparations have not been merged."):
        try:
            records, files = _validate_payload(staged, manifest, progress)
            metadata = [row for name, row in files.items() if name.startswith("study/metadata/")]
            for row in metadata:
                relative = row["path"][len("study/"):]
                target = _path(study, relative)
                if target.exists() and sha256(target) != row["sha256"]:
                    raise ValueError("An immutable local metadata file conflicts with the transfer: " + relative)
            for row in metadata:
                target = _path(study, row["path"][len("study/"):])
                if not target.exists():
                    _copy_atomic(_path(staged, row["path"]), target)
                    summary["metadata_copied"] += 1
            for number, (key, incoming, prepared) in enumerate(records, 1):
                relative = "records/" + key + ".json"
                target_record = _path(study, relative)
                target_bundle = _path(study, "prepared/" + key + ".zip")
                existing = read_json(target_record) if target_record.exists() else None
                if existing is not None:
                    if (existing.get("identity") != incoming["identity"] or
                            existing.get("source_fingerprint") not in (None, corpus["source_fingerprint"]) or
                            (existing.get("source_sha256") and incoming.get("source_sha256") and existing["source_sha256"] != incoming["source_sha256"])):
                        summary["conflicts"].append({"key": key, "reason": "Existing and incoming source identities differ.", "incoming_record": str(staged / "study" / relative)})
                        continue
                    valid_existing = existing.get("complete") is True and existing.get("status") in ("unchanged", "differences")
                    if existing.get("prepared") and target_bundle.is_file():
                        try:
                            if existing.get("identity") != incoming["identity"] or existing.get("source_fingerprint") != corpus["source_fingerprint"]:
                                raise ValueError("Existing preparation identity differs.")
                            _check_bundle(target_bundle, existing, manifest, staged / ".validation")
                            valid_existing = True
                        except (OSError, ValueError, KeyError, zipfile.BadZipFile, RuntimeError):
                            pass
                    if valid_existing or not prepared:
                        summary["retained_existing"] += 1
                        continue
                    history = study / "import-history" / digest
                    for path, name in ((target_record, relative), (target_bundle, "prepared/" + key + ".zip")):
                        if path.is_file() and not _path(history, name).exists():
                            _copy_atomic(path, _path(history, name))
                elif target_bundle.exists():
                    summary["conflicts"].append({"key": key, "reason": "A local bundle has no matching record."})
                    continue
                if prepared:
                    _copy_atomic(_path(staged, "study/prepared/" + key + ".zip"), target_bundle)
                _copy_atomic(_path(staged, "study/" + relative), target_record)
                summary["imported_prepared" if prepared else "imported_failed"] += 1
                if number % 25 == 0:
                    progress({"stage": "merging_records", "completed": number, "total": len(records)})
            summary.update(state="merged_with_conflicts" if summary["conflicts"] else "merged",
                           approved_preparation_builds=manifest["preparation_builds"], approved_builds=manifest["preparation_builds"])
            progress({"stage": "complete", "completed": len(records), "total": len(records)})
            return summary
        except Exception as error:
            summary.update(state="failed", error=str(error))
            progress({"stage": "failed"})
            raise


def main(argv=None):
    parser = argparse.ArgumentParser(description="Copy a preparation transfer with checksum verification.")
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args(argv)
    result = stage_transfer(args.source, args.destination)
    print(json.dumps(result, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
