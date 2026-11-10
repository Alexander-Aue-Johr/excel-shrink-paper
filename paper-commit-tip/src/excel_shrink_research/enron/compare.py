"""Compare original Enron attachment bytes with an existing spreadsheet corpus."""

from collections import Counter, defaultdict
from contextlib import ExitStack
import csv
import hashlib
import json
from pathlib import Path
import tempfile
import time
import zipfile
import zlib

from . import ingest
from .common import RunLock, long_path, read_json, rows, sha256, utc_now, write_json


SCHEMA = 1
_READ_ERRORS = (OSError, RuntimeError, ValueError, zipfile.BadZipFile, EOFError, zlib.error)
_LOCATION = ("archive", "member_chain", "index_chain")
_FILE_COLUMNS = ("sha256", "size", "kind", "classification") + _LOCATION


def _code_stamp():
    return {"scanner_sha256": sha256(Path(__file__)),
            "classifier_sha256": sha256(Path(ingest.__file__))}


def _stat(path):
    value = path.stat()
    return {"size": value.st_size, "mtime_ns": value.st_mtime_ns}


def _json_line(output, value):
    output.write(json.dumps(value, ensure_ascii=True) + "\n")


def _valid_hash(value):
    return isinstance(value, str) and len(value) == 64 and all(c in "0123456789abcdef" for c in value)


def _cached_json(path):
    try:
        value = read_json(path)
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError):
        return {}


def _location(archive, indices, names, **extra):
    return {"archive": archive, "id": ingest._entry_id(archive, indices),
            "index_chain": list(indices), "member_chain": list(names), **extra}


def index_old_archive(archive_path, destination, on_progress=None):
    """Hash every original file; known spreadsheet suffixes define the legacy scope."""
    archive_path = long_path(archive_path)
    destination = long_path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    index = destination / "old-files.jsonl"
    marker = destination / "old-index-summary.json"
    stamp = {"schema": SCHEMA, "source": str(archive_path),
             "source_size": archive_path.stat().st_size,
             "source_mtime_ns": archive_path.stat().st_mtime_ns, **_code_stamp()}
    with RunLock(destination / ".old-index"):
        if index.is_file() and marker.is_file():
            cached = _cached_json(marker)
            if all(cached.get(key) == value for key, value in stamp.items()) and cached.get("index_sha256") == sha256(index):
                return index
        pending = index.with_name(index.name + ".partial")
        counts = Counter()
        digests, scoped_digests = set(), set()
        errors = []
        started = time.monotonic()
        with zipfile.ZipFile(archive_path) as archive, pending.open("w", encoding="utf-8", newline="\n") as output:
            entries = archive.infolist()
            total = sum(not info.is_dir() for info in entries)
            for number, info in enumerate(entries):
                if info.is_dir():
                    counts["directories"] += 1
                    continue
                row = _location(archive_path.name, (number,), (info.orig_filename,),
                                member=info.orig_filename, index=number, size=info.file_size,
                                crc32=info.CRC, sha256=None, kind=None, classification=None,
                                spreadsheet_scope=Path(info.orig_filename).suffix.lower() in ingest.SPREADSHEET_SUFFIXES,
                                status="ok")
                try:
                    digest = hashlib.sha256()
                    size = 0
                    with archive.open(info) as stream:
                        prefix = stream.read(65536)
                        kind, classification, evidence = ingest._detect(prefix, info.orig_filename)
                        row.update(kind=kind, classification=classification, evidence=evidence)
                        row["spreadsheet_scope"] |= classification == "candidate"
                        block = prefix
                        while block:
                            digest.update(block)
                            size += len(block)
                            block = stream.read(ingest.CHUNK_SIZE)
                    if size != info.file_size:
                        raise ValueError("The member size differs from the ZIP directory.")
                    row["sha256"] = digest.hexdigest()
                    digests.add(row["sha256"])
                    if row["spreadsheet_scope"]:
                        scoped_digests.add(row["sha256"])
                    counts["hashed"] += 1
                except _READ_ERRORS as error:
                    row.update(status="unreadable", error=str(error))
                    errors.append({"member": info.orig_filename, "index": number, "error": str(error)})
                counts["files"] += 1
                counts["spreadsheet_scope_files"] += row["spreadsheet_scope"]
                _json_line(output, row)
                if on_progress and counts["files"] % 1000 == 0:
                    on_progress({"event": "legacy_index", "completed": counts["files"], "total": total})
        source_sha256 = sha256(archive_path)
        current = archive_path.stat()
        if current.st_size != stamp["source_size"] or current.st_mtime_ns != stamp["source_mtime_ns"]:
            raise RuntimeError("The existing corpus archive changed during indexing.")
        pending.replace(index)
        write_json(marker, {**stamp, "state": "complete" if not errors else "complete_with_errors",
                            "source_sha256": source_sha256, "index_sha256": sha256(index),
                            "counts": dict(counts), "unique_hashes": len(digests),
                            "spreadsheet_scope_unique_hashes": len(scoped_digests),
                            "errors": errors, "elapsed_seconds": round(time.monotonic() - started, 1)})
    return index


def _old_entries(path):
    entries = list(rows(path))
    by_hash = defaultdict(list)
    sizes = set()
    unreadable = 0
    for entry in entries:
        if entry.get("status", "ok") != "ok" or not _valid_hash(entry.get("sha256")):
            unreadable += 1
            continue
        by_hash[entry["sha256"]].append(entry)
        sizes.add(entry["size"])
    return entries, by_hash, sizes, unreadable


def _inventory_source(root, source):
    inventory = root / "inventory" / (source["name"] + ".jsonl")
    marker = inventory.with_name(source["name"] + ".complete.json")
    if not inventory.is_file() or not marker.is_file():
        return None
    try:
        recorded = read_json(marker)
        expected = {"manifest": ingest._fingerprint(source), **_stat(root / "raw" / source["name"])}
        if not all(recorded.get("source", {}).get(key) == value for key, value in expected.items()):
            return None
        return inventory, recorded
    except (OSError, ValueError, KeyError):
        return None


def _source_stamp(root, source, code, old_digest, inventory):
    raw = root / "raw" / source["name"]
    receipt_path = root / "raw" / ".verified" / (source["name"] + ".json")
    receipt = _cached_json(receipt_path) if receipt_path.is_file() else None
    stat = _stat(raw) if raw.is_file() else None
    verified = bool(receipt and stat and all(receipt.get(key) == value for key, value in stat.items())
                    and receipt.get("source") == ingest._fingerprint(source) and _valid_hash(receipt.get("sha256")))
    stamp = {"schema": SCHEMA, "source": source, "raw": stat, "verification": receipt,
             "old_index_sha256": old_digest, **code}
    if inventory:
        stamp["inventory"] = {**_stat(inventory[0]), "marker": inventory[1]}
    return stamp, verified


def _issue(output, location, reason, message):
    _json_line(output, {**location, "record_type": "issue", "reason": reason, "message": message})


def _coverage_issue(row):
    if row.get("read_error") or row.get("kind") == "unreadable":
        return "unreadable_member", json.dumps(row.get("read_error", row.get("evidence", "Unreadable member.")))
    if row.get("nested_status") not in (None, "scanned", "complete", "completed"):
        return "incomplete_nested_archive", "Nested archive status: " + str(row["nested_status"])
    kind, classification = row.get("kind"), row.get("classification")
    if kind == "symlink":
        return "symlink", "A ZIP symlink was not followed."
    if kind in ("archive", "encrypted-archive"):
        return kind, row.get("evidence", "The nested archive could not be inspected.")
    if kind == "zip" and classification == "ambiguous":
        return "uninspected_zip", row.get("evidence", "The nested ZIP could not be inspected.")
    return None


def _reuse_inventory(path, marker, source, output, old_hashes):
    counts = Counter()
    for row in rows(path):
        counts["members"] += 1
        if row.get("archive") != source["name"]:
            raise ValueError("The cached inventory belongs to another source archive.")
        gap = _coverage_issue(row)
        if gap:
            _issue(output, row, *gap)
        digest = row.get("sha256")
        if digest is None and row.get("classification") != "directory" and not row.get("read_error") and row.get("kind") != "unreadable":
            raise ValueError("The cached inventory is missing a file content hash.")
        if digest is not None and not _valid_hash(digest):
            raise ValueError("The cached inventory contains an invalid content hash.")
        if row.get("classification") in ("candidate", "ambiguous") or digest in old_hashes:
            _json_line(output, {**row, "record_type": "entry"})
        counts["hashed"] += digest is not None
    if counts["members"] != marker.get("counts", {}).get("members"):
        raise ValueError("The cached inventory is incomplete.")
    return dict(counts)


def _scan_source(root, source, output, old_hashes, old_sizes, work, progress):
    counts = Counter()
    last_progress = [0.0]

    def pulse():
        if time.monotonic() - last_progress[0] >= 2:
            progress({"event": "comparison_scan", "archive": source["name"], **dict(counts)})
            last_progress[0] = time.monotonic()

    def scan(path, indices=(), names=(), depth=0):
        try:
            archive = zipfile.ZipFile(path)
        except _READ_ERRORS as error:
            _issue(output, _location(source["name"], indices, names), "unreadable_archive", str(error))
            return
        with archive:
            for number, info in enumerate(archive.infolist()):
                counts["members"] += 1
                chain = indices + (number,)
                member_names = names + (info.orig_filename,)
                row = _location(source["name"], chain, member_names, member=info.orig_filename,
                                index=number, size=info.file_size, crc32=info.CRC)
                if info.is_dir():
                    continue
                if info.flag_bits & 1:
                    _issue(output, row, "encrypted_member", "The ZIP member requires a password.")
                    continue
                temporary = None
                try:
                    symlink = (info.external_attr >> 16) & 0o170000 == 0o120000
                    with archive.open(info) as member:
                        prefix = member.read(65536)
                        kind, classification, evidence = ingest._detect(prefix, info.orig_filename)
                        if classification == "other" and not symlink and info.file_size not in old_sizes:
                            counts["prefix_only"] += 1
                            pulse()
                            continue
                        if kind in ("ole", "zip"):
                            handle = tempfile.NamedTemporaryFile(prefix="member-", suffix=".bin", dir=work, delete=False)
                            temporary = Path(handle.name)
                        else:
                            handle = None
                        digest, size = hashlib.sha256(), 0
                        try:
                            block = prefix
                            while block:
                                digest.update(block)
                                size += len(block)
                                if handle:
                                    handle.write(block)
                                block = member.read(ingest.CHUNK_SIZE)
                                pulse()
                        finally:
                            if handle:
                                handle.close()
                    if size != info.file_size:
                        raise ValueError("The member size differs from the ZIP directory.")
                    if temporary:
                        kind, classification, evidence = ingest._detect(prefix, info.orig_filename, temporary)
                    if symlink:
                        kind, classification, evidence = "symlink", "ambiguous", "A ZIP symlink was not followed."
                    row.update(sha256=digest.hexdigest(), kind=kind, classification=classification, evidence=evidence)
                    counts["hashed"] += 1
                    if classification == "container":
                        if depth >= ingest.MAX_NESTING or info.file_size > ingest.MAX_NESTED_BYTES:
                            row.update(classification="ambiguous", evidence="Nested ZIP exceeds the configured depth or size limit.")
                        else:
                            if row["sha256"] in old_hashes:
                                _json_line(output, {**row, "record_type": "entry"})
                            scan(temporary, chain, member_names, depth + 1)
                            continue
                    gap = _coverage_issue(row)
                    if gap:
                        _issue(output, row, *gap)
                    if row["classification"] in ("candidate", "ambiguous") or row["sha256"] in old_hashes:
                        _json_line(output, {**row, "record_type": "entry"})
                except _READ_ERRORS as error:
                    _issue(output, row, "unreadable_member", str(error))
                finally:
                    if temporary is not None:
                        temporary.unlink(missing_ok=True)
                pulse()

    scan(root / "raw" / source["name"])
    return dict(counts)


def _csv_row(row):
    return {key: json.dumps(value, ensure_ascii=True) if isinstance(value, (list, dict)) else value
            for key, value in row.items()}


def _render_reports(destination, caches, old_entries, old_by_hash, old_unreadable, total, final=False):
    columns = {
        "new-only.csv": _FILE_COLUMNS,
        "review-only.csv": _FILE_COLUMNS,
        "old-not-found.csv": ("status", "sha256", "size", "kind", "archive", "member_chain", "index_chain"),
        "matches.csv": ("sha256", "old_member", "old_index_chain", "old_spreadsheet_scope") + _LOCATION,
        "scan-issues.csv": _LOCATION + ("reason", "message"),
    }
    counts = Counter()
    found, candidate_hashes, review_hashes, new_hashes, review_only_hashes = set(), set(), set(), set(), set()
    pending = []
    with ExitStack() as stack:
        writers = {}
        for name, fields in columns.items():
            path = destination / (name + ".partial")
            pending.append((path, destination / name))
            stream = stack.enter_context(path.open("w", encoding="utf-8", newline=""))
            writer = csv.DictWriter(stream, fieldnames=fields, extrasaction="ignore")
            writer.writeheader()
            writers[name] = writer
        for cache, marker in caches:
            for key, value in marker["counts"].items():
                counts[key] += value
            for row in rows(cache):
                if row["record_type"] == "issue":
                    writers["scan-issues.csv"].writerow(_csv_row(row))
                    counts["scan_issues"] += 1
                    continue
                digest = row.get("sha256")
                if not _valid_hash(digest):
                    continue
                classification = row.get("classification")
                if classification == "candidate":
                    counts["candidate_occurrences"] += 1
                    candidate_hashes.add(digest)
                elif classification == "ambiguous":
                    counts["review_occurrences"] += 1
                    review_hashes.add(digest)
                if digest in old_by_hash:
                    found.add(digest)
                    counts["matching_new_occurrences"] += 1
                    for old in old_by_hash[digest]:
                        writers["matches.csv"].writerow(_csv_row({**row, "old_member": old["member"],
                            "old_index_chain": old.get("index_chain", [old["index"]]),
                            "old_spreadsheet_scope": old.get("spreadsheet_scope", False)}))
                elif classification == "candidate":
                    writers["new-only.csv"].writerow(_csv_row(row))
                    counts["new_only_occurrences"] += 1
                    new_hashes.add(digest)
                elif classification == "ambiguous":
                    writers["review-only.csv"].writerow(_csv_row(row))
                    counts["review_only_occurrences"] += 1
                    review_only_hashes.add(digest)
        old_scope_hashes = {row["sha256"] for row in old_entries
                            if row.get("spreadsheet_scope") and _valid_hash(row.get("sha256"))}
        absence_definitive = final and len(caches) == total and counts["scan_issues"] == 0 and old_unreadable == 0
        for old in old_entries:
            if not old.get("spreadsheet_scope"):
                continue
            counts["old_spreadsheet_occurrences"] += 1
            if _valid_hash(old.get("sha256")) and old["sha256"] in found:
                counts["old_found_occurrences"] += 1
            else:
                counts["old_not_found_occurrences"] += 1
                status = ("not_found" if absence_definitive else "not_found_in_scanned_archives") if _valid_hash(old.get("sha256")) else "old_file_unreadable"
                writers["old-not-found.csv"].writerow(_csv_row({**old, "status": status}))
    for temporary, target in pending:
        temporary.replace(target)
    processed = len(caches)
    complete = final and processed == total
    coverage = complete and counts["scan_issues"] == 0
    summary = {"schema": SCHEMA, "state": "running" if not complete else "complete" if coverage and not old_unreadable else "complete_with_gaps",
               "updated_utc": utc_now(), "archives_processed": processed, "archives_total": total,
               "old_index_entries": len(old_entries), "old_unreadable_entries": old_unreadable,
               "old_spreadsheet_unique_sha256": len(old_scope_hashes),
               "old_found_unique_sha256": len(old_scope_hashes & found),
               "old_not_found_unique_sha256": len(old_scope_hashes - found),
               "candidate_unique_sha256": len(candidate_hashes), "review_unique_sha256": len(review_hashes),
               "new_only_unique_sha256": len(new_hashes), "review_only_unique_sha256": len(review_only_hashes),
               "new_only_definitive": old_unreadable == 0, "old_not_found_definitive": coverage and old_unreadable == 0,
               "coverage_complete": coverage, "counts": dict(counts),
               "scope": "Original attachment bytes, including nested ZIP members. Candidate and review files are reported separately.",
               "integrity_scope": "Source download receipts are checked. Clear non-spreadsheet members may be read only through their first 64 KiB."}
    write_json(destination / "summary.json", summary)
    return summary


def compare_corpora(root, old_index_path, destination, on_progress=None):
    """Resume a content comparison and keep the CSV reports current after each archive."""
    root, destination, old_index_path = long_path(root), long_path(destination), long_path(old_index_path)
    destination.mkdir(parents=True, exist_ok=True)
    work, cache_root = destination / ".work", destination / "sources"
    work.mkdir(exist_ok=True)
    cache_root.mkdir(exist_ok=True)
    code = _code_stamp()
    old_digest = sha256(old_index_path)
    old_stamp = _stat(old_index_path)
    old_entries, old_by_hash, old_sizes, old_unreadable = _old_entries(old_index_path)
    if _stat(old_index_path) != old_stamp:
        raise RuntimeError("The existing corpus index changed while it was being read.")
    manifest = read_json(root / "sources.json")
    sources = manifest["archives"]
    caches = []
    summary = None

    def progress(event):
        write_json(destination / "status.json", {"state": "running", "updated_utc": utc_now(),
                   "archives_processed": len(caches), "archives_total": len(sources), **event})
        if on_progress:
            on_progress(event)

    with RunLock(destination):
        for source in sources:
            progress({"event": "comparison", "archive": source["name"], "completed": len(caches), "total": len(sources)})
            inventory = _inventory_source(root, source)
            stamp, verified = _source_stamp(root, source, code, old_digest, inventory)
            cache = cache_root / (source["name"] + ".jsonl")
            marker_path = cache_root / (source["name"] + ".complete.json")
            marker = _cached_json(marker_path) if marker_path.is_file() else {}
            if not (cache.is_file() and marker.get("stamp") == stamp and marker.get("cache_sha256") == sha256(cache)):
                temporary = cache.with_name(cache.name + ".partial")
                with temporary.open("w", encoding="utf-8", newline="\n") as output:
                    if not verified:
                        _issue(output, _location(source["name"], (), ()), "source_not_verified", "No matching source-download verification receipt is available.")
                    if not stamp["raw"] or stamp["raw"]["size"] != source["size"]:
                        _issue(output, _location(source["name"], (), ()), "source_missing_or_incomplete", "The source ZIP is missing or differs from its recorded size.")
                        counts, method = {}, "unavailable"
                    elif inventory:
                        try:
                            counts = _reuse_inventory(*inventory, source, output, old_by_hash)
                            method = "inventory"
                        except (OSError, ValueError, KeyError):
                            output.seek(0)
                            output.truncate()
                            if not verified:
                                _issue(output, _location(source["name"], (), ()), "source_not_verified", "No matching source-download verification receipt is available.")
                            counts = _scan_source(root, source, output, old_by_hash, old_sizes, work, progress)
                            method = "scan"
                    else:
                        counts = _scan_source(root, source, output, old_by_hash, old_sizes, work, progress)
                        method = "scan"
                current_raw = _stat(root / "raw" / source["name"]) if (root / "raw" / source["name"]).is_file() else None
                if current_raw != stamp["raw"]:
                    raise RuntimeError("The source archive changed during comparison: " + source["name"])
                temporary.replace(cache)
                marker = {"stamp": stamp, "counts": counts, "method": method, "cache_sha256": sha256(cache)}
                write_json(marker_path, marker)
            caches.append((cache, marker))
            summary = _render_reports(destination, caches, old_entries, old_by_hash, old_unreadable, len(sources))
            progress({"event": "comparison", "archive": source["name"], "completed": len(caches), "total": len(sources),
                      "new_only_unique_sha256": summary["new_only_unique_sha256"], "old_found_unique_sha256": summary["old_found_unique_sha256"]})
        summary = _render_reports(destination, caches, old_entries, old_by_hash, old_unreadable, len(sources), final=True)
        summary.update(old_index_sha256=old_digest, source_manifest_sha256=sha256(root / "sources.json"), **code)
        write_json(destination / "summary.json", summary)
        write_json(destination / "status.json", summary)
    return summary
