"""Download the original EUSES release and preserve inputs for the study runner."""

import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import sys
import tempfile
import urllib.request
import zipfile

from .enron.common import RunLock, read_json, sha256, utc_now, write_json
from .enron.ingest import _download, _entry_id, _fingerprint, _Paths, _verify

RECORD_ID = 581673
DOI = "10.5281/zenodo.581673"
METADATA_URL = "https://zenodo.org/api/records/581673"
ARCHIVE_NAMES = {"cs101.zip", "database.zip", "filby.zip", "financial.zip", "forms3.zip",
                 "grades.zip", "homework.zip", "inventory.zip", "jackson.zip", "modeling.zip", "personal.zip"}
EXPECTED_FILES = ARCHIVE_NAMES | {"README.txt"}
CORPUS_ID = hashlib.sha256(("https://doi.org/" + DOI).encode("utf-8")).hexdigest()


def _manifest_from_metadata(raw):
    data = json.loads(raw)
    if data.get("id") != RECORD_ID or data.get("doi") != DOI:
        raise ValueError("The metadata does not identify the original EUSES release.")
    files = []
    for entry in data["files"]:
        name = entry["key"]
        checksum = entry["checksum"].split(":", 1)
        if name not in EXPECTED_FILES or checksum[0] != "md5" or not re.fullmatch(r"[a-fA-F0-9]{32}", checksum[1]):
            raise ValueError("Unexpected EUSES source file or checksum: " + name)
        size = int(entry["size"])
        if size <= 0:
            raise ValueError("The source file has no recorded size: " + name)
        url = "https://zenodo.org/api/records/581673/files/" + name + "/content"
        files.append({"name": name, "size": size, "md5": checksum[1].lower(), "url": url})
    if len(files) != len(EXPECTED_FILES) or {entry["name"] for entry in files} != EXPECTED_FILES:
        raise ValueError("The original EUSES release must contain eleven ZIP files and README.txt.")
    return {"schema_version": 1, "name": "EUSES", "record_id": RECORD_ID, "doi": DOI,
            "record_url": "https://zenodo.org/records/581673", "metadata_url": METADATA_URL,
            "metadata_sha256": hashlib.sha256(raw).hexdigest(),
            "published_counts": {"total_files": 5607, "processed_files": 4499},
            "files": sorted(files, key=lambda entry: entry["name"])}


def freeze_sources(root):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    metadata = root / "source-metadata.json"
    if not metadata.exists():
        request = urllib.request.Request(METADATA_URL, headers={"User-Agent": "excel-shrink-research/1"})
        with urllib.request.urlopen(request, timeout=90) as response:
            raw = response.read()
        _manifest_from_metadata(raw)
        pending = metadata.with_suffix(".json.pending")
        pending.write_bytes(raw)
        pending.replace(metadata)
    manifest = _manifest_from_metadata(metadata.read_bytes())
    path = root / "sources.json"
    if path.exists():
        if read_json(path) != manifest:
            raise ValueError("The frozen EUSES source manifest changed.")
    else:
        write_json(path, manifest)
    return manifest


def download_euses(root, on_progress=None):
    """Resume and verify all twelve files from the pinned, original Zenodo release."""
    root = Path(root)
    with RunLock(root / ".download"):
        def progress(event):
            write_json(root / "download-status.json", {"state": "running", "updated_utc": utc_now(), **event})
            if on_progress:
                on_progress(event)

        try:
            manifest = freeze_sources(root)
            for index, source in enumerate(manifest["files"]):
                path = root / "raw" / source["name"]
                receipt_path = root / "raw/.verified" / (source["name"] + ".json")
                try:
                    receipt = read_json(receipt_path) if receipt_path.exists() else {}
                except (OSError, ValueError):
                    receipt = {}
                stat = path.stat() if path.exists() else None
                stamp = {"source": _fingerprint(source)}
                if stat:
                    stamp.update(size=stat.st_size, mtime_ns=stat.st_mtime_ns)
                if not stat or any(receipt.get(key) != value for key, value in stamp.items()):
                    hashes = _verify(path, source) if stat else _download(source, path, progress)
                    stat = path.stat()
                    stamp.update(size=stat.st_size, mtime_ns=stat.st_mtime_ns, **hashes)
                    write_json(receipt_path, stamp)
                progress({"event": "source_downloaded", "file": source["name"],
                          "completed": index + 1, "total": len(manifest["files"])})
            write_json(root / "download-status.json", {"state": "complete", "updated_utc": utc_now(),
                       "files": len(manifest["files"]), "archives": len(ARCHIVE_NAMES),
                       "total_bytes": sum(source["size"] for source in manifest["files"]),
                       "source_manifest_sha256": sha256(root / "sources.json")})
            return manifest
        except Exception as error:
            write_json(root / "download-status.json", {"state": "failed", "updated_utc": utc_now(), "error": str(error)})
            raise


def _subset(member_name):
    parts = [part.lower() for part in member_name.replace("\\", "/").split("/")]
    for name in ("processed", "bad", "duplicates"):
        if name in parts:
            return name
    return "other"


def prepare_runner_archive(root, on_progress=None):
    """Assemble unique original bytes, retaining every category-file occurrence.

    This stage does not open Excel. Names without a final XLS/XLSX suffix receive
    an additional .xls suffix only in the runner ZIP; their original names remain
    in input-manifest.jsonl. Bad and duplicate source folders remain mapped, while
    byte-identical occurrences share one runner input and one Excel preparation.
    """
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    marker = root / "runner-archive.json"
    status_path = root / "archive-status.json"
    archive_path = root / "corpus/euses-original-inputs.zip"
    mapping_path = root / "input-manifest.jsonl"
    with RunLock(root / ".assemble"):
        def progress(event):
            write_json(status_path, {"state": "running", "updated_utc": utc_now(), **event})
            if on_progress:
                on_progress(event)

        try:
            manifest = freeze_sources(root)
            progress({"event": "checking_sources", "completed": 0, "total": len(manifest["files"])})
            sources = []
            for source in manifest["files"]:
                path = root / "raw" / source["name"]
                hashes = _verify(path, source)
                sources.append({"name": source["name"], "sha256": hashes["sha256"],
                                "bytes": path.stat().st_size})
            stamp = {"source_manifest_sha256": sha256(root / "sources.json"),
                     "builder_sha256": sha256(Path(__file__)), "source_files": sources}
            if marker.exists() and archive_path.is_file() and mapping_path.is_file():
                cached = read_json(marker)
                if (cached.get("stamp") == stamp and cached.get("sha256") == sha256(archive_path)
                        and cached.get("mapping_sha256") == sha256(mapping_path)):
                    write_json(status_path, {"state": "complete", "updated_utc": utc_now(), "cached": True,
                                            "workbooks": cached["workbooks"], "path": str(archive_path)})
                    return cached
            archive_path.parent.mkdir(parents=True, exist_ok=True)
            work = root / ".work"
            work.mkdir(exist_ok=True)
            pending_archive = archive_path.with_suffix(".zip.partial")
            pending_mapping = mapping_path.with_suffix(".jsonl.partial")
            paths = _Paths()
            counts = {"processed": 0, "bad": 0, "duplicates": 0, "other": 0, "renamed_for_runner": 0,
                      "unique_inputs": 0, "duplicate_occurrences": 0}
            canonical = {}
            total_files = 0
            with zipfile.ZipFile(pending_archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1,
                                 allowZip64=True) as output, pending_mapping.open("w", encoding="utf-8", newline="\n") as mapping:
                for source in manifest["files"]:
                    if source["name"] not in ARCHIVE_NAMES:
                        continue
                    with zipfile.ZipFile(root / "raw" / source["name"]) as archive:
                        for index, info in enumerate(archive.infolist()):
                            if info.is_dir():
                                continue
                            if info.flag_bits & 1:
                                raise ValueError("Encrypted source ZIP member: " + source["name"] + "/" + info.orig_filename)
                            entry_id = _entry_id(source["name"], (index,))
                            parts = [source["name"], *info.orig_filename.split("/")]
                            suffix = PurePosixPath(parts[-1]).suffix.lower()
                            renamed = suffix not in (".xls", ".xlsx")
                            if renamed:
                                parts[-1] += ".xls"
                            digest = hashlib.sha256()
                            size = 0
                            with tempfile.TemporaryFile(dir=work) as staged:
                                with archive.open(info) as member:
                                    for block in iter(lambda: member.read(1024 * 1024), b""):
                                        digest.update(block)
                                        size += len(block)
                                        staged.write(block)
                                if size != info.file_size:
                                    raise ValueError("A source member has an unexpected size: " + info.orig_filename)
                                content_hash = digest.hexdigest()
                                existing = canonical.get(content_hash)
                                if existing:
                                    counts["duplicate_occurrences"] += 1
                                else:
                                    runner_member = paths.allocate(tuple(parts), entry_id).as_posix()
                                    destination = zipfile.ZipInfo(runner_member, date_time=(1980, 1, 1, 0, 0, 0))
                                    destination.compress_type = zipfile.ZIP_DEFLATED
                                    destination.external_attr = 0o100600 << 16
                                    staged.seek(0)
                                    with output.open(destination, "w", force_zip64=True) as target:
                                        for block in iter(lambda: staged.read(1024 * 1024), b""):
                                            target.write(block)
                                    canonical[content_hash] = {"id": entry_id, "member": runner_member,
                                                               "index": counts["unique_inputs"]}
                                    counts["unique_inputs"] += 1
                            canonical_entry = canonical[content_hash]
                            subset = _subset(info.orig_filename)
                            counts[subset] += 1
                            counts["renamed_for_runner"] += renamed
                            total_files += 1
                            row = {"id": entry_id, "source_archive": source["name"], "source_member": info.orig_filename,
                                   "source_index": index, "runner_member": canonical_entry["member"],
                                   "runner_index": canonical_entry["index"],
                                   "runner_identity": "{}:{}:{}".format(CORPUS_ID, canonical_entry["index"], canonical_entry["member"]),
                                   "sha256": content_hash, "bytes": size, "subset": subset,
                                   "renamed_for_runner": renamed,
                                   "duplicate_of": canonical_entry["id"] if existing else None}
                            mapping.write(json.dumps(row, ensure_ascii=True) + "\n")
                            if total_files % 100 == 0:
                                progress({"event": "assembling_inputs", "completed": total_files,
                                          "source_archive": source["name"], "counts": dict(counts)})
            pending_archive.replace(archive_path)
            pending_mapping.replace(mapping_path)
            result = {"schema_version": 1, "name": "EUSES", "complete": True, "path": str(archive_path),
                      "sha256": sha256(archive_path), "bytes": archive_path.stat().st_size,
                      "workbooks": counts["unique_inputs"], "total_files": total_files, "counts": counts,
                      "corpus_id": CORPUS_ID, "mapping_path": str(mapping_path),
                      "mapping_sha256": sha256(mapping_path), "stamp": stamp,
                      "published_counts": manifest["published_counts"],
                      "scope": "All original category-file occurrences are mapped, including processed, bad and duplicates. Byte-identical originals share one runner input; bytes are unchanged."}
            write_json(marker, result)
            write_json(status_path, {"state": "complete", "updated_utc": utc_now(), "workbooks": counts["unique_inputs"],
                                    "counts": counts, "path": str(archive_path)})
            return result
        except Exception as error:
            write_json(status_path, {"state": "failed", "updated_utc": utc_now(), "error": str(error)})
            raise


def main(argv=None):
    parser = argparse.ArgumentParser(description="Download the original EUSES spreadsheet source files.")
    parser.add_argument("command", choices=("download", "assemble"))
    parser.add_argument("--root", type=Path, default=Path("F:/corpora/euses"))
    args = parser.parse_args(argv)
    try:
        def callback(event):
            if event["event"] in ("source_downloaded", "download_retry"):
                print("{}  {} {}".format(utc_now(), event["event"].replace("_", " "),
                      event.get("file", event.get("archive", ""))), flush=True)
        if args.command == "assemble":
            result = prepare_runner_archive(args.root)
            print("EUSES runner archive ready: {} original inputs.".format(result["workbooks"]), flush=True)
            return 0
        manifest = download_euses(args.root, callback)
        print("EUSES download complete: {} verified source files.".format(len(manifest["files"])), flush=True)
        return 0
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as error:
        print("Error: " + str(error), file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
