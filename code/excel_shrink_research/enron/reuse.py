"""Import complete study preparations matched by the original source hash."""

import hashlib
import os
from pathlib import Path, PureWindowsPath
import shutil
import zipfile
import zlib

from .common import contained, long_path, read_json, sha256, write_json


PROFILES = ("default", "maximum-cleaning", "too-aggressive-cleaning")
REQUIRED_FILES = ("original.xlsx",) + tuple(name + ".xlsx" for name in PROFILES)
OPTIONAL_FILES = ("converted.xlsx", "preparation.json")
MAX_WORKBOOK_BYTES = 16 * 1024 ** 3


def _digest(value):
    if not isinstance(value, str) or len(value) != 64:
        raise ValueError("A reuse checksum is missing or invalid.")
    if any(character not in "0123456789abcdef" for character in value):
        raise ValueError("Reuse checksums must be lowercase SHA-256 values.")
    return value


def _relative(value):
    if not isinstance(value, (str, Path)) or not str(value):
        raise ValueError("A reuse path is missing.")
    path = Path(value)
    windows = PureWindowsPath(str(value))
    if path.is_absolute() or windows.drive or windows.root or ".." in windows.parts:
        raise ValueError("A reuse path must stay within its output directory.")
    return path


def _candidates(cache_roots, source_sha256):
    candidates = []
    for cache_root in cache_roots:
        root = Path(cache_root).resolve()
        directory = root / "entries" / source_sha256
        if directory.is_dir():
            candidates.extend((root, path) for path in sorted(directory.glob("*.json")))
    return sorted(candidates, key=lambda item: (item[1].stem, str(item[0])))


def _file_metadata(entry):
    files = {}
    for item in entry["files"]:
        name = item["name"]
        if name in files or name not in REQUIRED_FILES + OPTIONAL_FILES:
            raise ValueError("The reuse manifest contains duplicate or unexpected workbook names.")
        _digest(item["sha256"])
        size = item.get("bytes", item.get("size"))
        if isinstance(size, bool) or not isinstance(size, int) or not 0 <= size <= MAX_WORKBOOK_BYTES:
            raise ValueError("The reuse manifest has an invalid workbook size.")
        files[name] = {"name": name, "sha256": item["sha256"], "bytes": size}
    if not all(name in files for name in REQUIRED_FILES):
        raise ValueError("The preparation does not contain the baseline and all three profiles.")
    return files


def _check_package(path):
    with zipfile.ZipFile(long_path(path)) as package:
        members = package.infolist()
        names = [item.filename for item in members]
        if len(names) != len(set(names)):
            raise ValueError("A reused XLSX package contains duplicate ZIP members.")
        if sum(item.file_size for item in members) > MAX_WORKBOOK_BYTES:
            raise ValueError("A reused XLSX package exceeds the size limit.")
        content_types = package.getinfo("[Content_Types].xml")
        if content_types.file_size > 8 * 1024 * 1024:
            raise ValueError("A reused XLSX package has an oversized content-type declaration.")
        if ("xl/workbook.xml" not in names or
                b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                not in package.read(content_types)):
            raise ValueError("A reused file is not an XLSX workbook.")
        invalid = package.testzip()
        if invalid:
            raise ValueError("A reused XLSX package has an invalid CRC: " + invalid)


def _unpack(bundle, files, staging, on_progress):
    with zipfile.ZipFile(long_path(bundle)) as archive:
        names = [item.filename for item in archive.infolist()]
        for name, metadata in files.items():
            if names.count(name) != 1:
                raise ValueError("A required workbook is missing or duplicated: " + name)
            member = archive.getinfo(name)
            if member.file_size != metadata["bytes"] or member.flag_bits & 1:
                raise ValueError("A reused workbook has an unexpected size or is encrypted: " + name)
            destination = staging / name
            digest = hashlib.sha256()
            copied = 0
            with archive.open(member) as source, long_path(destination).open("wb") as target:
                for block in iter(lambda: source.read(1024 * 1024), b""):
                    copied += len(block)
                    if copied > metadata["bytes"]:
                        raise ValueError("A reused workbook exceeds its recorded size.")
                    digest.update(block)
                    target.write(block)
            if copied != metadata["bytes"] or digest.hexdigest() != metadata["sha256"]:
                raise ValueError("A reused workbook differs from its recorded checksum: " + name)
            if name.endswith(".xlsx"):
                _check_package(destination)
            if on_progress:
                on_progress({"event": "reuse", "message": "Checking saved study workbooks", "workbook": name})


def _install(source, destination, expected_sha256):
    destination = Path(destination)
    long_path(destination.parent).mkdir(parents=True, exist_ok=True)
    if long_path(destination).exists():
        if sha256(destination) != expected_sha256:
            raise ValueError("An existing reuse output has different contents: " + str(destination))
        return
    pending = destination.with_name(destination.name + ".pending")
    shutil.copyfile(long_path(source), long_path(pending))
    os.replace(long_path(pending), long_path(destination))


def reuse_file(source_sha256, destination, variants_root, attempt_dir, cache_roots,
               *, relative_path, on_progress=None):
    """Return None on a cache miss; otherwise import or report a review issue.

    Imports use only exact original-source SHA-256 matches. The historical
    baseline is copied unchanged because the saved profiles were made from it.
    Package checks do not claim a fresh Excel reopen or conversion fidelity.
    """
    source_sha256 = _digest(source_sha256)
    relative = _relative(relative_path)
    candidates = _candidates(cache_roots, source_sha256)
    if not candidates:
        return None
    cache_root, entry_path = candidates[0]
    attempt_dir = Path(attempt_dir)
    staging = attempt_dir / "reuse"
    long_path(staging).mkdir(parents=True, exist_ok=True)
    result = {
        "status": "review", "reason": "reuse_invalid", "source_sha256": source_sha256,
        "method": "reuse_study_preparation", "cache_entry": str(entry_path),
        "prompt_check": "historical_study_baseline_not_freshly_reopened",
    }
    try:
        entry = read_json(entry_path)
        if entry.get("schema_version") != 1 or entry.get("source_sha256") != source_sha256:
            raise ValueError("The reuse entry does not match this original source.")
        bundle_sha256 = _digest(entry["bundle_sha256"])
        if entry_path.stem != bundle_sha256:
            raise ValueError("The reuse entry filename and bundle checksum disagree.")
        record = entry["record"]
        if not isinstance(record, dict) or record.get("source_sha256") != source_sha256:
            raise ValueError("The historical study record does not match this original source.")
        if record.get("prepared") is not True:
            raise ValueError("The historical study record is not a complete preparation.")
        files = _file_metadata(entry)
        baseline_sha256 = record.get("baseline_sha256")
        if baseline_sha256 is not None and baseline_sha256 != files["original.xlsx"]["sha256"]:
            raise ValueError("The historical baseline checksum does not match original.xlsx.")
        baseline_repaired = record.get("baseline_repaired", entry.get("baseline_repaired"))
        if baseline_repaired is not None and not isinstance(baseline_repaired, bool):
            raise ValueError("The historical baseline repair state is invalid.")
        bundle = contained(cache_root, _relative(entry["bundle_path"]))
        if sha256(bundle) != bundle_sha256:
            raise ValueError("The preparation bundle differs from its recorded checksum.")
        _unpack(bundle, files, staging, on_progress)
        outputs = {"original.xlsx": Path(destination)}
        outputs.update({profile + ".xlsx": contained(Path(variants_root) / profile, relative)
                        for profile in PROFILES})
        if "converted.xlsx" in files:
            outputs["converted.xlsx"] = contained(Path(variants_root) / "conversion", relative)
        # Check all existing destinations before installing any new output.
        for name, target in outputs.items():
            if long_path(target).exists() and sha256(target) != files[name]["sha256"]:
                raise ValueError("An existing reuse output has different contents: " + str(target))
        for name, target in outputs.items():
            _install(staging / name, target, files[name]["sha256"])
        profile_records = {item["profile"]: item for item in record.get("results", [])
                           if isinstance(item, dict) and "profile" in item}
        variants = {
            profile: {"path": str(outputs[profile + ".xlsx"]),
                      "sha256": files[profile + ".xlsx"]["sha256"],
                      "bytes": files[profile + ".xlsx"]["bytes"],
                      "config_hash": profile_records.get(profile, {}).get("config_hash")}
            for profile in PROFILES
        }
        result.update(
            status="accepted", reason="exact_source_match", stage="complete",
            output_sha256=files["original.xlsx"]["sha256"],
            output_bytes=files["original.xlsx"]["bytes"], variants=variants,
            baseline_repaired=baseline_repaired,
            validation="historical_study_baseline_package_checked",
            reuse={**entry, "cache_root": str(cache_root), "cache_entry": str(entry_path),
                   "candidate_count": len(candidates),
                   "alternative_bundle_sha256s": sorted({path.stem for _, path in candidates
                                                          if path.stem != bundle_sha256})},
        )
        if "converted.xlsx" in files:
            result["conversion_artifact"] = {
                "path": str(outputs["converted.xlsx"]),
                "sha256": files["converted.xlsx"]["sha256"],
                "bytes": files["converted.xlsx"]["bytes"],
            }
        return result
    except (OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile, RuntimeError, zlib.error) as error:
        result.update(stage="reuse", error=str(error))
        return result
    finally:
        write_json(attempt_dir / "reuse.json", result)
        for name in REQUIRED_FILES + OPTIONAL_FILES:
            long_path(staging / name).unlink(missing_ok=True)
