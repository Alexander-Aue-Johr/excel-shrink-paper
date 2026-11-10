"""Export original files from the two Enron difference lists, once per SHA-256."""

import argparse
import csv
import ctypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile
import uuid
import zipfile

from .ingest import _Paths
from .common import long_path, read_json, sha256, utc_now, write_json


def _target(destination, relative):
    path = destination / relative
    path.resolve().relative_to(destination)
    return path


def _read_list(path, group):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    unique = {}
    for row in rows:
        digest = row["sha256"]
        if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
            raise ValueError("An input list contains an invalid SHA-256.")
        names, indices = json.loads(row["member_chain"]), json.loads(row["index_chain"])
        if not names or len(names) != len(indices) or any(not isinstance(name, str) for name in names):
            raise ValueError("An input list contains an invalid member chain.")
        if any(type(index) is not int or index < 0 for index in indices) or int(row["size"]) < 0:
            raise ValueError("An input list contains an invalid index or size.")
        if Path(row["archive"]).name != row["archive"] or any(char in row["archive"] for char in "/\\:"):
            raise ValueError("Source archive names must be plain filenames.")
        if digest in unique and int(unique[digest]["size"]) != int(row["size"]):
            raise ValueError("Equal content hashes have different recorded sizes.")
        unique.setdefault(digest, {**row, "group": group, "names": names, "indices": indices})
    return rows, list(unique.values())


def _allocation_unit(path):
    if os.name != "nt":
        return 4096
    sectors, size, free, total = (ctypes.c_ulong() for _ in range(4))
    if not ctypes.windll.kernel32.GetDiskFreeSpaceW(str(path.anchor), ctypes.byref(sectors), ctypes.byref(size),
                                                   ctypes.byref(free), ctypes.byref(total)):
        raise OSError("Could not determine the destination allocation unit.")
    return sectors.value * size.value


def _required_space(destination, files):
    unit = _allocation_unit(destination)
    directories = set()
    allocated = 0
    for row in files:
        target = _target(destination, row["export_path"])
        if not long_path(target).exists():
            allocated += ((int(row["size"]) + unit - 1) // unit) * unit
        for parent in target.parents:
            if parent == destination:
                break
            if not long_path(parent).exists():
                directories.add(parent)
    return allocated + len(directories) * unit + max(16 * 1024 * 1024, 16 * unit)


def _verified_target(target, row):
    if not long_path(target).exists():
        return False
    if not long_path(target).is_file() or long_path(target).stat().st_size != int(row["size"]) or sha256(target) != row["sha256"]:
        raise ValueError("An existing output has different contents; it was retained: " + str(target))
    return True


def _copy_member(archive, info, target, row):
    long_path(target.parent).mkdir(parents=True, exist_ok=True)
    pending = long_path(target.with_name(target.name + "." + uuid.uuid4().hex + ".pending"))
    try:
        digest, size = hashlib.sha256(), 0
        with archive.open(info) as source, pending.open("xb") as output:
            for block in iter(lambda: source.read(1024 * 1024), b""):
                size += len(block)
                if size > int(row["size"]):
                    raise ValueError("A member exceeds the expected original size.")
                digest.update(block)
                output.write(block)
        if size != int(row["size"]) or digest.hexdigest() != row["sha256"]:
            raise ValueError("Original source bytes differ from the comparison list: " + info.orig_filename)
        if sha256(pending) != row["sha256"]:
            raise ValueError("The destination read-back checksum differs: " + str(target))
        # Another writer must never be replaced between the initial check and publication.
        if long_path(target).exists():
            _verified_target(target, row)
        else:
            pending.rename(long_path(target))
    finally:
        pending.unlink(missing_ok=True)


def _tree(files):
    root = {}
    for row in files:
        current = root
        for depth, (index, name) in enumerate(zip(row["indices"], row["names"])):
            node = current.setdefault(index, {"name": name, "children": {}, "row": None})
            if node["name"] != name:
                raise ValueError("Source member names disagree for the same archive index.")
            if depth == len(row["indices"]) - 1:
                if node["children"] or node["row"]:
                    raise ValueError("Source member paths conflict.")
                node["row"] = row
            else:
                if node["row"]:
                    raise ValueError("A source member is listed both as a file and as a container.")
                current = node["children"]
    return root


def _extract_tree(source_path, nodes, destination, work_root, completed):
    with zipfile.ZipFile(long_path(source_path)) as archive:
        entries = archive.infolist()
        for index, node in nodes.items():
            if index >= len(entries) or entries[index].orig_filename != node["name"]:
                raise ValueError("The source archive index or member name has changed.")
            info = entries[index]
            if info.is_dir() or info.flag_bits & 1:
                raise ValueError("A selected source member is a directory or is encrypted.")
            if node["row"]:
                row = node["row"]
                _copy_member(archive, info, _target(destination, row["export_path"]), row)
                completed(row)
            else:
                with tempfile.NamedTemporaryFile(prefix="nested-", suffix=".zip", dir=work_root, delete=False) as output:
                    temporary = Path(output.name)
                    try:
                        with archive.open(info) as member:
                            shutil.copyfileobj(member, output, 1024 * 1024)
                    except Exception:
                        output.close()
                        temporary.unlink(missing_ok=True)
                        raise
                try:
                    if temporary.stat().st_size != info.file_size:
                        raise ValueError("A nested archive has an unexpected size.")
                    _extract_tree(temporary, node["children"], destination, work_root, completed)
                finally:
                    temporary.unlink(missing_ok=True)


def export_differences(root, old_archive, destination, work_root, status_path=None, on_progress=None):
    root, old_archive, destination, work_root = (Path(value).resolve() for value in (root, old_archive, destination, work_root))
    status_path = Path(status_path) if status_path else destination / "export-status.json"
    destination.mkdir(parents=True, exist_ok=True)
    work_root.mkdir(parents=True, exist_ok=True)
    marker = _target(destination, "COPY_COMPLETE.txt")
    marker.unlink(missing_ok=True)
    status = {"state": "running", "files_verified": 0, "bytes_verified": 0, "destination": str(destination)}
    def report():
        status["updated_utc"] = utc_now()
        write_json(status_path, status)
        if on_progress:
            on_progress(dict(status))
    try:
        report()
        source_lists = {name: root / "comparison" / (name + ".csv") for name in ("new-only", "old-not-found")}
        input_hashes = {name: sha256(path) for name, path in source_lists.items()}
        manifest_path = _target(destination, "manifest.json")
        if manifest_path.exists() and read_json(manifest_path).get("source_list_sha256") != input_hashes:
            raise ValueError("The source lists differ from this export; use a separate destination.")
        all_rows, selected, allocations = {}, [], _Paths()
        by_archive = {}
        for group, path in source_lists.items():
            aliases, files = _read_list(path, group)
            all_rows[group] = aliases
            for row in files:
                parts = [group, row["archive"]]
                for depth, name in enumerate(row["names"]):
                    pieces = name.split("/")
                    if depth < len(row["names"]) - 1:
                        pieces[-1] += ".contents"
                    parts.extend(pieces)
                relative = allocations.allocate(tuple(parts), row["sha256"][:24])
                row["export_path"] = relative.as_posix()
                selected.append(row)
        status.update(files_total=len(selected), bytes_total=sum(int(row["size"]) for row in selected),
                      counts={name: sum(row["group"] == name for row in selected) for name in source_lists})
        for row in selected:
            target = _target(destination, row["export_path"])
            if _verified_target(target, row):
                status["files_verified"] += 1
                status["bytes_verified"] += int(row["size"])
            else:
                archive_path = old_archive if row["group"] == "old-not-found" else root / "raw" / row["archive"]
                if row["group"] == "old-not-found" and row["archive"] != old_archive.name:
                    raise ValueError("The old-corpus list names another archive.")
                by_archive.setdefault(archive_path, []).append(row)
        required = _required_space(destination, selected)
        if shutil.disk_usage(destination).free < required:
            raise OSError("The destination needs at least %d free bytes, including allocation padding." % required)
        report()
        def completed(row):
            status["files_verified"] += 1
            status["bytes_verified"] += int(row["size"])
            if status["files_verified"] % 10 == 0:
                report()
        for path, files in by_archive.items():
            _extract_tree(path, _tree(files), destination, work_root, completed)
        if any(sha256(path) != input_hashes[name] for name, path in source_lists.items()):
            raise ValueError("The comparison lists changed during export; no completion marker was written.")
        manifest_files = [{"path": row["export_path"], "sha256": row["sha256"], "bytes": int(row["size"]),
                           "group": row["group"], "archive": row["archive"], "member_chain": row["names"],
                           "index_chain": row["indices"]} for row in selected]
        for group, aliases in all_rows.items():
            paths = {row["sha256"]: row["export_path"] for row in selected if row["group"] == group}
            directory = _target(destination, group)
            long_path(directory).mkdir(parents=True, exist_ok=True)
            pending = long_path(directory / "files.csv.pending")
            with pending.open("w", encoding="utf-8-sig", newline="") as stream:
                fields = list(aliases[0]) + ["export_path"] if aliases else ["sha256", "export_path"]
                writer = csv.DictWriter(stream, fieldnames=fields)
                writer.writeheader()
                writer.writerows({**row, "export_path": paths[row["sha256"]]} for row in aliases)
            pending.replace(long_path(directory / "files.csv"))
        sources = _target(destination, "source-lists")
        sources.mkdir(exist_ok=True)
        for name, path in source_lists.items():
            target = _target(destination, "source-lists/" + name + ".csv")
            shutil.copyfile(path, target)
            if sha256(target) != input_hashes[name]:
                raise ValueError("A copied source list differs from its recorded checksum.")
        summary = root / "comparison/summary.json"
        if summary.exists():
            shutil.copyfile(summary, sources / "summary.json")
        manifest = {"schema_version": 1, "state": "complete", "created_utc": utc_now(),
                    "source_list_sha256": input_hashes, "files": manifest_files,
                    "counts": status["counts"], "total_bytes": status["bytes_total"],
                    "old_not_found_definitive": False}
        write_json(manifest_path, manifest)
        _target(destination, "README.txt").write_text(
            "Original files, one copy per SHA-256. No conversion was performed.\n"
            "new-only contains candidates absent from the old archive.\n"
            "old-not-found contains old files not found in readable scanned sources; absence remains provisional.\n"
            "Each files.csv preserves all source occurrences and their copied paths.\n", encoding="utf-8")
        marker.write_text("All %d original files passed source and destination SHA-256 verification.\n" % len(selected), encoding="utf-8")
        status["state"] = "complete"
        report()
        return manifest
    except Exception as error:
        status.update(state="failed", error=str(error))
        report()
        raise


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("root", "old-archive", "destination", "work-root"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--status", type=Path)
    args = parser.parse_args(argv)
    result = export_differences(args.root, args.old_archive, args.destination, args.work_root, args.status)
    print("Export complete: %d original files." % len(result["files"]), flush=True)


if __name__ == "__main__":
    main()
