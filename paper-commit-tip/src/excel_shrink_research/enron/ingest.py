"""Download the frozen EDRM source set and inventory its spreadsheet attachments."""

import hashlib
import http.client
import json
import os
from pathlib import Path
import re
import shutil
import struct
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import zlib


ITEM = "edrm.enron.email.data.set.v2.xml"
METADATA_URL = "https://archive.org/metadata/" + ITEM
EXPECTED_ARCHIVES = 159
CHUNK_SIZE = 1024 * 1024
MAX_NESTING = 3
MAX_NESTED_BYTES = 2 * 1024 * 1024 * 1024
SPREADSHEET_SUFFIXES = {
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".xla",
    ".xlam", ".xlw", ".csv", ".tsv", ".slk", ".sylk", ".dif", ".wks",
    ".wk1", ".wk3", ".wk4", ".123", ".wq1", ".wb1", ".wb2", ".wb3", ".qpw",
}


class SourceError(RuntimeError):
    """The source set is incomplete or failed an integrity check."""


def _progress(callback, **event):
    if callback:
        callback(event)


def _write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    temporary.replace(path)


def _fingerprint(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True).encode("utf-8")).hexdigest()


def _long_path(path):
    path = Path(path).absolute()
    text = str(path)
    if os.name == "nt" and not text.startswith("\\\\?\\"):
        return Path("\\\\?\\UNC\\" + text[2:] if text.startswith("\\\\") else "\\\\?\\" + text)
    return path


def _request(url, headers=None):
    return urllib.request.urlopen(
        urllib.request.Request(url, headers={"User-Agent": "excel-shrink-research/1", **(headers or {})}),
        timeout=90,
    )


def freeze_sources(root):
    """Reuse sources.json, or pin the 159 ZIP files in Archive.org metadata."""
    root = Path(root)
    path = root / "sources.json"
    if path.exists():
        manifest = json.loads(path.read_text(encoding="utf-8"))
        _validate_manifest(manifest)
        return manifest
    with _request(METADATA_URL) as response:
        raw = response.read()
    metadata = json.loads(raw)
    archives = []
    for item in metadata["files"]:
        if not item["name"].lower().endswith(".zip"):
            continue
        archives.append({
            "name": item["name"], "size": int(item["size"]),
            "md5": item.get("md5"), "sha1": item.get("sha1"),
            "url": "https://archive.org/download/" + ITEM + "/" + urllib.parse.quote(item["name"], safe=""),
        })
    manifest = {
        "schema": 1, "item": ITEM, "metadata_url": METADATA_URL,
        "metadata_sha256": hashlib.sha256(raw).hexdigest(),
        "archives": sorted(archives, key=lambda entry: entry["name"]),
    }
    _validate_manifest(manifest)
    root.mkdir(parents=True, exist_ok=True)
    metadata_temporary = root / "source-metadata.json.tmp"
    metadata_temporary.write_bytes(raw)
    metadata_temporary.replace(root / "source-metadata.json")
    _write_json(path, manifest)
    return manifest


def _validate_manifest(manifest):
    archives = manifest.get("archives", [])
    if manifest.get("schema") != 1 or manifest.get("item") != ITEM:
        raise SourceError("The source manifest does not describe the expected EDRM collection.")
    if len(archives) != EXPECTED_ARCHIVES:
        raise SourceError("Expected 159 source ZIPs; found %d. Review the source manifest." % len(archives))
    names = set()
    for source in archives:
        name = source["name"]
        if name.casefold() in names or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*\.zip", name, re.I):
            raise SourceError("Duplicate or unsupported source archive name: " + name)
        names.add(name.casefold())
        if source["size"] <= 0 or not source.get("md5") or not source.get("sha1"):
            raise SourceError("Missing size or checksums for " + name)
        if not re.fullmatch(r"[a-fA-F0-9]{32}", source["md5"]) or not re.fullmatch(r"[a-fA-F0-9]{40}", source["sha1"]):
            raise SourceError("Invalid source checksum for " + name)


def _hash_file(path):
    hashes = {name: hashlib.new(name) for name in ("md5", "sha1", "sha256")}
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(CHUNK_SIZE), b""):
            for digest in hashes.values():
                digest.update(chunk)
    return {name: digest.hexdigest() for name, digest in hashes.items()}


def _verify(path, source):
    if path.stat().st_size != source["size"]:
        raise SourceError("Incorrect download size for " + source["name"])
    hashes = _hash_file(path)
    for algorithm in ("md5", "sha1"):
        if source.get(algorithm) and hashes[algorithm] != source[algorithm].lower():
            raise SourceError("%s checksum mismatch for %s" % (algorithm.upper(), source["name"]))
    return hashes


def _download(source, destination, callback, retries=4):
    partial = destination.with_name(destination.name + ".partial")
    destination.parent.mkdir(parents=True, exist_ok=True)
    last_error = None
    for attempt in range(retries):
        try:
            offset = partial.stat().st_size if partial.exists() else 0
            if offset > source["size"]:
                partial.unlink()
                offset = 0
            if offset < source["size"]:
                headers = {"Range": "bytes=%d-" % offset} if offset else {}
                with _request(source["url"], headers) as response:
                    status = getattr(response, "status", 200)
                    if status == 206:
                        expected = "bytes %d-" % offset
                        if not response.headers.get("Content-Range", "").startswith(expected):
                            raise SourceError("Server returned an unexpected byte range for " + source["name"])
                    elif status == 200:
                        offset = 0
                    else:
                        raise SourceError("Unexpected HTTP status %s for %s" % (status, source["name"]))
                    last_update = 0.0
                    with partial.open("ab" if offset else "wb") as output:
                        for chunk in iter(lambda: response.read(CHUNK_SIZE), b""):
                            output.write(chunk)
                            offset += len(chunk)
                            if offset > source["size"]:
                                raise SourceError("Download exceeds the recorded size for " + source["name"])
                            if time.monotonic() - last_update >= 5:
                                _progress(callback, event="download", archive=source["name"], completed=offset, total=source["size"])
                                last_update = time.monotonic()
            try:
                hashes = _verify(partial, source)
            except SourceError:
                if partial.exists() and partial.stat().st_size >= source["size"]:
                    partial.unlink()
                raise
            partial.replace(destination)
            return hashes
        except (OSError, SourceError, urllib.error.URLError, http.client.HTTPException) as error:
            last_error = error
            _progress(callback, event="download_retry", archive=source["name"], attempt=attempt + 1, message=str(error))
            if attempt + 1 < retries:
                time.sleep(min(30, 2 ** attempt))
    raise SourceError("Download failed after %d attempts: %s (%s)" % (retries, source["name"], last_error))


def download_sources(root, manifest, on_progress=None):
    """Resume interrupted downloads and verify their recorded MD5 and SHA-1."""
    root = Path(root)
    for index, source in enumerate(manifest["archives"]):
        destination = root / "raw" / source["name"]
        receipt_path = root / "raw" / ".verified" / (source["name"] + ".json")
        receipt = json.loads(receipt_path.read_text(encoding="utf-8")) if receipt_path.exists() else {}
        stamp = {"source": _fingerprint(source)}
        if destination.exists():
            stat = destination.stat()
            stamp.update(size=stat.st_size, mtime_ns=stat.st_mtime_ns)
        if not destination.exists() or any(receipt.get(key) != value for key, value in stamp.items()):
            if destination.exists():
                hashes = _verify(destination, source)
            else:
                hashes = _download(source, destination, on_progress)
            stat = destination.stat()
            stamp.update(size=stat.st_size, mtime_ns=stat.st_mtime_ns, **hashes)
            _write_json(receipt_path, stamp)
        _progress(on_progress, event="archive_downloaded", archive=source["name"], completed=index + 1, total=len(manifest["archives"]))


def _escape_component(component):
    if not component:
        return "%00"
    escaped = ""
    for index, char in enumerate(component):
        unsafe = ord(char) < 32 or char in '%<>:"/\\|?*' or (index >= len(component.rstrip(" .")) and char in " .")
        escaped += "%%%02X" % ord(char) if unsafe else char
    if re.match(r"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)", escaped, re.I):
        escaped = "%%%02X" % ord(escaped[0]) + escaped[1:]
    if len(escaped.encode("utf-16-le")) > 320:
        shortened = escaped[:120]
        while len(shortened.encode("utf-16-le")) > 240:
            shortened = shortened[:-1]
        escaped = shortened + "~" + hashlib.sha256(component.encode("utf-8")).hexdigest()[:20]
    return escaped


class _Paths:
    """Keep a readable tree; inventory paths record every escaped name and collision."""

    def __init__(self):
        self.files = set()
        self.directories = set()

    def allocate(self, parts, entry_id):
        output = []
        for part in parts[:-1]:
            name = _escape_component(part)
            key = "/".join(output + [name]).casefold()
            if key in self.files:
                name += "~" + entry_id
                while "/".join(output + [name]).casefold() in self.files:
                    name += "~"
            output.append(name)
            self.directories.add("/".join(output).casefold())
        name = _escape_component(parts[-1])
        key = "/".join(output + [name]).casefold()
        if key in self.files or key in self.directories:
            suffix = Path(name).suffix
            stem = name[:-len(suffix)] if suffix else name
            name = stem + "~" + entry_id + suffix
            number = 1
            while "/".join(output + [name]).casefold() in self.files | self.directories:
                name = stem + "~" + entry_id + "-%d" % number + suffix
                number += 1
        output.append(name)
        self.files.add("/".join(output).casefold())
        return Path(*output)


def _ole_names(path):
    """Read the CFB directory, without loading workbook streams or invoking Office."""
    with path.open("rb") as stream:
        header = stream.read(512)
        if len(header) != 512:
            raise ValueError("Truncated OLE header")
        shift = struct.unpack_from("<H", header, 30)[0]
        if shift not in (9, 12):
            raise ValueError("Unsupported OLE sector size")
        sector_size = 1 << shift
        sector_count = max(0, path.stat().st_size // sector_size - 1)

        def sector(number):
            if number >= sector_count:
                raise ValueError("OLE sector outside the file")
            stream.seek((number + 1) * sector_size)
            return stream.read(sector_size)

        fat_ids = list(struct.unpack_from("<109I", header, 76))
        difat = struct.unpack_from("<I", header, 68)[0]
        visited = set()
        while difat < 0xFFFFFFFA:
            if difat in visited:
                raise ValueError("Cyclic OLE allocation table")
            visited.add(difat)
            values = struct.unpack("<%dI" % (sector_size // 4), sector(difat))
            fat_ids.extend(values[:-1])
            difat = values[-1]
        fat = []
        for number in fat_ids:
            if number < 0xFFFFFFFA:
                fat.extend(struct.unpack("<%dI" % (sector_size // 4), sector(number)))
        directory = struct.unpack_from("<I", header, 48)[0]
        names = set()
        visited.clear()
        while directory < 0xFFFFFFFA:
            if directory in visited or directory >= len(fat):
                raise ValueError("Invalid OLE directory chain")
            visited.add(directory)
            data = sector(directory)
            for offset in range(0, sector_size, 128):
                length = struct.unpack_from("<H", data, offset + 64)[0]
                if 2 <= length <= 64 and data[offset + 66] in (1, 2, 5):
                    names.add(data[offset:offset + length - 2].decode("utf-16-le", errors="replace").casefold())
            directory = fat[directory]
        return names


def _detect(prefix, name, path=None):
    suffix = Path(name).suffix.lower()
    if prefix.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        if path is None:
            return "ole", "ambiguous", "Inspect the OLE directory."
        try:
            names = _ole_names(path)
        except (ValueError, OSError, struct.error) as error:
            return "ole", "ambiguous", "OLE directory could not be read: " + str(error)
        if "workbook" in names or "book" in names:
            return "xls", "candidate", "OLE Workbook or Book stream."
        if "encryptedpackage" in names:
            return "encrypted-office", "ambiguous", "Encrypted Office package; its document type is unknown."
        if "worddocument" in names or "powerpoint document" in names:
            return "other-office", "other", "OLE Word or PowerPoint document."
        return "ole", "ambiguous", "OLE container without an identified document stream."
    if prefix.startswith((b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")):
        if path is None:
            return "zip", "ambiguous", "Inspect the ZIP package."
        try:
            with zipfile.ZipFile(path) as package:
                entries = package.infolist()
                if any(entry.flag_bits & 1 for entry in entries):
                    return "encrypted-archive", "ambiguous", "Encrypted ZIP members; the original archive is retained for review."
                names = {entry.filename.replace("\\", "/").lower() for entry in entries}
                if "xl/workbook.bin" in names:
                    return "xlsb", "candidate", "Office binary workbook package."
                if "xl/workbook.xml" in names:
                    kind = "xlsm" if "xl/vbaproject.bin" in names else "xlsx"
                    return kind, "candidate", "Office workbook package."
                if "word/document.xml" in names or "ppt/presentation.xml" in names:
                    return "other-office", "other", "Office Word or PowerPoint package."
                if "mimetype" in names:
                    entry = next(entry for entry in entries if entry.filename.lower() == "mimetype")
                    if entry.file_size < 256 and b"opendocument.spreadsheet" in package.read(entry):
                        return "ods", "candidate", "OpenDocument spreadsheet package."
        except (OSError, zipfile.BadZipFile, RuntimeError) as error:
            return "zip", "ambiguous", "ZIP package could not be read: " + str(error)
        return "zip", "container", "Nested ZIP archive."
    if prefix[:2] in (b"\x09\x00", b"\x09\x02", b"\x09\x04", b"\x09\x08") and len(prefix) >= 4:
        return "biff", "candidate", "BIFF beginning-of-file record."
    if prefix[:4] in (b"\x00\x00\x02\x00", b"\x00\x00\x1a\x00"):
        return "lotus", "candidate", "Lotus workbook header."
    if prefix.startswith((b"%PDF", b"\x89PNG", b"GIF87a", b"GIF89a", b"\xff\xd8\xff", b"MZ", b"BM", b"{\\rtf")):
        return "other", "other", "Recognized non-spreadsheet signature."
    if prefix.startswith((b"Rar!", b"7z\xbc\xaf\x27\x1c", b"\x1f\x8b", b"BZh")):
        return "archive", "ambiguous", "Nested archive format requires a separate extraction step."
    if prefix.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = prefix.decode("utf-16", errors="replace")
    else:
        text = prefix.decode("utf-8-sig", errors="replace")
    start = text.lstrip()
    lower = start.lower()
    if "urn:schemas-microsoft-com:office:spreadsheet" in lower:
        return "spreadsheetml", "candidate", "Excel SpreadsheetML namespace."
    if start.startswith("ID;"):
        return "sylk", "candidate", "SYLK header."
    if start.startswith("TABLE\n") or start.startswith("TABLE\r\n"):
        return "dif", "candidate", "Data Interchange Format header."
    if "urn:schemas-microsoft-com:office:excel" in lower:
        return "html", "candidate", "Excel HTML namespace."
    if "<table" in lower and not lower.startswith("<?xml"):
        return "html", "ambiguous", "HTML table; spreadsheet intent needs review."
    if suffix in SPREADSHEET_SUFFIXES:
        return "text" if "\x00" not in text else "unknown", "ambiguous", "Spreadsheet suffix without a recognized workbook signature."
    lines = [line for line in text.splitlines()[:20] if line.strip()]
    if len(lines) >= 3 and all("\t" in line for line in lines) and len({line.count("\t") for line in lines}) == 1:
        return "tsv", "ambiguous", "Consistent tab-separated rows; spreadsheet intent needs review."
    if lower.startswith(("<?xml", "<", "from:", "received:", "mime-version:")) or (text and "\x00" not in text and text.count("\ufffd") < len(text) // 100 + 1):
        return "text", "other", "No spreadsheet signature in the first 64 KiB."
    return "unknown", "ambiguous", "Unrecognized binary content."


def _entry_id(archive, chain):
    return hashlib.sha256((archive + "\0" + "/".join(map(str, chain))).encode("utf-8")).hexdigest()[:24]


def _inventory_archive(root, source, output, callback):
    paths = _Paths()
    work = root / ".work"
    work.mkdir(parents=True, exist_ok=True)
    counts = {"members": 0, "candidate": 0, "ambiguous": 0, "read_errors": 0}

    def retain_staged(temporary, row, parts):
        if temporary is None:
            raise SourceError("Missing staged data for " + row["member"])
        relative = Path("extracted") / _escape_component(source["name"]) / paths.allocate(parts, row["id"])
        destination = root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary.replace(destination)
        row["extracted_path"] = relative.as_posix()
        counts[row["classification"]] += 1

    def scan(path, parent_parts=(), chain=(), member_chain=(), depth=0, sink=output):
        with zipfile.ZipFile(path) as archive:
            for index, info in enumerate(archive.infolist()):
                member_name = info.orig_filename
                indices = chain + (index,)
                entry_id = _entry_id(source["name"], indices)
                original_parts = parent_parts + tuple(member_name.split("/"))
                row = {
                    "id": entry_id, "archive": source["name"], "member": member_name,
                    "index": index, "member_chain": list(member_chain + (member_name,)),
                    "index_chain": list(indices), "size": info.file_size, "sha256": None,
                    "kind": "directory", "classification": "directory", "extracted_path": None,
                    "attempt_excel": False,
                }
                counts["members"] += 1
                if info.is_dir():
                    sink.write(json.dumps(row, ensure_ascii=True) + "\n")
                    continue
                if info.flag_bits & 1:
                    raise SourceError("Encrypted ZIP member cannot be inventoried: %s / %s" % (source["name"], member_name))
                symlink = (info.external_attr >> 16) & 0o170000 == 0o120000
                temporary = None
                try:
                    with archive.open(info) as member:
                        prefix = member.read(65536)
                        kind, classification, evidence = _detect(prefix, member_name)
                        retain = classification != "other" or symlink
                        digest = hashlib.sha256()
                        total = 0
                        if retain:
                            handle = tempfile.NamedTemporaryFile(prefix="member-", suffix=".bin", dir=work, delete=False)
                            temporary = Path(handle.name)
                        else:
                            handle = None
                        try:
                            chunk = prefix
                            while chunk:
                                digest.update(chunk)
                                total += len(chunk)
                                if handle:
                                    handle.write(chunk)
                                chunk = member.read(CHUNK_SIZE)
                        finally:
                            if handle:
                                handle.close()
                    if total != info.file_size:
                        raise SourceError("Incorrect extracted size for " + member_name)
                    if temporary and kind in ("ole", "zip"):
                        kind, classification, evidence = _detect(prefix, member_name, temporary)
                    if symlink:
                        kind, classification, evidence = "symlink", "ambiguous", "ZIP symlink stored as regular bytes; it was not followed."
                    row.update(sha256=digest.hexdigest(), kind=kind, classification=classification, evidence=evidence)
                    row["attempt_excel"] = not symlink and (
                        classification == "candidate"
                        or kind in ("ole", "encrypted-office", "tsv")
                        or (classification == "ambiguous" and kind not in ("encrypted-archive", "archive") and Path(member_name).suffix.lower() in SPREADSHEET_SUFFIXES)
                    )
                    if classification == "container":
                        if depth >= MAX_NESTING or info.file_size > MAX_NESTED_BYTES:
                            row.update(classification="ambiguous", evidence="Nested ZIP exceeds the configured depth or size limit.")
                        else:
                            # Spool child rows so the container can report whether
                            # every readable member was inspected, without keeping
                            # a large nested inventory in memory.
                            with tempfile.TemporaryFile(mode="w+", encoding="utf-8", newline="\n", dir=work) as nested_output:
                                before = counts["read_errors"]
                                scan(temporary, original_parts, indices, member_chain + (member_name,), depth + 1, nested_output)
                                errors = counts["read_errors"] - before
                                if errors:
                                    row.update(classification="ambiguous", nested_status="partial", nested_read_errors=errors,
                                               evidence="Some nested ZIP members could not be read; the original archive is retained for review.")
                                    retain_staged(temporary, row, original_parts)
                                    temporary = None
                                else:
                                    row.update(classification="other", nested_status="scanned")
                                sink.write(json.dumps(row, ensure_ascii=True) + "\n")
                                nested_output.seek(0)
                                shutil.copyfileobj(nested_output, sink)
                            continue
                    if row["classification"] in ("candidate", "ambiguous"):
                        retain_staged(temporary, row, original_parts)
                        temporary = None
                    sink.write(json.dumps(row, ensure_ascii=True) + "\n")
                    if counts["members"] % 250 == 0:
                        _progress(callback, event="inventory", archive=source["name"], **counts)
                except (zipfile.BadZipFile, RuntimeError, EOFError, zlib.error) as error:
                    if depth == 0 or isinstance(error, SourceError):
                        raise SourceError("Cannot inventory %s / %s: %s" % (source["name"], member_name, error)) from error
                    # A downloaded source ZIP must pass integrity checks. Damage
                    # inside an attached ZIP is a documented source exception;
                    # retain its enclosing bytes and continue readable siblings.
                    row.update(kind="unreadable", classification="unreadable", sha256=None,
                               evidence="ZIP member could not be read: " + str(error),
                               read_error={"type": type(error).__name__, "message": str(error)})
                    counts["read_errors"] += 1
                    sink.write(json.dumps(row, ensure_ascii=True) + "\n")
                    _progress(callback, event="inventory_read_error", archive=source["name"],
                              member_chain=row["member_chain"], message=str(error), **counts)
                except OSError as error:
                    raise SourceError("Cannot inventory %s / %s: %s" % (source["name"], member_name, error)) from error
                finally:
                    if temporary and temporary.exists():
                        temporary.unlink()

    scan(root / "raw" / source["name"])
    return counts


def _checkpoint_valid(root, inventory, marker, stamp):
    if not inventory.exists() or not marker.exists():
        return False
    try:
        if json.loads(marker.read_text(encoding="utf-8"))["source"] != stamp:
            return False
        with inventory.open(encoding="utf-8") as stream:
            for line in stream:
                row = json.loads(line)
                if row.get("extracted_path"):
                    path = root / row["extracted_path"]
                    if not path.is_file() or path.stat().st_size != row["size"]:
                        return False
    except (OSError, ValueError, KeyError):
        return False
    return True


def inventory_sources(root, manifest, on_progress=None):
    """Stream every ZIP member, check CRCs, and retain candidate and ambiguous files.

    The returned JSONL lists every occurrence. Equal candidate bytes retain their
    source paths and receive duplicate_of, referring to the first occurrence ID.
    Unknown binary files and unsupported nested archives remain available for review.
    """
    root = _long_path(root)
    directory = root / "inventory"
    directory.mkdir(parents=True, exist_ok=True)
    inventories = []
    for index, source in enumerate(manifest["archives"]):
        raw = root / "raw" / source["name"]
        if not raw.is_file() or raw.stat().st_size != source["size"]:
            raise SourceError("Source archive is missing or incomplete: " + source["name"])
        stamp = {"manifest": _fingerprint(source), "size": raw.stat().st_size, "mtime_ns": raw.stat().st_mtime_ns, "inventory_schema": 1}
        inventory = directory / (source["name"] + ".jsonl")
        marker = directory / (source["name"] + ".complete.json")
        if not _checkpoint_valid(root, inventory, marker, stamp):
            temporary = inventory.with_name(inventory.name + ".partial")
            with temporary.open("w", encoding="utf-8", newline="\n") as output:
                counts = _inventory_archive(root, source, output, on_progress)
            temporary.replace(inventory)
            _write_json(marker, {"source": stamp, "counts": counts,
                                 "scan_status": "completed_with_read_errors" if counts["read_errors"] else "completed"})
        inventories.append(inventory)
        _progress(on_progress, event="archive_inventoried", archive=source["name"], completed=index + 1, total=len(manifest["archives"]))
    combined = root / "inventory.jsonl"
    temporary = combined.with_name(combined.name + ".partial")
    seen = {}
    with temporary.open("w", encoding="utf-8", newline="\n") as output:
        for inventory in inventories:
            with inventory.open(encoding="utf-8") as stream:
                for line in stream:
                    row = json.loads(line)
                    if row["classification"] in ("candidate", "ambiguous"):
                        previous = seen.setdefault(row["sha256"], row["id"])
                        if previous != row["id"]:
                            row["duplicate_of"] = previous
                    output.write(json.dumps(row, ensure_ascii=True) + "\n")
    temporary.replace(combined)
    return combined
