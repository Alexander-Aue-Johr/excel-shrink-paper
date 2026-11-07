import json
import os
import zipfile
from collections import Counter
from pathlib import Path

from .common import RunLock, contained, long_path, read_json, rows, sha256, utc_now, write_json


def accepted_outputs(root):
    root = Path(root).resolve()
    summary = read_json(root / "summary.json")
    if summary["limited"]:
        raise ValueError("Finish preparation before packaging or benchmarking.")
    for name, key in (
        ("sources.json", "source_manifest_sha256"),
        ("inventory.jsonl", "inventory_sha256"),
        ("conversions.jsonl", "conversion_manifest_sha256"),
    ):
        if summary.get(key) != sha256(root / name):
            raise ValueError("Preparation metadata changed or is incomplete: " + name)
    candidates = {}
    inventory_entries = 0
    for source in rows(root / "inventory.jsonl"):
        inventory_entries += 1
        if source["classification"] in ("candidate", "ambiguous"):
            if source["id"] in candidates:
                raise ValueError("The inventory repeats an occurrence ID: " + source["id"])
            candidates[source["id"]] = source
    outputs = {}
    seen = set()
    counts = Counter()
    accepted_contents = {}
    for row in rows(root / "conversions.jsonl"):
        source = row["source"]
        identifier = source["id"]
        if identifier in seen or identifier not in candidates:
            raise ValueError("The conversion manifest repeats or adds an occurrence: " + identifier)
        if source != candidates[identifier]:
            raise ValueError("The conversion source differs from its inventory entry: " + identifier)
        seen.add(identifier)
        result = row["result"]
        if result["status"] not in ("accepted", "review"):
            raise ValueError("An occurrence has no final conversion outcome: " + identifier)
        counts[result["status"]] += 1
        if result["status"] == "accepted":
            relative, digest = result["output_path"], result["output_sha256"]
            contained(root, relative).relative_to(root / "xlsx")
            if relative in outputs and outputs[relative] != digest:
                raise ValueError("An output has conflicting checksums: " + relative)
            previous = accepted_contents.setdefault(source["sha256"], (relative, digest))
            if previous != (relative, digest):
                raise ValueError("Duplicate input contents map to different outputs: " + identifier)
            outputs[relative] = digest
            for artifact in list(result.get("variants", {}).values()) + ([result["conversion_artifact"]] if result.get("conversion_artifact") else []):
                saved = contained(root, artifact["path"])
                saved.relative_to(root / "variants")
                if sha256(saved) != artifact["sha256"]:
                    raise ValueError("A saved study variant changed: " + artifact["path"])
    if seen != set(candidates):
        raise ValueError("The conversion manifest is missing inventory occurrences.")
    expected_counts = {
        "inventory_entries": inventory_entries,
        "candidate_occurrences": len(candidates),
        "accepted_occurrences": counts["accepted"],
        "review_occurrences": counts["review"],
    }
    if any(summary.get(key, 0) != count for key, count in expected_counts.items()):
        raise ValueError("The occurrence counts and preparation summary disagree.")
    if len(outputs) != summary["unique_accepted"]:
        raise ValueError("The conversion manifest and summary disagree.")
    if len(accepted_contents) != len(outputs):
        raise ValueError("Different input contents share an output path.")
    if not outputs:
        raise ValueError("There are no accepted XLSX files.")
    for relative, digest in outputs.items():
        if sha256(contained(root, relative)) != digest:
            raise ValueError("A converted output changed: " + relative)
    actual = {
        path.relative_to(long_path(root)).as_posix()
        for path in long_path(root / "xlsx").rglob("*") if path.is_file()
    }
    if actual != set(outputs):
        raise ValueError("The XLSX directory contains unaccounted files. Review it before proceeding.")
    return outputs


def package(root, destination):
    root, destination = Path(root).resolve(), Path(destination).resolve()
    with RunLock(root):
        if destination.exists():
            raise ValueError("Choose a new archive name; published versions are not overwritten.")
        outputs = accepted_outputs(root)
        summary = read_json(root / "summary.json")
        metadata = {
            "created_utc": utc_now(), "scope": summary["scope"],
            "source_manifest_sha256": sha256(root / "sources.json"),
            "conversion_manifest_sha256": sha256(root / "conversions.jsonl"),
            "summary": summary,
            "validation": "New conversions: fresh Excel COM reopen. Reused study files: historical validation and current package/hash checks; see each conversion record.",
            "normalization": "Excel SaveAs XLSX; VBA and other unsupported features may be lost. Originals remain in the source archives.",
        }
        destination.parent.mkdir(parents=True, exist_ok=True)
        pending = destination.with_suffix(destination.suffix + ".partial")
        with zipfile.ZipFile(pending, "w", zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=True) as archive:
            for relative in sorted(outputs):
                archive.write(long_path(contained(root, relative)), relative)
            variants = {}
            for row in rows(root / "conversions.jsonl"):
                result = row["result"]
                for artifact in list(result.get("variants", {}).values()) + ([result["conversion_artifact"]] if result.get("conversion_artifact") else []):
                    variants[artifact["path"]] = artifact["sha256"]
            for relative in sorted(variants):
                archive.write(long_path(contained(root, relative)), relative)
            for name in ("sources.json", "inventory.jsonl", "conversions.jsonl", "summary.json"):
                archive.write(root / name, "metadata/" + name)
            for path in sorted((root / "state" / "code").glob("*.json")):
                archive.write(path, "metadata/code/" + path.name)
            for path in sorted((root / "runs").glob("*/run.json")):
                archive.write(path, "metadata/runs/" + path.parent.name + "/run.json")
            archive.writestr("metadata/release.json", json.dumps(metadata, indent=2) + "\n")
            checksums = "".join("{}  {}\n".format(digest, path) for path, digest in sorted({**outputs, **variants}.items()))
            archive.writestr("SHA256SUMS", checksums)
            archive.writestr("README.txt", (
                "Enron spreadsheets normalized with Microsoft Excel\n\n"
                "This release reconstructs spreadsheet candidates from the frozen EDRM v2 XML archive list.\n"
                "It does not represent every document ever held by Enron.\n"
                "xlsx/ contains one file per accepted distinct input content.\n"
                "metadata/conversions.jsonl maps every candidate occurrence to its output or exception.\n"
                "Duplicate occurrences retain their source paths in that manifest.\n"
                "variants/ retains imported study outputs with their original producer metadata.\n"
                "Imported baselines were package-checked, without reopening Excel or repeating shrinking.\n"
                "Unresolved formats and failed conversions remain explicit exceptions.\n"
                "Excel conversion can change workbook features; this is a normalized research input.\n"
                "The source manifest identifies the original archives. No new license is assigned to their contents.\n"
            ))
        os.replace(pending, destination)
        digest = sha256(destination)
        destination.with_suffix(destination.suffix + ".sha256").write_text(digest + "  " + destination.name + "\n", encoding="ascii")
        write_json(destination.with_suffix(destination.suffix + ".json"), {**metadata, "archive_sha256": digest})
        print("Archive ready: " + str(destination))
