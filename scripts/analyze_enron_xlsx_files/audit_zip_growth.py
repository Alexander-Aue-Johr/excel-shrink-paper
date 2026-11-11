"""Audit exact XLSX pairs and compare payload sizes under common DEFLATE settings."""
import argparse
import csv
import hashlib
import json
from pathlib import Path
import zipfile
import zlib


def inspect_archive(path):
    members = {}
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            if info.filename in members:
                raise ValueError(f"Duplicate archive member in {path}: {info.filename}")
            digest = hashlib.sha256()
            compressors = {level: zlib.compressobj(level, zlib.DEFLATED, -15) for level in (2, 6)}
            sizes = dict.fromkeys(compressors, 0)
            with archive.open(info) as stream:
                while chunk := stream.read(1024 * 1024):
                    digest.update(chunk)
                    for level, compressor in compressors.items():
                        sizes[level] += len(compressor.compress(chunk))
            for level, compressor in compressors.items():
                sizes[level] += len(compressor.flush())
            members[info.filename] = dict(uncompressed=info.file_size, compressed=info.compress_size,
                method=info.compress_type, sha256=digest.hexdigest(), deflate_2=sizes[2], deflate_6=sizes[6])
    return members


def verified_pair(case, directory):
    paths = []
    for side, size_key, hash_key in (("original", "input_bytes", "original_sha256"),
                                     ("output", "output_bytes", "output_sha256")):
        path = directory / f"{side}.xlsx"
        with path.open("rb") as stream:
            digest = hashlib.file_digest(stream, "sha256").hexdigest()
        if path.stat().st_size != case[size_key] or digest != case[hash_key]:
            raise ValueError(f"Transferred identity differs for {case['case_id']}/{side}")
        paths.append(path)
    return paths


def analyze(case, pair_directory):
    original, output = verified_pair(case, pair_directory)
    left, right = inspect_archive(original), inspect_archive(output)
    members = []
    for name in sorted(left.keys() | right.keys()):
        a, b = left.get(name), right.get(name)
        members.append(dict(part=name, original=a, output=b,
            payload_identical=bool(a and b and a["sha256"] == b["sha256"]),
            uncompressed_delta=(b or {}).get("uncompressed", 0)-(a or {}).get("uncompressed", 0),
            compressed_delta=(b or {}).get("compressed", 0)-(a or {}).get("compressed", 0)))
    summary = dict(case_id=case["case_id"], file=case["file"],
        archive_growth_bytes=output.stat().st_size-original.stat().st_size,
        uncompressed_growth_bytes=sum(m["uncompressed_delta"] for m in members),
        xml_growth_bytes=sum(m["uncompressed_delta"] for m in members if m["part"].endswith((".xml", ".rels"))),
        compressed_payload_growth_bytes=sum(m["compressed_delta"] for m in members),
        changed_members=sum(not m["payload_identical"] for m in members))
    summary["archive_overhead_growth_bytes"] = summary["archive_growth_bytes"]-summary["compressed_payload_growth_bytes"]
    for level in (2, 6):
        key = f"deflate_{level}"
        delta = sum(m[key] for m in right.values())-sum(m[key] for m in left.values())
        summary[f"common_deflate_{level}_payload_delta"] = delta
        summary[f"actual_minus_common_deflate_{level}_delta"] = summary["compressed_payload_growth_bytes"]-delta
    identity = {key: case[key] for key in ("original_sha256", "output_sha256")}
    return dict(identity=identity, summary=summary, members=members)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("request", type=Path)
    parser.add_argument("pairs", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    summaries = []
    missing = []
    for case in json.loads(args.request.read_text(encoding="utf-8-sig"))["cases"]:
        pair = args.pairs / case["case_id"]
        if not all((pair / f"{side}.xlsx").is_file() for side in ("original", "output")):
            missing.append(case["case_id"])
            continue
        receipt = args.output / f"{case['case_id']}.json"
        verified_pair(case, pair)
        result = None
        if receipt.exists():
            result = json.loads(receipt.read_text())
        identity = {key: case[key] for key in ("original_sha256", "output_sha256")}
        if result is None or result.get("identity") != identity:
            result = analyze(case, pair)
            receipt.write_text(json.dumps(result, indent=2), encoding="utf-8")
        summaries.append(result["summary"])
        print(json.dumps(result["summary"]), flush=True)
    if summaries:
        with (args.output / "summary.csv").open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=list(summaries[0]))
            writer.writeheader()
            writer.writerows(summaries)
    (args.output / "status.json").write_text(json.dumps(dict(completed=len(summaries), missing=missing), indent=2))


if __name__ == "__main__":
    main()
