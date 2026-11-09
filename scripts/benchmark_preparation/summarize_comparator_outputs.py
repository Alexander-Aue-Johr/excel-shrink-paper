"""Verify retained comparator packages and summarize the frozen six-file panel."""

import argparse
import csv
import hashlib
import json
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NAMESPACE = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def identify_package(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    with ZipFile(path) as archive:
        invalid_member = archive.testzip()
    if invalid_member:
        raise ValueError(f"Invalid ZIP member in {path}: {invalid_member}")
    return {"bytes": path.stat().st_size, "sha256": digest.hexdigest(), "zip_crc_valid": True}


def count_wanderungen_zero_tail(path):
    cells = 0
    numeric_zeros = 0
    with ZipFile(path) as archive, archive.open("xl/worksheets/sheet31.xml") as stream:
        for _, element in ET.iterparse(stream, events=("end",)):
            if element.tag == NAMESPACE + "c":
                cells += 1
                value = element.find(NAMESPACE + "v")
                if element.get("t", "n") == "n" and element.find(NAMESPACE + "f") is None and value is not None:
                    try:
                        numeric_zeros += float(value.text) == 0
                    except (TypeError, ValueError):
                        pass
                element.clear()
            elif element.tag == NAMESPACE + "row":
                element.clear()
    return {"worksheet_part": "xl/worksheets/sheet31.xml", "cells": cells, "numeric_zero_constants": numeric_zeros}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--study-root", type=Path, required=True)
    root = parser.parse_args().study_root / "comparators"
    panel = json.loads((root.parent / "results/provisional-panel.json").read_text())
    rows = []
    for candidate in panel["destatis_candidates"]:
        name = candidate["file"]
        for tool in ("nxpowerlite", "check-performance", "xlstyles", "sheetshrink"):
            path = root / tool / "outputs" / name
            identity = identify_package(path)
            row = {"file": name, "tool": tool, "input_bytes": candidate["input_bytes"], **identity}
            row["reduction_percent"] = 100 * (1 - identity["bytes"] / candidate["input_bytes"])
            if name == "wanderungen-2010120217005.xlsx":
                row["zero_tail"] = count_wanderungen_zero_tail(path)
            rows.append(row)
    (root / "output-manifest.json").write_text(json.dumps(rows, indent=2) + "\n")
    with (root / "sizes.csv").open("w", newline="") as stream:
        fields = ["file", "tool", "input_bytes", "bytes", "reduction_percent", "sha256", "zip_crc_valid"]
        writer = csv.DictWriter(stream, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    print(json.dumps([row for row in rows if "zero_tail" in row], indent=2))


if __name__ == "__main__":
    main()
