#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any


SHEET_PART_RE = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
WORKBOOK_XML = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"

NS = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wb": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


# ----------------------------
# Phase A: fast (compressed only)
# ----------------------------


@dataclass(frozen=True)
class MatchFast:
    relative_path: str
    original_path: str
    shrinked_path: str
    original_bytes: int  # compressed / on-disk
    shrinked_bytes: int  # compressed / on-disk
    bytes_saved: int  # compressed saved
    percent_saved: float  # compressed % saved


# ----------------------------
# Phase B: details (for selected only)
# ----------------------------


@dataclass(frozen=True)
class MatchDetails:
    original_uncompressed_bytes: int
    shrinked_uncompressed_bytes: int
    uncompressed_bytes_saved: int
    uncompressed_percent_saved: float


def iter_files_recursive(root: Path) -> List[Path]:
    return [p for p in root.rglob("*") if p.is_file()]


def relpath_under(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def build_lookup(root: Path) -> Dict[str, Path]:
    out: Dict[str, Path] = {}
    for p in iter_files_recursive(root):
        out[relpath_under(p, root)] = p
    return out


def compute_matches_fast(orig_root: Path, shrink_root: Path) -> List[MatchFast]:
    orig_lookup = build_lookup(orig_root)
    results: List[MatchFast] = []

    for sp in iter_files_recursive(shrink_root):
        rel = relpath_under(sp, shrink_root)
        op = orig_lookup.get(rel)
        if op is None:
            continue

        osize = int(op.stat().st_size)
        ssize = int(sp.stat().st_size)
        saved = osize - ssize
        pct = (saved / osize * 100.0) if osize > 0 else 0.0

        results.append(
            MatchFast(
                relative_path=rel,
                original_path=str(op),
                shrinked_path=str(sp),
                original_bytes=osize,
                shrinked_bytes=ssize,
                bytes_saved=saved,
                percent_saved=float(round(pct, 6)),
            )
        )

    return results


def top_by_absolute_fast(rows: List[MatchFast], n: int) -> List[MatchFast]:
    # absolute reduction by compressed bytes saved (fast criteria)
    return sorted(rows, key=lambda r: r.bytes_saved, reverse=True)[:n]


def top_by_relative_fast(rows: List[MatchFast], n: int) -> List[MatchFast]:
    # relative reduction by compressed percent saved (fast criteria)
    return sorted(rows, key=lambda r: r.percent_saved, reverse=True)[:n]


# ----------------------------
# ZIP + XML helpers (Phase B)
# ----------------------------


def safe_open_zip(path: Path) -> Optional[zipfile.ZipFile]:
    try:
        return zipfile.ZipFile(path, "r")
    except Exception:
        return None


def zip_uncompressed_total(zf: zipfile.ZipFile) -> int:
    total = 0
    for info in zf.infolist():
        total += int(info.file_size)
    return total


def try_read_xml(zf: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        with zf.open(name) as f:
            data = f.read()
        return ET.fromstring(data)
    except Exception:
        return None


def resolve_sheet_names(zf: zipfile.ZipFile) -> Dict[str, str]:
    """
    Map "xl/worksheets/sheetN.xml" -> Excel-visible sheet name.
    """
    result: Dict[str, str] = {}

    wb_root = try_read_xml(zf, WORKBOOK_XML)
    rels_root = try_read_xml(zf, WORKBOOK_RELS)
    if wb_root is None or rels_root is None:
        return result

    rid_to_target: Dict[str, str] = {}
    for rel in rels_root.findall(".//pr:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rid or not target:
            continue
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        rid_to_target[rid] = target.replace("\\", "/")

    for sheet in wb_root.findall(".//wb:sheets/wb:sheet", NS):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(f"{{{NS['r']}}}id")  # r:id
        if not name or not rid:
            continue
        target = rid_to_target.get(rid)
        if not target:
            continue
        if SHEET_PART_RE.match(target):
            result[target] = name

    return result


def list_sheet_infos(zf: zipfile.ZipFile) -> Dict[str, zipfile.ZipInfo]:
    out: Dict[str, zipfile.ZipInfo] = {}
    for info in zf.infolist():
        if SHEET_PART_RE.match(info.filename):
            out[info.filename] = info
    return out


def compute_details_for_workbook(
    orig_xlsx: Path, shrink_xlsx: Path
) -> Tuple[MatchDetails, Dict[str, Any]]:
    """
    Returns:
      - MatchDetails (workbook-level uncompressed totals)
      - per_sheet_details dict:
          { "sheets": [...], "errors": [...] }
    """
    details = MatchDetails(
        original_uncompressed_bytes=0,
        shrinked_uncompressed_bytes=0,
        uncompressed_bytes_saved=0,
        uncompressed_percent_saved=0.0,
    )
    per_sheet: Dict[str, Any] = {"sheets": [], "errors": []}

    oz = safe_open_zip(orig_xlsx)
    if oz is None:
        per_sheet["errors"].append("failed_to_open_original_as_zip")
        return details, per_sheet

    sz = safe_open_zip(shrink_xlsx)
    if sz is None:
        oz.close()
        per_sheet["errors"].append("failed_to_open_shrinked_as_zip")
        return details, per_sheet

    try:
        o_un = zip_uncompressed_total(oz)
        s_un = zip_uncompressed_total(sz)
        un_saved = o_un - s_un
        un_pct = (un_saved / o_un * 100.0) if o_un > 0 else 0.0
        details = MatchDetails(
            original_uncompressed_bytes=int(o_un),
            shrinked_uncompressed_bytes=int(s_un),
            uncompressed_bytes_saved=int(un_saved),
            uncompressed_percent_saved=float(round(un_pct, 6)),
        )

        o_sheets = list_sheet_infos(oz)
        s_sheets = list_sheet_infos(sz)
        o_names = resolve_sheet_names(oz)
        s_names = resolve_sheet_names(sz)

        common = sorted(set(o_sheets.keys()) & set(s_sheets.keys()))
        rows = []
        for part in common:
            oi = o_sheets[part]
            si = s_sheets[part]

            orig_comp = int(oi.compress_size)
            shr_comp = int(si.compress_size)
            comp_saved = orig_comp - shr_comp
            comp_pct = (comp_saved / orig_comp * 100.0) if orig_comp > 0 else 0.0

            orig_un = int(oi.file_size)
            shr_un = int(si.file_size)
            un_saved_sheet = orig_un - shr_un
            un_pct_sheet = (un_saved_sheet / orig_un * 100.0) if orig_un > 0 else 0.0

            excel_name = o_names.get(part) or s_names.get(part) or "(unknown)"

            rows.append(
                {
                    "part": part,
                    "xml_name": part.split("/")[-1],
                    "excel_sheet_name": excel_name,
                    "original_compressed_bytes": orig_comp,
                    "shrinked_compressed_bytes": shr_comp,
                    "compressed_bytes_saved": comp_saved,
                    "compressed_percent_saved": round(comp_pct, 6),
                    "original_uncompressed_bytes": orig_un,
                    "shrinked_uncompressed_bytes": shr_un,
                    "uncompressed_bytes_saved": un_saved_sheet,
                    "uncompressed_percent_saved": round(un_pct_sheet, 6),
                }
            )

        per_sheet["sheets"] = rows
        return details, per_sheet

    finally:
        oz.close()
        sz.close()


# ----------------------------
# Printing
# ----------------------------


def print_table_with_multiline_headers(
    title: str,
    rows: List[MatchFast],
    details_by_rel: Dict[str, MatchDetails],
) -> None:
    print(title)
    if not rows:
        print("  (none)\n")
        return

    # Headers with line breaks as you requested
    headers = [
        "Bytes Saved\n(Compressed)",
        "% Saved\n(Compressed)",
        "Original Bytes\n(Compressed)",
        "Shrinked Bytes\n(Compressed)",
        "Bytes Saved\n(Uncompressed)",
        "% Saved\n(Uncompressed)",
        "Original Bytes\n(Uncompressed)",
        "Shrinked Bytes\n(Uncompressed)",
        "Relative Path",
    ]

    data: List[List[str]] = []
    for r in rows:
        d = details_by_rel.get(r.relative_path)
        if d is None:
            # not analysed (shouldn't happen for top lists, but keep safe)
            d = MatchDetails(0, 0, 0, 0.0)

        data.append(
            [
                str(r.bytes_saved),
                f"{r.percent_saved:.6f}",
                str(r.original_bytes),
                str(r.shrinked_bytes),
                str(d.uncompressed_bytes_saved),
                f"{d.uncompressed_percent_saved:.6f}",
                str(d.original_uncompressed_bytes),
                str(d.shrinked_uncompressed_bytes),
                r.relative_path,
            ]
        )

    widths = [0] * len(headers)

    # Width calc must consider multi-line headers
    header_lines = [h.splitlines() for h in headers]
    for i, lines in enumerate(header_lines):
        widths[i] = max(widths[i], max(len(line) for line in lines))

    for row in data:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(cell))

    def fmt_cells(cells: List[str]) -> str:
        return "  " + "  ".join(cells[i].ljust(widths[i]) for i in range(len(cells)))

    # Print header in multiple lines
    max_header_height = max(len(lines) for lines in header_lines)
    for line_idx in range(max_header_height):
        line_cells = []
        for col_idx, lines in enumerate(header_lines):
            line_cells.append(lines[line_idx] if line_idx < len(lines) else "")
        print(fmt_cells(line_cells))

    print("  " + "  ".join("-" * w for w in widths))
    for row in data:
        print(fmt_cells(row))
    print()


def pick_top_sheets(
    sheet_rows: List[Dict[str, Any]], mode: str, k: int
) -> List[Dict[str, Any]]:
    if mode == "absolute":
        return sorted(
            sheet_rows, key=lambda s: s["uncompressed_bytes_saved"], reverse=True
        )[:k]
    if mode == "relative":
        return sorted(
            sheet_rows, key=lambda s: s["uncompressed_percent_saved"], reverse=True
        )[:k]
    return []


def print_sheet_block(per_sheet: Dict[str, Any], mode: str, k: int) -> None:
    indent = "    "
    if per_sheet.get("errors"):
        print(f"{indent}- (sheet breakdown failed) errors={per_sheet['errors']}")
        return

    sheets = per_sheet.get("sheets", [])
    top = pick_top_sheets(sheets, mode=mode, k=k)
    if not top:
        print(f"{indent}- (no worksheet parts found under xl/worksheets/sheetN.xml)")
        return

    for s in top:
        xml = s["xml_name"]
        excel_name = s["excel_sheet_name"]
        un_saved = s["uncompressed_bytes_saved"]
        un_pct = s["uncompressed_percent_saved"]
        comp_saved = s["compressed_bytes_saved"]
        comp_pct = s["compressed_percent_saved"]

        print(
            f"{indent}- {xml}  |  {excel_name}\n"
            f"{indent}  saved: {un_saved} ({un_pct:.6f}%) (Uncompressed)"
            f"  |  {comp_saved} ({comp_pct:.6f}%) (Compressed)"
        )


# ----------------------------
# JSON report
# ----------------------------


def to_jsonable_fast(rows: List[MatchFast]) -> List[Dict[str, Any]]:
    return [
        {
            "relative_path": r.relative_path,
            "original_path": r.original_path,
            "shrinked_path": r.shrinked_path,
            "original_bytes_compressed": r.original_bytes,
            "shrinked_bytes_compressed": r.shrinked_bytes,
            "bytes_saved_compressed": r.bytes_saved,
            "percent_saved_compressed": r.percent_saved,
        }
        for r in rows
    ]


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Compare XLSX reductions: pick top by compressed size first, then analyse selected for uncompressed + sheets."
    )
    ap.add_argument("orig_dir", type=Path)
    ap.add_argument("shrink_dir", type=Path)
    ap.add_argument("--out", type=Path, default=Path("comparison_report.json"))
    ap.add_argument("--topn", type=int, default=10)
    ap.add_argument("--per-sheet-top", type=int, default=3)
    args = ap.parse_args()

    orig_dir = args.orig_dir.resolve()
    shrink_dir = args.shrink_dir.resolve()
    if not orig_dir.is_dir():
        print(f"Original directory not found: {orig_dir}", file=sys.stderr)
        return 2
    if not shrink_dir.is_dir():
        print(f"Shrinked directory not found: {shrink_dir}", file=sys.stderr)
        return 2

    # Phase A: fast compressed-only matching
    matches_fast = compute_matches_fast(orig_dir, shrink_dir)
    print(f"Matched files: {len(matches_fast)}\n")

    top_abs = top_by_absolute_fast(matches_fast, args.topn)
    top_rel = top_by_relative_fast(matches_fast, args.topn)

    # Union of selected (avoid double work)
    selected: Dict[str, MatchFast] = {}
    for r in top_abs:
        selected[r.relative_path] = r
    for r in top_rel:
        selected[r.relative_path] = r

    # Phase B: details only for selected
    details_by_rel: Dict[str, MatchDetails] = {}
    per_sheet_by_rel: Dict[str, Any] = {}
    for rel, r in selected.items():
        details, per_sheet = compute_details_for_workbook(
            Path(r.original_path), Path(r.shrinked_path)
        )
        details_by_rel[rel] = details
        per_sheet_by_rel[rel] = per_sheet

    # Print absolute table + sheet breakdown
    print_table_with_multiline_headers(
        f"Top {args.topn} by absolute reduction (Workbook; ranked by Bytes Saved (Compressed)):",
        top_abs,
        details_by_rel,
    )
    for r in top_abs:
        print(f"  {r.relative_path}")
        print_sheet_block(
            per_sheet_by_rel[r.relative_path], mode="absolute", k=args.per_sheet_top
        )
        print()

    # Print relative table + sheet breakdown
    print_table_with_multiline_headers(
        f"Top {args.topn} by relative reduction (Workbook; ranked by % Saved (Compressed)):",
        top_rel,
        details_by_rel,
    )
    for r in top_rel:
        print(f"  {r.relative_path}")
        print_sheet_block(
            per_sheet_by_rel[r.relative_path], mode="relative", k=args.per_sheet_top
        )
        print()

    # JSON report
    report = {
        "inputs": {"original_dir": str(orig_dir), "shrinked_dir": str(shrink_dir)},
        "matched_files_count": len(matches_fast),
        "topn": args.topn,
        "per_sheet_topk": args.per_sheet_top,
        "selection_criteria": "Top lists selected by compressed XLSX file size only; details computed only for selected union.",
        "top_absolute": to_jsonable_fast(top_abs),
        "top_relative": to_jsonable_fast(top_rel),
        "details_uncompressed_for_selected": {
            rel: {
                "original_bytes_uncompressed": details_by_rel[
                    rel
                ].original_uncompressed_bytes,
                "shrinked_bytes_uncompressed": details_by_rel[
                    rel
                ].shrinked_uncompressed_bytes,
                "bytes_saved_uncompressed": details_by_rel[
                    rel
                ].uncompressed_bytes_saved,
                "percent_saved_uncompressed": details_by_rel[
                    rel
                ].uncompressed_percent_saved,
            }
            for rel in selected.keys()
        },
        "per_sheet_details_for_selected": per_sheet_by_rel,
    }

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(f"Wrote report: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
