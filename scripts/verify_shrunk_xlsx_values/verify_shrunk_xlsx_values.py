#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Verify that shrunk XLSX files still produce the same calculated values as originals,
but SKIP files when all worksheet XML files in xl/worksheets have identical sizes.

Key changes vs multiprocessing version:
- Excel COM comparison runs ONLY in the main thread (no multiprocessing).
- Fast pre-check:
    If *all* files in xl/worksheets (sheet*.xml) exist in both and have identical sizes,
    skip the file entirely (log as skipped).
- If any worksheet file differs (name missing/extra/size mismatch), then:
    open both in Excel, CalculateFull, and compare computed values.

Requirements (Windows):
  pip install pywin32

Example:
  python verify_shrunk_xlsx_values_skip_if_ws_equal.py "C:\\orig" "C:\\shrunk" --out report.json
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import logging
import os
import re
import time
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import hashlib
import shutil
import tempfile

import pythoncom
import win32com.client


# ----------------------------
# Helpers
# ----------------------------
_SHEET_XML_RE = re.compile(r"^xl/worksheets/sheet\d+\.xml$", re.IGNORECASE)


def col_to_letters(col: int) -> str:
    if col <= 0:
        return "?"
    letters: List[str] = []
    n = col
    while n:
        n, r = divmod(n - 1, 26)
        letters.append(chr(ord("A") + r))
    return "".join(reversed(letters))


def a1_addr(row: int, col: int) -> str:
    return f"{col_to_letters(col)}{row}"


def normalise_empty(v: Any) -> Any:
    """Treat whitespace-only strings as empty (None)."""
    if v is None:
        return None
    if isinstance(v, str):
        return None if v.strip() == "" else v
    return v


def safe_value_for_report(v: Any, max_len: int = 200) -> Any:
    try:
        if v is None:
            return None
        if isinstance(v, (int, float, bool, str)):
            if isinstance(v, str) and len(v) > max_len:
                return v[:max_len] + "…"
            return v
        s = str(v)
        if len(s) > max_len:
            s = s[:max_len] + "…"
        return s
    except Exception:
        return "<unrepr>"


def safe_cell_formula(
    sheet: Any, row: int, col: int, max_len: int = 500
) -> Optional[str]:
    """
    Returns the formula string for a cell if present, else None.
    Uses Range.Formula (not FormulaR1C1). For non-formula cells Excel returns the value itself
    or an empty string depending on type; we only keep strings that look like formulas.
    """
    try:
        f = sheet.Cells(row, col).Formula
        if f is None:
            return None
        s = str(f)
        # In Excel COM, formulas typically start with "="
        if not s.startswith("="):
            return None
        if len(s) > max_len:
            return s[:max_len] + "…"
        return s
    except Exception:
        return None


def coerce_2d(value2: Any, rows: int, cols: int) -> Tuple[Tuple[Any, ...], ...]:
    """
    Coerce Excel Range.Value2 output to a rows x cols tuple-of-tuples.
    """
    if rows == 1 and cols == 1:
        return ((value2,),)

    # Already 2D?
    if isinstance(value2, tuple) and len(value2) == rows and rows > 1:
        if isinstance(value2[0], tuple):
            return value2  # type: ignore

    # 1 x N
    if rows == 1:
        if isinstance(value2, tuple):
            if (
                len(value2) == 1
                and isinstance(value2[0], tuple)
                and len(value2[0]) == cols
            ):
                return (value2[0],)  # type: ignore
            if len(value2) == cols:
                return (value2,)  # type: ignore
        return ((value2,),)

    # N x 1
    if cols == 1:
        if isinstance(value2, tuple) and len(value2) == rows:
            if isinstance(value2[0], tuple):
                return value2  # type: ignore
            return tuple((v,) for v in value2)  # type: ignore
        return tuple((None,) for _ in range(rows))

    # Fallback: try best-effort 2D
    if isinstance(value2, tuple):
        out: List[Tuple[Any, ...]] = []
        for r in range(rows):
            row_val = value2[r] if r < len(value2) else None
            if isinstance(row_val, tuple):
                out.append(
                    tuple(row_val[c] if c < len(row_val) else None for c in range(cols))
                )
            else:
                out.append(tuple(None for _ in range(cols)))
        return tuple(out)

    return tuple(tuple(None for _ in range(cols)) for _ in range(rows))


def list_xlsx_files(folder: Path) -> Dict[str, Path]:
    out: Dict[str, Path] = {}
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() == ".xlsx":
            out[p.name] = p.absolute()
    return out


# ----------------------------
# Path length handling (Windows MAX_PATH workaround)
# ----------------------------
MAX_EXCEL_PATH_LEN = 200  # as requested


def _abs_path_str(p: Path) -> str:
    # resolve() can fail on some weird paths; fall back gracefully
    try:
        return str(p.resolve())
    except Exception:
        return str(p.absolute())


def _temp_name_from_abs_path(abs_path: str, suffix: str = ".xlsx") -> str:
    """
    Stable filename derived from the original absolute path.
    Using sha256(abs_path) avoids collisions across orig/shrunk that have different paths.
    """
    h = hashlib.sha256(abs_path.encode("utf-8", errors="ignore")).hexdigest()
    return f"{h}{suffix}"


def ensure_short_path_for_excel(p: Path, temp_dir: Path) -> Path:
    """
    If absolute path length > MAX_EXCEL_PATH_LEN, copy to temp_dir and return the temp path.
    Otherwise return original path.
    """
    abs_str = _abs_path_str(p)
    if len(abs_str) <= MAX_EXCEL_PATH_LEN:
        return p

    temp_dir.mkdir(parents=True, exist_ok=True)

    # keep .xlsx suffix
    suffix = p.suffix if p.suffix else ".xlsx"
    tmp_name = _temp_name_from_abs_path(abs_str, suffix=suffix)
    tmp_path = temp_dir / tmp_name

    # Copy only once per run if already present
    if not tmp_path.exists():
        logging.warning(
            "Path too long (%d > %d). Copying to temp: %s -> %s",
            len(abs_str),
            MAX_EXCEL_PATH_LEN,
            abs_str,
            str(tmp_path),
        )
        shutil.copy2(p, tmp_path)
    else:
        logging.debug(
            "Temp copy already exists for long path (%d chars): %s -> %s",
            len(abs_str),
            abs_str,
            str(tmp_path),
        )

    return tmp_path


# ----------------------------
# Resume from existing report.json
# ----------------------------
def load_previous_report(out_path: Path) -> tuple[set[str], dict[str, str]]:
    """
    Returns:
      - same_files_prev: set of filenames that were "same"
      - skipped_reason_prev: filename -> reason for skipped files
    If report doesn't exist or is unreadable, returns empty sets.
    """
    if not out_path.exists():
        return set(), {}

    try:
        with out_path.open("r", encoding="utf-8") as f:
            rep = json.load(f)

        same_files_prev = set(rep.get("same_files", []) or [])

        skipped_reason_prev: dict[str, str] = {}
        for item in rep.get("skipped_files", []) or []:
            fn = item.get("filename")
            rsn = item.get("reason")
            if isinstance(fn, str) and isinstance(rsn, str):
                skipped_reason_prev[fn] = rsn

        return same_files_prev, skipped_reason_prev

    except Exception as e:
        logging.warning("Could not read existing report %s: %s", out_path, e)
        return set(), {}


# ----------------------------
# Pre-check: worksheet xml sizes
# ----------------------------
def worksheet_xml_sizes(xlsx_path: Path) -> Dict[str, int]:
    """
    Return mapping: 'xl/worksheets/sheetN.xml' -> uncompressed file_size within zip.
    Only includes the canonical sheetN.xml pattern (ignores xl/worksheets/sheet.xml etc.).
    """
    sizes: Dict[str, int] = {}
    with zipfile.ZipFile(xlsx_path, "r") as z:
        for info in z.infolist():
            name = info.filename
            if _SHEET_XML_RE.match(name):
                sizes[name.lower()] = int(info.file_size)
    return sizes


def should_skip_by_worksheet_sizes(orig: Path, shr: Path) -> Tuple[bool, str]:
    """
    Returns (skip, reason).
    skip=True when:
      - both have identical set of sheetN.xml files AND
      - all corresponding file sizes match exactly.
    """
    try:
        a = worksheet_xml_sizes(orig)
        b = worksheet_xml_sizes(shr)
    except Exception as e:
        # If we can't read zip content reliably, be safe: do NOT skip
        return (False, f"precheck_error: {type(e).__name__}: {e}")

    if not a and not b:
        # No sheets found by our pattern -> safest: do NOT skip (force Excel check)
        return (False, "precheck_no_sheet_xml_found")

    if set(a.keys()) != set(b.keys()):
        extra_a = sorted(set(a.keys()) - set(b.keys()))
        extra_b = sorted(set(b.keys()) - set(a.keys()))
        return (
            False,
            f"worksheet_file_set_diff extra_in_orig={extra_a} extra_in_shrunk={extra_b}",
        )

    for k in a.keys():
        if a[k] != b[k]:
            return (False, f"worksheet_size_diff {k} orig={a[k]} shrunk={b[k]}")

    return (True, "all_xl_worksheets_sheetN_xml_sizes_equal")


# ----------------------------
# Excel COM compare
# ----------------------------
def open_excel_app() -> Any:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    try:
        excel.AutomationSecurity = 3  # disable macros
    except Exception:
        pass

    return excel


def open_workbook(excel: Any, file_path: Path) -> Any:
    return excel.Workbooks.Open(
        str(file_path),
        UpdateLinks=0,
        ReadOnly=True,
        IgnoreReadOnlyRecommended=True,
        AddToMru=False,
        Notify=False,
    )


def get_used_bounds(sheet: Any) -> Optional[Tuple[int, int, int, int]]:
    try:
        used = sheet.UsedRange
        r0 = int(used.Row)
        c0 = int(used.Column)
        nrows = int(used.Rows.Count)
        ncols = int(used.Columns.Count)
        if nrows <= 0 or ncols <= 0:
            return None
        r1 = r0 + nrows - 1
        c1 = c0 + ncols - 1
        return r0, c0, r1, c1
    except Exception:
        return None


@dataclasses.dataclass
class Diff:
    sheet: str
    cell: str
    original: Any
    shrunk: Any
    original_formula: Optional[str] = None
    shrunk_formula: Optional[str] = None


@dataclasses.dataclass
class FileResult:
    filename: str
    status: str  # "same" | "different" | "error" | "skipped"
    reason: Optional[str] = None
    diff_count: int = 0
    diffs: Optional[List[Diff]] = None
    elapsed_sec: float = 0.0


def compare_sheet_values(
    sheet_a: Any,
    sheet_b: Any,
    sheet_name: str,
    row_chunk: int,
    col_chunk: int,
    max_diffs: int,
) -> Tuple[int, List[Diff]]:

    # if sheet_a.UsedRange.Value2 == sheet_b.UsedRange.Value2:
    #     return 0, []

    bounds_a = get_used_bounds(sheet_a)
    bounds_b = get_used_bounds(sheet_b)

    if bounds_a is None and bounds_b is None:
        return 0, []

    if bounds_a is None:
        min_r, min_c, max_r, max_c = bounds_b  # type: ignore
    elif bounds_b is None:
        min_r, min_c, max_r, max_c = bounds_a
    else:
        min_r = min(bounds_a[0], bounds_b[0])
        min_c = min(bounds_a[1], bounds_b[1])
        max_r = max(bounds_a[2], bounds_b[2])
        max_c = max(bounds_a[3], bounds_b[3])

    diffs: List[Diff] = []
    diff_count = 0

    r = min_r
    while r <= max_r:
        r2 = min(r + row_chunk - 1, max_r)
        c = min_c
        while c <= max_c:
            c2 = min(c + col_chunk - 1, max_c)

            rng_a = sheet_a.Range(sheet_a.Cells(r, c), sheet_a.Cells(r2, c2))
            rng_b = sheet_b.Range(sheet_b.Cells(r, c), sheet_b.Cells(r2, c2))

            va = rng_a.Value2
            vb = rng_b.Value2

            rows = r2 - r + 1
            cols = c2 - c + 1

            a2d = coerce_2d(va, rows, cols)
            b2d = coerce_2d(vb, rows, cols)

            for i in range(rows):
                ra = a2d[i]
                rb = b2d[i]
                row_idx = r + i
                for j in range(cols):
                    col_idx = c + j

                    na = normalise_empty(ra[j] if j < len(ra) else None)
                    nb = normalise_empty(rb[j] if j < len(rb) else None)

                    if na is None and nb is None:
                        continue
                    if na == nb:
                        continue

                    diff_count += 1
                    if len(diffs) < max_diffs:
                        diffs.append(
                            Diff(
                                sheet=sheet_name,
                                cell=a1_addr(row_idx, col_idx),
                                original=safe_value_for_report(
                                    ra[j] if j < len(ra) else None
                                ),
                                shrunk=safe_value_for_report(
                                    rb[j] if j < len(rb) else None
                                ),
                                original_formula=safe_cell_formula(
                                    sheet_a, row_idx, col_idx
                                ),
                                shrunk_formula=safe_cell_formula(
                                    sheet_b, row_idx, col_idx
                                ),
                            )
                        )

            del rng_a, rng_b, va, vb, a2d, b2d
            c = c2 + 1
        r = r2 + 1

    return diff_count, diffs


def compare_workbooks_in_excel(
    excel_a: Any,
    excel_b: Any,
    original_path: Path,
    shrunk_path: Path,
    temp_dir: Path,
    row_chunk: int,
    col_chunk: int,
    max_diffs_per_file: int,
) -> FileResult:

    t0 = time.time()
    filename = original_path.name

    wb_a = None
    wb_b = None
    try:
        op = ensure_short_path_for_excel(original_path, temp_dir)
        sp = ensure_short_path_for_excel(shrunk_path, temp_dir)

        wb_a = open_workbook(excel_a, op)
        wb_b = open_workbook(excel_b, sp)

        excel_a.CalculateFull()
        excel_b.CalculateFull()

        count_a = int(wb_a.Worksheets.Count)
        count_b = int(wb_b.Worksheets.Count)
        if count_a != count_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=f"sheet_count_mismatch: original={count_a}, shrunk={count_b}",
                diffs=[],
                diff_count=0,
                elapsed_sec=time.time() - t0,
            )

        names_a = [str(wb_a.Worksheets(i).Name) for i in range(1, count_a + 1)]
        names_b = [str(wb_b.Worksheets(i).Name) for i in range(1, count_b + 1)]
        if names_a != names_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=f"sheet_name_mismatch: original={names_a}, shrunk={names_b}",
                diffs=[],
                diff_count=0,
                elapsed_sec=time.time() - t0,
            )

        sheet_b_by_name = {
            str(wb_b.Worksheets(i).Name): wb_b.Worksheets(i)
            for i in range(1, count_b + 1)
        }

        total_diff_count = 0
        all_diffs: List[Diff] = []

        for i in range(1, count_a + 1):
            sheet_a = wb_a.Worksheets(i)
            name = str(sheet_a.Name)
            sheet_b = sheet_b_by_name[name]

            remaining = max(0, max_diffs_per_file - len(all_diffs))
            diff_count, diffs = compare_sheet_values(
                sheet_a,
                sheet_b,
                name,
                row_chunk=row_chunk,
                col_chunk=col_chunk,
                max_diffs=remaining,
            )
            total_diff_count += diff_count
            if diffs:
                all_diffs.extend(diffs)

            del sheet_a, sheet_b, diffs

        if total_diff_count == 0:
            return FileResult(
                filename=filename,
                status="same",
                reason=None,
                diffs=[],
                diff_count=0,
                elapsed_sec=time.time() - t0,
            )

        return FileResult(
            filename=filename,
            status="different",
            reason="value_mismatch",
            diffs=all_diffs,
            diff_count=total_diff_count,
            elapsed_sec=time.time() - t0,
        )

    except Exception as e:
        return FileResult(
            filename=filename,
            status="error",
            reason=f"{type(e).__name__}: {e}",
            diffs=[],
            diff_count=0,
            elapsed_sec=time.time() - t0,
        )
    finally:
        try:
            if wb_a is not None:
                wb_a.Close(SaveChanges=False)
                logging.info("Closed workbook A")
        except Exception:
            pass
        try:
            if wb_b is not None:
                wb_b.Close(SaveChanges=False)
                logging.info("Closed workbook B")
        except Exception:
            pass
        try:
            del wb_a, wb_b
        except Exception:
            pass


# ----------------------------
# CLI / main
# ----------------------------
def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Verify shrunk .xlsx values; skip if worksheet XML sizes match."
    )
    ap.add_argument(
        "originals",
        type=Path,
        help="Folder containing original .xlsx files (non-recursive).",
    )
    ap.add_argument(
        "shrunk",
        type=Path,
        help="Folder containing shrunk .xlsx files (non-recursive).",
    )
    ap.add_argument(
        "--out", type=Path, default=Path("verify_report.json"), help="Output JSON path."
    )
    ap.add_argument(
        "--row-chunk",
        type=int,
        default=2000,
        help="Rows per COM block read (memory control).",
    )
    ap.add_argument(
        "--col-chunk",
        type=int,
        default=256,
        help="Cols per COM block read (memory control).",
    )
    ap.add_argument(
        "--max-diffs-per-file",
        type=int,
        default=200,
        help="Max diff examples stored per file.",
    )
    ap.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level.",
    )
    ap.add_argument(
        "--recompare-all",
        action="store_true",
        help="Ignore existing --out report and compare all common files again.",
    )

    return ap.parse_args()


def main() -> int:
    args = parse_args()
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
    )

    if not args.originals.is_dir():
        logging.error("Originals folder invalid: %s", args.originals)
        return 2
    if not args.shrunk.is_dir():
        logging.error("Shrunk folder invalid: %s", args.shrunk)
        return 2

    same_prev, skipped_prev = load_previous_report(args.out)

    orig_map = list_xlsx_files(args.originals)
    shr_map = list_xlsx_files(args.shrunk)

    common_names = sorted(set(orig_map.keys()) & set(shr_map.keys()))
    missing_in_shrunk = sorted(set(orig_map.keys()) - set(shr_map.keys()))
    missing_in_originals = sorted(set(shr_map.keys()) - set(orig_map.keys()))

    logging.info(
        "Original files: %d | Shrunk files: %d | Common: %d",
        len(orig_map),
        len(shr_map),
        len(common_names),
    )
    if missing_in_shrunk:
        logging.warning("Missing in shrunk (skipped): %d", len(missing_in_shrunk))
    if missing_in_originals:
        logging.warning("Missing in originals (ignored): %d", len(missing_in_originals))

    results: List[FileResult] = []

    # Initialise COM + Excel once in main thread
    pythoncom.CoInitialize()
    excel_a = None
    excel_b = None
    try:
        excel_a = open_excel_app()
        excel_b = open_excel_app()

        with tempfile.TemporaryDirectory(prefix="xlsx_compare_") as td:
            temp_dir = Path(td)

            # loop with index
            for idx, name in enumerate(common_names, start=1):
                op = orig_map[name]
                sp = shr_map[name]

                # 0) Resume logic: skip files already proven same,
                # or previously skipped for sheetN.xml size equality,
                # unless --recompare-all was specified.
                if not args.recompare_all:
                    if name in same_prev:
                        logging.debug(
                            "SKIP    %-60s reason=%s",
                            name,
                            "already_same_in_existing_report",
                        )
                        results.append(
                            FileResult(
                                filename=name,
                                status="same",
                                reason="already_same_in_existing_report",
                                diffs=[],
                                diff_count=0,
                                elapsed_sec=0.0,
                            )
                        )
                        continue

                    prev_reason = skipped_prev.get(name)
                    if prev_reason == "all_xl_worksheets_sheetN_xml_sizes_equal":
                        logging.debug("SKIP    %-60s reason=%s", name, prev_reason)
                        results.append(
                            FileResult(
                                filename=name,
                                status="skipped",
                                reason=prev_reason,
                                diffs=[],
                                diff_count=0,
                                elapsed_sec=0.0,
                            )
                        )
                        continue

                # 1) Pre-check: skip if xl/worksheets/sheetN.xml sizes equal
                skip, skip_reason = should_skip_by_worksheet_sizes(op, sp)
                if skip:
                    logging.debug("SKIP    %-60s reason=%s", name, skip_reason)
                    results.append(
                        FileResult(
                            filename=name,
                            status="skipped",
                            reason=skip_reason,
                            diffs=[],
                            diff_count=0,
                            elapsed_sec=0.0,
                        )
                    )
                    continue

                logging.debug("NEEDCHK %-60s reason=%s", name, skip_reason)

                # 2) Excel compare (main thread)
                logging.info(
                    "Comparing workbook %s from %s: %s ", name, idx, len(common_names)
                )
                res = compare_workbooks_in_excel(
                    excel_a,
                    excel_b,
                    op,
                    sp,
                    temp_dir=temp_dir,
                    row_chunk=int(args.row_chunk),
                    col_chunk=int(args.col_chunk),
                    max_diffs_per_file=int(args.max_diffs_per_file),
                )
                results.append(res)

                if res.status == "same":
                    logging.info("SAME    %-60s (%.1fs)", res.filename, res.elapsed_sec)
                elif res.status == "different":
                    logging.warning(
                        "DIFF    %-60s diffs=%d (%.1fs) reason=%s",
                        res.filename,
                        res.diff_count,
                        res.elapsed_sec,
                        res.reason,
                    )
                else:
                    logging.error(
                        "ERROR   %-60s (%.1fs) %s",
                        res.filename,
                        res.elapsed_sec,
                        res.reason,
                    )

                # Drop refs where possible
                del res

    finally:
        try:
            if excel_a is not None:
                excel_a.Quit()
        except Exception:
            pass
        try:
            del excel_a
        except Exception:
            pass
        try:
            if excel_b is not None:
                excel_b.Quit()
        except Exception:
            pass
        try:
            del excel_b
        except Exception:
            pass
        pythoncom.CoUninitialize()

    # JSON report
    same_files = sorted([r.filename for r in results if r.status == "same"])
    skipped_files = sorted([r.filename for r in results if r.status == "skipped"])
    diff_files = sorted([r.filename for r in results if r.status == "different"])
    err_files = sorted([r.filename for r in results if r.status == "error"])

    report: Dict[str, Any] = {
        "meta": {
            "originals_folder": str(args.originals),
            "shrunk_folder": str(args.shrunk),
            "row_chunk": int(args.row_chunk),
            "col_chunk": int(args.col_chunk),
            "max_diffs_per_file": int(args.max_diffs_per_file),
            "common_files": len(common_names),
            "missing_in_shrunk": missing_in_shrunk,
            "missing_in_originals": missing_in_originals,
        },
        "summary": {
            "same": len(same_files),
            "skipped": len(skipped_files),
            "different": len(diff_files),
            "error": len(err_files),
        },
        "same_files": same_files,
        "skipped_files": [],
        "different_files": [],
        "errors": [],
    }

    for r in results:
        if r.status == "skipped":
            report["skipped_files"].append({"filename": r.filename, "reason": r.reason})
        elif r.status == "different":
            report["different_files"].append(
                {
                    "filename": r.filename,
                    "reason": r.reason,
                    "diff_count": r.diff_count,
                    "elapsed_sec": r.elapsed_sec,
                    "diffs": [
                        {
                            "sheet": d.sheet,
                            "cell": d.cell,
                            "original": d.original,
                            "shrunk": d.shrunk,
                            "original_formula": d.original_formula,
                            "shrunk_formula": d.shrunk_formula,
                        }
                        for d in (r.diffs or [])
                    ],
                }
            )
        elif r.status == "error":
            report["errors"].append(
                {
                    "filename": r.filename,
                    "reason": r.reason,
                    "elapsed_sec": r.elapsed_sec,
                }
            )

    report["skipped_files"].sort(key=lambda x: x["filename"])
    report["different_files"].sort(key=lambda x: x["filename"])
    report["errors"].sort(key=lambda x: x["filename"])

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with args.out.open("w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    logging.info("Wrote report: %s", args.out)

    # Drop references (symbolic but fine)
    del results, report, orig_map, shr_map

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
