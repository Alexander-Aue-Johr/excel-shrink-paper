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
    repair workbook.xml in temporary copies, open both in Excel, CalculateFull,
    and compare computed values.

Requirements (Windows):
  pip install pywin32

Example:
  python verify_shrunk_xlsx_values_skip_if_ws_equal.py "C:\\orig" "C:\\shrunk" --out report.json
"""

from __future__ import annotations

import argparse
import csv
import contextlib
import ctypes
import ctypes.wintypes
import dataclasses
import datetime
import json
import logging
import os
import re
import subprocess
import sys
import time
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import hashlib
import shutil
import tempfile

import pythoncom
import win32com.client
import win32gui
import win32process


# ----------------------------
# Helpers
# ----------------------------
_SHEET_XML_RE = re.compile(r"^xl/worksheets/sheet\d+\.xml$", re.IGNORECASE)
_WORKBOOK_XML = "xl/workbook.xml"
_DEFINED_NAMES_RE = re.compile(rb"(<definedNames\b[^>]*>)(.*?)(</definedNames>)", re.S)
_DEFINED_NAME_START_RE = re.compile(
    rb'<definedName\b[^>]*?name="([^"]*?)"'
    rb'(?:[^>]*?localSheetId="([^"]*?)")?'
    rb"[^>]*?>",
    re.S,
)
_DEFINED_NAME_NODE_RE = re.compile(
    rb'<definedName\b[^>]*?name="([^"]*?)"'
    rb'(?:[^>]*?localSheetId="([^"]*?)")?'
    rb"[^>]*?>.*?</definedName>",
    re.S,
)


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


class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.wintypes.DWORD),
        ("PageFaultCount", ctypes.wintypes.DWORD),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
        ("PrivateUsage", ctypes.c_size_t),
    ]


def excel_process_id(excel: Any) -> Optional[int]:
    try:
        hwnd = int(excel.Hwnd)
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid) if pid else None
    except Exception:
        return None


def process_memory_snapshot(pid: Optional[int]) -> Optional[Dict[str, Any]]:
    if not pid:
        return None

    process_query_limited_information = 0x1000
    process_vm_read = 0x0010
    handle = ctypes.windll.kernel32.OpenProcess(
        process_query_limited_information | process_vm_read,
        False,
        int(pid),
    )
    if not handle:
        return None

    try:
        counters = PROCESS_MEMORY_COUNTERS_EX()
        counters.cb = ctypes.sizeof(PROCESS_MEMORY_COUNTERS_EX)
        ok = ctypes.windll.psapi.GetProcessMemoryInfo(
            handle,
            ctypes.byref(counters),
            counters.cb,
        )
        if not ok:
            return None

        def mb(value: int) -> float:
            return round(float(value) / (1024.0 * 1024.0), 3)

        return {
            "pid": int(pid),
            "working_set_mb": mb(counters.WorkingSetSize),
            "peak_working_set_mb": mb(counters.PeakWorkingSetSize),
            "private_usage_mb": mb(counters.PrivateUsage),
            "pagefile_usage_mb": mb(counters.PagefileUsage),
            "peak_pagefile_usage_mb": mb(counters.PeakPagefileUsage),
        }
    except Exception:
        return None
    finally:
        try:
            ctypes.windll.kernel32.CloseHandle(handle)
        except Exception:
            pass


def is_cell_hidden(sheet: Any, row: int, col: int) -> bool:
    """
    Excel cells are hidden through their worksheet, row and/or column.
    If the COM check itself fails, treat the cell as visible so real mismatches
    are not accidentally suppressed.
    """
    try:
        if int(sheet.Visible) != -1:  # -1 = xlSheetVisible
            return True
    except Exception:
        pass
    try:
        if bool(sheet.Rows(row).Hidden):
            return True
    except Exception:
        pass
    try:
        if bool(sheet.Columns(col).Hidden):
            return True
    except Exception:
        pass
    return False


def sheet_visible_state(sheet: Any) -> Optional[int]:
    try:
        return int(sheet.Visible)
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


def repair_originals_with_python_excel_shrink(
    originals_dir: Path,
    repair_dir: Path,
) -> None:
    """
    Materialise workbook.xml-only repaired originals via the Python Excel Shrink.
    These repaired files are then used as the verify baseline.
    """
    project_root = Path(__file__).resolve().parents[2]
    excel_shrink_script = project_root / "scripts" / "excel_shrink" / "excel_shrink.py"
    if not excel_shrink_script.exists():
        raise FileNotFoundError(f"Excel Shrink script not found: {excel_shrink_script}")

    originals_count = len(list_xlsx_files(originals_dir))
    repair_dir.mkdir(parents=True, exist_ok=True)
    command = [
        sys.executable,
        str(excel_shrink_script),
        str(originals_dir),
        str(repair_dir),
        "--only-clean-workbook",
        "--disable-multiprocessing",
    ]
    logging.info(
        "Repairing workbook.xml for %d original files into %s using Python Excel Shrink",
        originals_count,
        repair_dir,
    )
    subprocess.run(command, cwd=str(project_root), check=True)


def _sample_rust_shrink_process_memory(started_at: datetime.datetime) -> Dict[str, Any]:
    cutoff = started_at.isoformat()
    command = [
        "powershell.exe",
        "-NoProfile",
        "-Command",
        (
            "$cutoff=[datetime]'%s'; "
            "Get-Process python,cargo,excel-shrink-cli -ErrorAction SilentlyContinue | "
            "Where-Object { $_.StartTime -ge $cutoff } | "
            "Select-Object Id,ProcessName,WorkingSet64,PrivateMemorySize64 | "
            "ConvertTo-Json -Depth 3"
        )
        % cutoff,
    ]
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=5,
        )
        if completed.returncode != 0 or not completed.stdout.strip():
            return {"working_set_bytes": 0, "private_bytes": 0, "processes": []}
        raw = json.loads(completed.stdout)
        processes = raw if isinstance(raw, list) else [raw]
        cleaned: List[Dict[str, Any]] = []
        working_set = 0
        private = 0
        for proc in processes:
            ws = int(proc.get("WorkingSet64") or 0)
            pm = int(proc.get("PrivateMemorySize64") or 0)
            working_set += ws
            private += pm
            cleaned.append(
                {
                    "pid": int(proc.get("Id") or 0),
                    "process_name": str(proc.get("ProcessName") or ""),
                    "working_set_mb": round(ws / (1024.0 * 1024.0), 3),
                    "private_memory_mb": round(pm / (1024.0 * 1024.0), 3),
                }
            )
        return {
            "working_set_bytes": working_set,
            "private_bytes": private,
            "processes": cleaned,
        }
    except Exception:
        return {"working_set_bytes": 0, "private_bytes": 0, "processes": []}


def shrink_workbook_with_rust_for_verify(
    original_path: Path,
    output_dir: Path,
    *,
    analysis_dir: Path,
    log_dir: Path,
    max_active_workbooks: int,
    writer_workers: int,
    remove_unreferenced_hidden_content_cells: bool,
) -> Tuple[Optional[Path], Dict[str, Any]]:
    project_root = Path(__file__).resolve().parents[2]
    shrink_runner = project_root / "scripts" / "excel_shrink_rust" / "run_excel_shrink_rust.py"
    rust_cli = (
        project_root
        / "external"
        / "excel-shrink-rust"
        / "target"
        / "release"
        / "excel-shrink-cli.exe"
    )
    if not rust_cli.exists() and not shrink_runner.exists():
        raise FileNotFoundError(f"Rust shrink runner not found: {shrink_runner}")

    output_dir.mkdir(parents=True, exist_ok=True)
    analysis_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", original_path.stem)[:120]
    digest = hashlib.sha256(str(original_path.resolve()).encode("utf-8")).hexdigest()[:12]
    analysis_csv = analysis_dir / f"{safe_stem}_{digest}.csv"
    log_path = log_dir / f"{safe_stem}_{digest}.log"
    expected_output = output_dir / original_path.name

    rust_args = [
        str(original_path),
        str(output_dir),
        "--force-overwrite",
        "--max-active-workbooks",
        str(max(1, int(max_active_workbooks))),
        "--writer-workers",
        str(max(1, int(writer_workers))),
        "--analysis-csv",
        str(analysis_csv),
        "--console-log-level",
        "warn",
        "--file-log-level",
        "info",
    ]
    if remove_unreferenced_hidden_content_cells:
        rust_args.append("--remove-unreferenced-hidden-content-cells")

    if rust_cli.exists():
        command = [str(rust_cli), *rust_args]
        command_kind = "release_exe"
    else:
        command = [sys.executable, str(shrink_runner), *rust_args]
        command_kind = "python_runner_fallback"

    started_at = datetime.datetime.now(datetime.timezone.utc).astimezone()
    peak_working_set = 0
    peak_private = 0
    peak_processes: List[Dict[str, Any]] = []

    with log_path.open("w", encoding="utf-8", errors="replace") as log_file:
        proc = subprocess.Popen(
            command,
            cwd=str(project_root),
            stdout=log_file,
            stderr=subprocess.STDOUT,
            text=True,
        )
        while proc.poll() is None:
            memory = process_memory_snapshot(proc.pid)
            if memory:
                working_set = int(
                    float(memory.get("working_set_mb") or 0.0) * 1024.0 * 1024.0
                )
                private = int(
                    float(memory.get("private_usage_mb") or 0.0) * 1024.0 * 1024.0
                )
                if working_set > peak_working_set:
                    peak_working_set = working_set
                    peak_processes = [
                        {
                            "pid": proc.pid,
                            "process_name": Path(command[0]).stem,
                            "working_set_mb": memory.get("working_set_mb"),
                            "private_memory_mb": memory.get("private_usage_mb"),
                        }
                    ]
                if private > peak_private:
                    peak_private = private
            time.sleep(0.1)
        exit_code = proc.wait()

    finished_at = datetime.datetime.now(datetime.timezone.utc).astimezone()
    output_path = expected_output if expected_output.exists() else None
    metrics: Dict[str, Any] = {
        "enabled": True,
        "status": "ok" if exit_code == 0 and output_path is not None else "error",
        "input": str(original_path),
        "output": str(expected_output),
        "command_kind": command_kind,
        "analysis_csv": str(analysis_csv),
        "log_file": str(log_path),
        "started_at": started_at.isoformat(),
        "finished_at": finished_at.isoformat(),
        "elapsed_sec": round((finished_at - started_at).total_seconds(), 6),
        "exit_code": exit_code,
        "peak_working_set_mb": round(peak_working_set / (1024.0 * 1024.0), 3),
        "peak_private_memory_mb": round(peak_private / (1024.0 * 1024.0), 3),
        "peak_process_sample": peak_processes,
    }
    if output_path is None:
        metrics["reason"] = "shrunk_output_missing"
    return output_path, metrics


def attach_shrink_performance(
    result: FileResult,
    shrink_metrics: Optional[Dict[str, Any]],
) -> FileResult:
    if not shrink_metrics:
        return result

    performance = result.performance or {"enabled": True}
    performance["shrink"] = shrink_metrics

    timings = performance.setdefault("timings_sec", {})
    derived = performance.setdefault("derived_timings_sec", {})
    shrink_elapsed = float(shrink_metrics.get("elapsed_sec") or 0.0)
    open_original = float(timings.get("open_original_a") or 0.0)
    open_shrunk = float(timings.get("open_shrunk_b") or 0.0)
    recalc_original = float(timings.get("recalc_original_a") or 0.0)
    recalc_shrunk = float(timings.get("recalc_shrunk_b") or 0.0)

    derived["excel_open_original_only"] = round(open_original, 6)
    derived["excel_open_shrunk_only"] = round(open_shrunk, 6)
    derived["excel_open_recalc_original"] = round(open_original + recalc_original, 6)
    derived["excel_open_recalc_shrunk"] = round(open_shrunk + recalc_shrunk, 6)
    derived["rust_shrink_only"] = round(shrink_elapsed, 6)
    derived["rust_shrink_plus_excel_open_shrunk"] = round(
        shrink_elapsed + open_shrunk,
        6,
    )
    derived["rust_shrink_plus_excel_open_recalc_shrunk"] = round(
        shrink_elapsed + open_shrunk + recalc_shrunk,
        6,
    )

    memory = performance.setdefault("memory", {})
    memory["rust_shrink_peak_working_set_mb"] = shrink_metrics.get(
        "peak_working_set_mb"
    )
    memory["rust_shrink_peak_private_memory_mb"] = shrink_metrics.get(
        "peak_private_memory_mb"
    )
    result.performance = performance
    return result


# ----------------------------
# Path length handling (Windows MAX_PATH workaround)
# ----------------------------
MAX_EXCEL_PATH_LEN = 180  # as requested


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


def clean_workbook_xml_for_excel(workbook_xml: bytes) -> Tuple[bytes, int]:
    """
    Remove duplicate defined names that can make Excel refuse to open workbooks.
    This mirrors excel_shrink.py --only-clean-workbook for the verification path.
    """
    block_match = _DEFINED_NAMES_RE.search(workbook_xml)
    if block_match is None:
        return workbook_xml, 0

    defined_names_content = block_match.group(2)
    duplicate_pairs = [
        (b"Print_Area", b"_xlnm.Print_Area"),
        (b"Print_Titles", b"_xlnm.Print_Titles"),
        (b"_FilterDatabase", b"_xlnm._FilterDatabase"),
        (b"Database", b"_xlnm.Database"),
    ]

    seen: Dict[bytes, set[bytes]] = {}
    for name, local_sheet_id in _DEFINED_NAME_START_RE.findall(defined_names_content):
        seen.setdefault(name, set()).add(local_sheet_id or b"")

    names_to_delete: set[Tuple[bytes, bytes]] = set()
    for plain_name, xlnm_name in duplicate_pairs:
        plain_sheet_ids = seen.get(plain_name, set())
        xlnm_sheet_ids = seen.get(xlnm_name, set())
        for local_sheet_id in plain_sheet_ids & xlnm_sheet_ids:
            names_to_delete.add((plain_name, local_sheet_id))

    if not names_to_delete:
        return workbook_xml, 0

    removed = 0

    def remove_duplicate_defined_name(match: re.Match[bytes]) -> bytes:
        nonlocal removed
        name = match.group(1)
        local_sheet_id = match.group(2) or b""
        if (name, local_sheet_id) in names_to_delete:
            removed += 1
            return b""
        return match.group(0)

    cleaned_defined_names = _DEFINED_NAME_NODE_RE.sub(
        remove_duplicate_defined_name, defined_names_content
    )
    cleaned_workbook = (
        workbook_xml[: block_match.start(2)]
        + cleaned_defined_names
        + workbook_xml[block_match.end(2) :]
    )
    return cleaned_workbook, removed


def repair_workbook_xml_copy_for_excel(xlsx_path: Path, temp_dir: Path) -> Path:
    """
    If workbook.xml contains duplicate defined names, write a repaired temporary XLSX.
    The source file is never modified.
    """
    temp_dir.mkdir(parents=True, exist_ok=True)
    abs_str = _abs_path_str(xlsx_path)
    repaired_path = temp_dir / _temp_name_from_abs_path(
        abs_str + "|workbook_xml_repaired", suffix=xlsx_path.suffix or ".xlsx"
    )

    if repaired_path.exists():
        return repaired_path

    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            try:
                workbook_xml = zin.read(_WORKBOOK_XML)
            except KeyError:
                return xlsx_path

            cleaned_workbook_xml, removed = clean_workbook_xml_for_excel(workbook_xml)
            if removed == 0:
                return xlsx_path

            logging.info(
                "Repairing workbook.xml before Excel open: %s removed_defined_names=%d",
                xlsx_path.name,
                removed,
            )
            with zipfile.ZipFile(repaired_path, "w") as zout:
                for info in zin.infolist():
                    if info.filename == _WORKBOOK_XML:
                        data = cleaned_workbook_xml
                    else:
                        data = zin.read(info)
                    zout.writestr(info, data)
    except Exception:
        if repaired_path.exists():
            try:
                repaired_path.unlink()
            except Exception:
                pass
        raise

    return repaired_path


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
        excel.ErrorCheckingOptions.BackgroundChecking = False
    except Exception:
        pass
    try:
        excel.AutomationSecurity = 3  # disable macros
    except Exception:
        pass

    return excel


def close_excel_app(excel: Any) -> None:
    try:
        if excel is not None:
            excel.Quit()
    except Exception:
        pass
    try:
        del excel
    except Exception:
        pass


def list_excel_process_ids() -> set[int]:
    def from_powershell_get_process() -> set[int]:
        try:
            completed = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    "Get-Process EXCEL -ErrorAction SilentlyContinue | "
                    "ForEach-Object { $_.Id }",
                ],
                check=False,
                capture_output=True,
                text=True,
                timeout=15,
            )
        except Exception as e:
            logging.warning("Could not list EXCEL.EXE processes via PowerShell: %s", e)
            return set()

        if completed.returncode != 0:
            logging.warning(
                "PowerShell Get-Process failed while listing EXCEL.EXE processes: %s",
                (completed.stderr or completed.stdout or "").strip(),
            )
            return set()

        pids: set[int] = set()
        for line in completed.stdout.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                pids.add(int(line))
            except ValueError:
                continue
        return pids

    try:
        completed = subprocess.run(
            [
                "tasklist",
                "/FI",
                "IMAGENAME eq EXCEL.EXE",
                "/FO",
                "CSV",
                "/NH",
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=15,
        )
    except Exception as e:
        logging.warning("Could not list EXCEL.EXE processes: %s", e)
        return from_powershell_get_process()

    if completed.returncode != 0:
        logging.warning(
            "tasklist failed while listing EXCEL.EXE processes: %s",
            (completed.stderr or completed.stdout or "").strip(),
        )
        return from_powershell_get_process()

    pids: set[int] = set()
    for row in csv.reader(completed.stdout.splitlines()):
        if len(row) < 2:
            continue
        image_name = row[0].strip('"').lower()
        if image_name != "excel.exe":
            continue
        try:
            pids.add(int(row[1].strip('"')))
        except ValueError:
            continue
    return pids


def list_visible_window_process_ids() -> set[int]:
    visible_pids: set[int] = set()

    def collect(hwnd: int, _extra: Any) -> bool:
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid:
                visible_pids.add(int(pid))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(collect, None)
    except Exception as e:
        logging.warning("Could not enumerate visible windows: %s", e)

    return visible_pids


def kill_hung_excel_processes(
    *,
    baseline_pids: set[int],
    reason: str,
    include_visible: bool = False,
) -> int:
    """
    Kill Excel processes that were spawned after this verify run started.
    Visible windows are protected by default so a user's normal Excel session is
    not closed while cleaning up orphaned COM automation instances.
    """
    excel_pids = list_excel_process_ids()
    if not excel_pids:
        logging.info("No EXCEL.EXE processes found for cleanup after %s", reason)
        return 0

    visible_pids = set() if include_visible else list_visible_window_process_ids()
    candidates = sorted(excel_pids - set(baseline_pids) - visible_pids)
    if not candidates:
        logging.info(
            "No hidden post-start EXCEL.EXE processes to kill after %s "
            "(excel=%d baseline=%d visible=%d)",
            reason,
            len(excel_pids),
            len(baseline_pids),
            len(visible_pids),
        )
        return 0

    killed = 0
    for pid in candidates:
        try:
            completed = subprocess.run(
                ["taskkill", "/PID", str(pid), "/F", "/T"],
                check=False,
                capture_output=True,
                text=True,
                timeout=15,
            )
            if completed.returncode == 0:
                killed += 1
                logging.warning(
                    "Killed hidden EXCEL.EXE pid=%s after %s", pid, reason
                )
            else:
                logging.warning(
                    "Could not kill EXCEL.EXE pid=%s after %s: %s",
                    pid,
                    reason,
                    (completed.stderr or completed.stdout or "").strip(),
                )
        except Exception as e:
            logging.warning(
                "Exception while killing EXCEL.EXE pid=%s after %s: %s",
                pid,
                reason,
                e,
            )

    if killed:
        time.sleep(2.0)
    return killed


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


def is_excel_rpc_unavailable(reason: Optional[str]) -> bool:
    if not reason:
        return False
    return (
        "RPC server is unavailable" in reason
        or "-2147023174" in reason
        or "Object disconnected" in reason
    )


def is_temp_permission_error(reason: Optional[str]) -> bool:
    if not reason:
        return False
    return "PermissionError" in reason and "xlsx_compare_" in reason


def is_retryable_excel_error(reason: Optional[str]) -> bool:
    if not reason:
        return False
    return (
        is_excel_rpc_unavailable(reason)
        or "Open method of Workbooks class failed" in reason
        or "-2146827284" in reason
        or "The remote procedure call failed" in reason
        or "-2147023170" in reason
        or "Not enough memory resources are available" in reason
        or "-2147024882" in reason
        or "The paging file is too small" in reason
        or "-2147023441" in reason
        or "Server execution failed" in reason
        or "-2146959355" in reason
        or "Property 'Excel.Application.Visible' can not be set" in reason
        or "Call was rejected by callee" in reason
        or "The message filter indicated" in reason
    )


@dataclasses.dataclass
class Diff:
    sheet: str
    cell: str
    original: Any
    shrunk: Any
    original_formula: Optional[str] = None
    shrunk_formula: Optional[str] = None
    category: Optional[str] = None


@dataclasses.dataclass
class FileResult:
    filename: str
    status: str  # "same" | "different" | "error" | "skipped"
    reason: Optional[str] = None
    diff_count: int = 0
    volatile_baseline_diff_count: int = 0
    cleanup_diff_count: int = 0
    volatile_repeated_diff_count: int = 0
    hidden_ignored_diff_count: int = 0
    diffs: Optional[List[Diff]] = None
    visual: Optional[Dict[str, Any]] = None
    performance: Optional[Dict[str, Any]] = None
    elapsed_sec: float = 0.0


def file_result_to_dict(result: FileResult) -> Dict[str, Any]:
    return dataclasses.asdict(result)


def file_result_from_dict(data: Dict[str, Any]) -> FileResult:
    diffs = [
        Diff(
            sheet=d.get("sheet", ""),
            cell=d.get("cell", ""),
            original=d.get("original"),
            shrunk=d.get("shrunk"),
            original_formula=d.get("original_formula"),
            shrunk_formula=d.get("shrunk_formula"),
            category=d.get("category"),
        )
        for d in (data.get("diffs") or [])
    ]
    return FileResult(
        filename=data.get("filename", ""),
        status=data.get("status", "error"),
        reason=data.get("reason"),
        diff_count=int(data.get("diff_count") or 0),
        volatile_baseline_diff_count=int(
            data.get("volatile_baseline_diff_count") or 0
        ),
        cleanup_diff_count=int(data.get("cleanup_diff_count") or 0),
        volatile_repeated_diff_count=int(
            data.get("volatile_repeated_diff_count") or 0
        ),
        hidden_ignored_diff_count=int(data.get("hidden_ignored_diff_count") or 0),
        diffs=diffs,
        visual=data.get("visual"),
        performance=data.get("performance"),
        elapsed_sec=float(data.get("elapsed_sec") or 0.0),
    )


def progress_journal_path(out_path: Path) -> Path:
    return out_path.with_name(out_path.name + ".progress.jsonl")


def append_progress_result(progress_path: Path, result: FileResult) -> None:
    progress_path.parent.mkdir(parents=True, exist_ok=True)
    with progress_path.open("a", encoding="utf-8") as f:
        json.dump(file_result_to_dict(result), f, ensure_ascii=False)
        f.write("\n")
        f.flush()


def load_previous_completed_results(out_path: Path) -> Dict[str, FileResult]:
    """
    Load resumable per-workbook results. The JSONL journal is written after each
    workbook, so it survives crashes that happen before the final report write.
    """
    progress_path = progress_journal_path(out_path)
    results: Dict[str, FileResult] = {}

    if progress_path.exists():
        try:
            with progress_path.open("r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        result = file_result_from_dict(json.loads(line))
                    except Exception as e:
                        logging.warning(
                            "Could not read progress entry from %s: %s",
                            progress_path,
                            e,
                        )
                        continue
                    if result.filename:
                        results[result.filename] = result
        except Exception as e:
            logging.warning("Could not read progress journal %s: %s", progress_path, e)

    if results or not out_path.exists():
        return results

    try:
        with out_path.open("r", encoding="utf-8") as f:
            report = json.load(f)
    except Exception as e:
        logging.warning("Could not read existing report %s: %s", out_path, e)
        return results

    for filename in report.get("same_files", []) or []:
        if isinstance(filename, str):
            results[filename] = FileResult(
                filename=filename,
                status="same",
                reason="already_same_in_existing_report",
                diffs=[],
                diff_count=0,
                elapsed_sec=0.0,
            )

    for item in report.get("skipped_files", []) or []:
        filename = item.get("filename") if isinstance(item, dict) else None
        if isinstance(filename, str):
            results[filename] = FileResult(
                filename=filename,
                status="skipped",
                reason=item.get("reason"),
                diffs=[],
                diff_count=0,
                elapsed_sec=float(item.get("elapsed_sec") or 0.0),
            )

    for item in report.get("different_files", []) or []:
        filename = item.get("filename") if isinstance(item, dict) else None
        if isinstance(filename, str):
            results[filename] = file_result_from_dict(
                {
                    "filename": filename,
                    "status": "different",
                    **item,
                }
            )

    for item in report.get("errors", []) or []:
        filename = item.get("filename") if isinstance(item, dict) else None
        if isinstance(filename, str):
            results[filename] = FileResult(
                filename=filename,
                status="error",
                reason=item.get("reason"),
                diffs=[],
                diff_count=0,
                elapsed_sec=float(item.get("elapsed_sec") or 0.0),
            )

    visual_by_name = {
        item.get("filename"): {k: v for k, v in item.items() if k != "filename"}
        for item in (report.get("visual_files", []) or [])
        if isinstance(item, dict) and isinstance(item.get("filename"), str)
    }
    performance_by_name = {
        item.get("filename"): {
            k: v for k, v in item.items() if k not in ("filename", "status")
        }
        for item in (report.get("performance_files", []) or [])
        if isinstance(item, dict) and isinstance(item.get("filename"), str)
    }
    for filename, result in results.items():
        result.visual = visual_by_name.get(filename)
        result.performance = performance_by_name.get(filename)

    return results


def import_pillow_for_visual_compare() -> Tuple[Any, Any, Any]:
    try:
        from PIL import Image, ImageChops, ImageGrab

        return Image, ImageChops, ImageGrab
    except ImportError as e:
        raise RuntimeError(
            "Visual compare requires Pillow. Install it with: pip install pillow"
        ) from e


EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLS = 16_384
VISUAL_ROW_PADDING = 15


def get_sheet_content_bounds_for_visual(sheet: Any) -> Dict[str, int]:
    """
    Return bounds for real cell contents/formulas only. UsedRange also expands for
    formatting, which would reintroduce grid-only screenshot areas after shrinking.
    """
    try:
        first_cell = sheet.Cells(1, 1)
        last_row_cell = sheet.Cells.Find(
            What="*",
            After=first_cell,
            LookIn=-4123,  # xlFormulas
            LookAt=2,  # xlPart
            SearchOrder=1,  # xlByRows
            SearchDirection=2,  # xlPrevious
            MatchCase=False,
        )
        last_col_cell = sheet.Cells.Find(
            What="*",
            After=first_cell,
            LookIn=-4123,
            LookAt=2,
            SearchOrder=2,  # xlByColumns
            SearchDirection=2,
            MatchCase=False,
        )
        if last_row_cell is None or last_col_cell is None:
            return {"min_row": 1, "min_col": 1, "max_row": 1, "max_col": 1}
        return {
            "min_row": 1,
            "min_col": 1,
            "max_row": max(1, int(last_row_cell.Row)),
            "max_col": max(1, int(last_col_cell.Column)),
        }
    except Exception:
        bounds = get_used_bounds(sheet)
        if bounds is None:
            return {"min_row": 1, "min_col": 1, "max_row": 1, "max_col": 1}
        return {
            "min_row": max(1, int(bounds[0])),
            "min_col": max(1, int(bounds[1])),
            "max_row": max(1, int(bounds[2])),
            "max_col": max(1, int(bounds[3])),
        }


def build_visual_regions(max_row: int, max_col: int, rows: int, cols: int) -> List[Dict[str, int]]:
    rows = max(1, int(rows))
    cols = max(1, int(cols))
    max_row = max(1, min(EXCEL_MAX_ROWS, int(max_row)))
    max_col = max(1, min(EXCEL_MAX_COLS, int(max_col)))

    def region_ending_at(end_row: int, end_col: int) -> Dict[str, int]:
        row2 = max(1, min(EXCEL_MAX_ROWS, int(end_row)))
        col2 = max(1, min(EXCEL_MAX_COLS, int(end_col)))
        row1 = max(1, row2 - rows + 1)
        col1 = max(1, col2 - cols + 1)
        return {"row1": row1, "col1": col1, "row2": row2, "col2": col2}

    regions = [
        {"row1": 1, "col1": 1, "row2": min(rows, max_row), "col2": min(cols, max_col)}
    ]
    if max_row > rows:
        regions.append(region_ending_at(max_row + VISUAL_ROW_PADDING, min(cols, max_col)))
    if max_col > cols:
        regions.append(region_ending_at(min(rows, max_row), max_col))
    if max_row > rows and max_col > cols:
        regions.append(
            region_ending_at(
                max_row + VISUAL_ROW_PADDING,
                max_col,
            )
        )

    return merge_overlapping_visual_regions(regions)


def visual_regions_overlap(a: Dict[str, int], b: Dict[str, int]) -> bool:
    return (
        a["row1"] <= b["row2"]
        and b["row1"] <= a["row2"]
        and a["col1"] <= b["col2"]
        and b["col1"] <= a["col2"]
    )


def merge_overlapping_visual_regions(regions: List[Dict[str, int]]) -> List[Dict[str, int]]:
    merged: List[Dict[str, int]] = []
    for region in regions:
        current = dict(region)
        changed = True
        while changed:
            changed = False
            next_regions: List[Dict[str, int]] = []
            for existing in merged:
                if visual_regions_overlap(current, existing):
                    current = {
                        "row1": min(current["row1"], existing["row1"]),
                        "col1": min(current["col1"], existing["col1"]),
                        "row2": max(current["row2"], existing["row2"]),
                        "col2": max(current["col2"], existing["col2"]),
                    }
                    changed = True
                else:
                    next_regions.append(existing)
            merged = next_regions
        merged.append(current)

    for idx, region in enumerate(merged, start=1):
        region["index"] = idx
    return merged


def capture_sheet_range_picture(
    sheet: Any,
    row1: int,
    col1: int,
    row2: int,
    col2: int,
    retry_count: int,
    retry_sleep_sec: float,
) -> Any:
    _, _, image_grab = import_pillow_for_visual_compare()
    app = sheet.Application
    workbook = sheet.Parent
    workbook.Activate()
    sheet.Activate()
    try:
        app.DisplayFullScreen = True
        app.WindowState = -4137  # xlMaximized
    except Exception:
        pass
    try:
        app.ActiveWindow.Zoom = 100
        app.ActiveWindow.ScrollRow = max(1, int(row1))
        app.ActiveWindow.ScrollColumn = max(1, int(col1))
        app.ActiveWindow.DisplayGridlines = False
    except Exception:
        pass

    rng = sheet.Range(sheet.Cells(row1, col1), sheet.Cells(row2, col2))

    last_error: Optional[str] = None
    for _ in range(max(1, retry_count)):
        try:
            rng.CopyPicture(Appearance=1, Format=2)
            time.sleep(retry_sleep_sec)
            img = image_grab.grabclipboard()
            if img is not None and hasattr(img, "convert") and hasattr(img, "copy"):
                return img.copy()
            last_error = f"clipboard did not contain an image: {type(img).__name__}"
        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            time.sleep(retry_sleep_sec)

    raise RuntimeError(last_error or "failed to capture Excel picture")


def compare_visual_images(
    image_a: Any,
    image_b: Any,
    pixel_tolerance: int,
) -> Dict[str, Any]:
    image, image_chops, _ = import_pillow_for_visual_compare()
    a = image_a.convert("RGB")
    b = image_b.convert("RGB")

    size_a = list(a.size)
    size_b = list(b.size)
    max_width = max(a.size[0], b.size[0])
    max_height = max(a.size[1], b.size[1])
    if a.size != (max_width, max_height):
        canvas = image.new("RGB", (max_width, max_height), "white")
        canvas.paste(a, (0, 0))
        a = canvas
    if b.size != (max_width, max_height):
        canvas = image.new("RGB", (max_width, max_height), "white")
        canvas.paste(b, (0, 0))
        b = canvas

    diff = image_chops.difference(a, b)
    bbox = diff.getbbox()
    total_pixels = max_width * max_height
    if bbox is None:
        return {
            "same": True,
            "size_original": size_a,
            "size_shrunk": size_b,
            "compared_size": [max_width, max_height],
            "different_pixels": 0,
            "different_pixel_ratio": 0.0,
            "mean_abs_channel_diff": 0.0,
            "max_channel_diff": 0,
            "bbox": None,
        }

    pixels = diff.getdata()
    different_pixels = 0
    channel_sum = 0
    max_channel_diff = 0
    tolerance = max(0, int(pixel_tolerance))
    for r, g, b_channel in pixels:
        pixel_max = max(r, g, b_channel)
        if pixel_max > tolerance:
            different_pixels += 1
        channel_sum += r + g + b_channel
        if pixel_max > max_channel_diff:
            max_channel_diff = pixel_max

    return {
        "same": different_pixels == 0,
        "size_original": size_a,
        "size_shrunk": size_b,
        "compared_size": [max_width, max_height],
        "different_pixels": different_pixels,
        "different_pixel_ratio": (
            float(different_pixels) / float(total_pixels) if total_pixels else 0.0
        ),
        "mean_abs_channel_diff": (
            float(channel_sum) / float(total_pixels * 3) if total_pixels else 0.0
        ),
        "max_channel_diff": max_channel_diff,
        "bbox": list(bbox),
    }


def safe_visual_filename(filename: str, sheet_name: str, side: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", f"{filename}_{sheet_name}_{side}")
    return safe[:180] + ".png"


def compare_workbook_visuals(
    wb_a: Any,
    wb_b: Any,
    filename: str,
    rows: int,
    cols: int,
    max_sheets: int,
    pixel_tolerance: int,
    screenshot_dir: Optional[Path],
    retry_count: int,
    retry_sleep_sec: float,
) -> Dict[str, Any]:
    t0 = time.time()
    count_a = int(wb_a.Worksheets.Count)
    count_b = int(wb_b.Worksheets.Count)
    sheet_count = min(count_a, count_b)
    if max_sheets > 0:
        sheet_count = min(sheet_count, max_sheets)

    sheet_results: List[Dict[str, Any]] = []
    total_different_pixels = 0
    total_pixels = 0
    error_count = 0
    skipped_hidden_sheets = 0

    if screenshot_dir is not None:
        screenshot_dir.mkdir(parents=True, exist_ok=True)

    for i in range(1, sheet_count + 1):
        sheet_t0 = time.perf_counter()
        sheet_a = wb_a.Worksheets(i)
        sheet_b = wb_b.Worksheets(i)
        sheet_name = str(sheet_a.Name)
        result: Dict[str, Any] = {"sheet": sheet_name, "status": "ok"}
        try:
            logging.info(
                "SHEET_START visual workbook=%s sheet=%d/%d name=%s",
                filename,
                i,
                sheet_count,
                sheet_name,
            )
            visible_a = sheet_visible_state(sheet_a)
            visible_b = sheet_visible_state(sheet_b)
            if visible_a != visible_b:
                result.update(
                    {
                        "status": "different",
                        "same": False,
                        "reason": (
                            "sheet_visibility_mismatch: "
                            f"original={visible_a}, shrunk={visible_b}"
                        ),
                    }
                )
                continue

            if visible_a is not None and visible_a != -1:  # -1 = xlSheetVisible
                skipped_hidden_sheets += 1
                result.update(
                    {
                        "status": "skipped_hidden_sheet",
                        "same": True,
                        "reason": f"sheet_hidden: visible_state={visible_a}",
                    }
                )
                continue

            bounds_a = get_sheet_content_bounds_for_visual(sheet_a)
            bounds_b = get_sheet_content_bounds_for_visual(sheet_b)
            max_row = max(int(bounds_a["max_row"]), int(bounds_b["max_row"]))
            max_col = max(int(bounds_a["max_col"]), int(bounds_b["max_col"]))
            regions = build_visual_regions(max_row, max_col, rows, cols)

            result["content_bounds_original"] = bounds_a
            result["content_bounds_shrunk"] = bounds_b
            result["regions"] = []

            sheet_different_pixels = 0
            sheet_total_pixels = 0
            sheet_same = True
            sheet_max_channel_diff = 0
            sheet_channel_weighted_sum = 0.0
            sheet_channel_weighted_pixels = 0

            for region in regions:
                region_label = (
                    f"region{region['index']}_"
                    f"r{region['row1']}-{region['row2']}_"
                    f"c{region['col1']}-{region['col2']}"
                )
                img_a = capture_sheet_range_picture(
                    sheet_a,
                    region["row1"],
                    region["col1"],
                    region["row2"],
                    region["col2"],
                    retry_count,
                    retry_sleep_sec,
                )
                img_b = capture_sheet_range_picture(
                    sheet_b,
                    region["row1"],
                    region["col1"],
                    region["row2"],
                    region["col2"],
                    retry_count,
                    retry_sleep_sec,
                )
                metrics = compare_visual_images(img_a, img_b, pixel_tolerance)

                compared_size = metrics.get("compared_size") or [0, 0]
                region_pixels = int(compared_size[0]) * int(compared_size[1])
                region_diff_pixels = int(metrics.get("different_pixels") or 0)
                sheet_total_pixels += region_pixels
                sheet_different_pixels += region_diff_pixels
                sheet_same = sheet_same and bool(metrics.get("same"))
                sheet_max_channel_diff = max(
                    sheet_max_channel_diff, int(metrics.get("max_channel_diff") or 0)
                )
                sheet_channel_weighted_sum += float(
                    metrics.get("mean_abs_channel_diff") or 0.0
                ) * float(region_pixels)
                sheet_channel_weighted_pixels += region_pixels

                region_result = {
                    "label": region_label,
                    "row1": region["row1"],
                    "col1": region["col1"],
                    "row2": region["row2"],
                    "col2": region["col2"],
                    **metrics,
                }

                if screenshot_dir is not None:
                    original_path = screenshot_dir / safe_visual_filename(
                        filename, f"{sheet_name}_{region_label}", "original"
                    )
                    shrunk_path = screenshot_dir / safe_visual_filename(
                        filename, f"{sheet_name}_{region_label}", "shrunk"
                    )
                    img_a.save(original_path)
                    img_b.save(shrunk_path)
                    region_result["screenshot_original"] = str(original_path)
                    region_result["screenshot_shrunk"] = str(shrunk_path)

                    if len(regions) == 1:
                        result["screenshot_original"] = str(original_path)
                        result["screenshot_shrunk"] = str(shrunk_path)

                result["regions"].append(region_result)

            total_pixels += sheet_total_pixels
            total_different_pixels += sheet_different_pixels
            result.update(
                {
                    "same": sheet_same,
                    "different_pixels": sheet_different_pixels,
                    "different_pixel_ratio": (
                        float(sheet_different_pixels) / float(sheet_total_pixels)
                        if sheet_total_pixels
                        else 0.0
                    ),
                    "mean_abs_channel_diff": (
                        sheet_channel_weighted_sum
                        / float(sheet_channel_weighted_pixels)
                        if sheet_channel_weighted_pixels
                        else 0.0
                    ),
                    "max_channel_diff": sheet_max_channel_diff,
                    "bbox": None,
                    "status": "ok" if sheet_same else "different",
                }
            )
        except Exception as e:
            error_count += 1
            result.update(
                {
                    "status": "error",
                    "same": False,
                    "reason": f"{type(e).__name__}: {e}",
                }
            )
        finally:
            elapsed = time.perf_counter() - sheet_t0
            result["elapsed_sec"] = round(elapsed, 6)
            logging.info(
                "SHEET_DONE visual workbook=%s sheet=%d/%d name=%s status=%s elapsed=%.3fs",
                filename,
                i,
                sheet_count,
                sheet_name,
                result.get("status"),
                elapsed,
            )
            sheet_results.append(result)
            try:
                del sheet_a, sheet_b
            except Exception:
                pass

    same = error_count == 0 and all(bool(s.get("same")) for s in sheet_results)
    return {
        "enabled": True,
        "status": "same" if same else ("error" if error_count else "different"),
        "same": same,
        "scope": {
            "region_rows": rows,
            "region_cols": cols,
            "row_padding": VISUAL_ROW_PADDING,
            "col_padding": 0,
            "zoom": 100,
            "sheets_compared": sheet_count,
            "max_sheets": max_sheets,
            "pixel_tolerance": pixel_tolerance,
        },
        "summary": {
            "different_pixels": total_different_pixels,
            "different_pixel_ratio": (
                float(total_different_pixels) / float(total_pixels)
                if total_pixels
                else 0.0
            ),
            "errors": error_count,
            "skipped_hidden_sheets": skipped_hidden_sheets,
            "elapsed_sec": time.time() - t0,
        },
        "sheets": sheet_results,
    }


def compare_sheet_values(
    sheet_a: Any,
    sheet_b: Any,
    sheet_name: str,
    row_chunk: int,
    col_chunk: int,
    max_diffs: int,
    *,
    collect_diff_cells: Optional[set[Tuple[str, str]]] = None,
    volatile_cells: Optional[set[Tuple[str, str]]] = None,
) -> Tuple[int, int, int, int, List[Diff]]:

    # if sheet_a.UsedRange.Value2 == sheet_b.UsedRange.Value2:
    #     return 0, []

    bounds_a = get_used_bounds(sheet_a)
    bounds_b = get_used_bounds(sheet_b)

    if bounds_a is None and bounds_b is None:
        return 0, 0, 0, 0, []

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
    volatile_repeated_diff_count = 0
    cleanup_diff_count = 0
    hidden_ignored_diff_count = 0

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

                    cell_addr = a1_addr(row_idx, col_idx)
                    cell_key = (sheet_name, cell_addr)

                    if is_cell_hidden(sheet_a, row_idx, col_idx) or is_cell_hidden(
                        sheet_b, row_idx, col_idx
                    ):
                        hidden_ignored_diff_count += 1
                        continue

                    if collect_diff_cells is not None:
                        collect_diff_cells.add(cell_key)

                    category: Optional[str] = None
                    if volatile_cells is not None:
                        if cell_key in volatile_cells:
                            category = "already_different_in_original"
                            volatile_repeated_diff_count += 1
                        else:
                            category = "changed_after_cleanup"
                            cleanup_diff_count += 1

                    diff_count += 1
                    if len(diffs) < max_diffs:
                        diffs.append(
                            Diff(
                                sheet=sheet_name,
                                cell=cell_addr,
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
                                category=category,
                            )
                        )

            del rng_a, rng_b, va, vb, a2d, b2d
            c = c2 + 1
        r = r2 + 1

    return (
        diff_count,
        volatile_repeated_diff_count,
        cleanup_diff_count,
        hidden_ignored_diff_count,
        diffs,
    )


def compare_workbooks_in_excel(
    excel_a: Any,
    excel_b: Any,
    original_path: Path,
    shrunk_path: Path,
    temp_dir: Path,
    row_chunk: int,
    col_chunk: int,
    max_diffs_per_file: int,
    visual_compare: bool = False,
    visual_rows: int = 60,
    visual_cols: int = 20,
    visual_max_sheets: int = 0,
    visual_pixel_tolerance: int = 0,
    visual_screenshot_dir: Optional[Path] = None,
    visual_retry_count: int = 5,
    visual_retry_sleep_sec: float = 0.25,
) -> FileResult:

    t0 = time.time()
    filename = original_path.name

    wb_a = None
    wb_b = None
    performance: Dict[str, Any] = {
        "enabled": True,
        "timings_sec": {},
        "memory": {},
        "sheet_timings": [],
    }
    try:
        op = repair_workbook_xml_copy_for_excel(original_path, temp_dir)
        sp = repair_workbook_xml_copy_for_excel(shrunk_path, temp_dir)
        op = ensure_short_path_for_excel(op, temp_dir)
        sp = ensure_short_path_for_excel(sp, temp_dir)

        excel_a_pid = excel_process_id(excel_a)
        excel_b_pid = excel_process_id(excel_b)
        performance["excel_pids"] = {
            "original_app": excel_a_pid,
            "comparison_app": excel_b_pid,
        }
        performance["memory"]["original_app_before_open"] = process_memory_snapshot(
            excel_a_pid
        )
        performance["memory"]["comparison_app_before_open"] = process_memory_snapshot(
            excel_b_pid
        )

        open_t0 = time.perf_counter()
        wb_a = open_workbook(excel_a, op)
        performance["timings_sec"]["open_original_a"] = round(
            time.perf_counter() - open_t0, 6
        )
        performance["memory"]["original_app_after_open_original_a"] = (
            process_memory_snapshot(excel_a_pid)
        )

        open_t0 = time.perf_counter()
        wb_b = open_workbook(excel_b, op)
        performance["timings_sec"]["open_original_b"] = round(
            time.perf_counter() - open_t0, 6
        )
        performance["memory"]["comparison_app_after_open_original_b"] = (
            process_memory_snapshot(excel_b_pid)
        )

        recalc_t0 = time.perf_counter()
        excel_a.CalculateFull()
        performance["timings_sec"]["recalc_original_a"] = round(
            time.perf_counter() - recalc_t0, 6
        )
        performance["memory"]["original_app_after_recalc_original_a"] = (
            process_memory_snapshot(excel_a_pid)
        )

        recalc_t0 = time.perf_counter()
        excel_b.CalculateFull()
        performance["timings_sec"]["recalc_original_b"] = round(
            time.perf_counter() - recalc_t0, 6
        )
        performance["memory"]["comparison_app_after_recalc_original_b"] = (
            process_memory_snapshot(excel_b_pid)
        )

        count_a = int(wb_a.Worksheets.Count)
        count_b = int(wb_b.Worksheets.Count)
        if count_a != count_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=(
                    "original_baseline_sheet_count_mismatch: "
                    f"original_a={count_a}, original_b={count_b}"
                ),
                diffs=[],
                diff_count=0,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        names_a = [str(wb_a.Worksheets(i).Name) for i in range(1, count_a + 1)]
        names_b = [str(wb_b.Worksheets(i).Name) for i in range(1, count_b + 1)]
        if names_a != names_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=(
                    "original_baseline_sheet_name_mismatch: "
                    f"original_a={names_a}, original_b={names_b}"
                ),
                diffs=[],
                diff_count=0,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        volatile_cells: set[Tuple[str, str]] = set()
        volatile_baseline_diff_count = 0
        baseline_sheet_b_by_name = {
            str(wb_b.Worksheets(i).Name): wb_b.Worksheets(i)
            for i in range(1, count_b + 1)
        }
        for i in range(1, count_a + 1):
            sheet_a = wb_a.Worksheets(i)
            name = str(sheet_a.Name)
            sheet_b = baseline_sheet_b_by_name[name]

            sheet_t0 = time.perf_counter()
            bounds_a = get_used_bounds(sheet_a)
            bounds_b = get_used_bounds(sheet_b)
            logging.info(
                "SHEET_START volatile_baseline workbook=%s sheet=%d/%d name=%s bounds_original=%s bounds_baseline=%s",
                filename,
                i,
                count_a,
                name,
                bounds_a,
                bounds_b,
            )
            diff_count, _, _, hidden_ignored_count, _ = compare_sheet_values(
                sheet_a,
                sheet_b,
                name,
                row_chunk=row_chunk,
                col_chunk=col_chunk,
                max_diffs=0,
                collect_diff_cells=volatile_cells,
            )
            sheet_elapsed = time.perf_counter() - sheet_t0
            performance["sheet_timings"].append(
                {
                    "phase": "volatile_baseline",
                    "sheet_index": i,
                    "sheet": name,
                    "elapsed_sec": round(sheet_elapsed, 6),
                    "diff_count": diff_count,
                    "hidden_ignored_diff_count": hidden_ignored_count,
                    "bounds_original": bounds_a,
                    "bounds_baseline": bounds_b,
                }
            )
            logging.info(
                "SHEET_DONE volatile_baseline workbook=%s sheet=%d/%d name=%s diffs=%d hidden_ignored=%d elapsed=%.3fs",
                filename,
                i,
                count_a,
                name,
                diff_count,
                hidden_ignored_count,
                sheet_elapsed,
            )
            volatile_baseline_diff_count += diff_count
            if hidden_ignored_count:
                logging.info(
                    "Ignored %d hidden baseline diffs in original %s sheet=%s",
                    hidden_ignored_count,
                    filename,
                    name,
                )
            del sheet_a, sheet_b

        if volatile_baseline_diff_count:
            logging.info(
                "Volatile baseline diffs in original %s: %d cells",
                filename,
                volatile_baseline_diff_count,
            )

        try:
            wb_b.Close(SaveChanges=False)
            logging.info("Closed workbook B original baseline")
        except Exception:
            pass
        open_t0 = time.perf_counter()
        wb_b = open_workbook(excel_b, sp)
        performance["timings_sec"]["open_shrunk_b"] = round(
            time.perf_counter() - open_t0, 6
        )
        performance["memory"]["comparison_app_after_open_shrunk_b"] = (
            process_memory_snapshot(excel_b_pid)
        )

        recalc_t0 = time.perf_counter()
        excel_b.CalculateFull()
        performance["timings_sec"]["recalc_shrunk_b"] = round(
            time.perf_counter() - recalc_t0, 6
        )
        performance["memory"]["comparison_app_after_recalc_shrunk_b"] = (
            process_memory_snapshot(excel_b_pid)
        )

        count_b = int(wb_b.Worksheets.Count)
        if count_a != count_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=f"sheet_count_mismatch: original={count_a}, shrunk={count_b}",
                diffs=[],
                diff_count=0,
                volatile_baseline_diff_count=volatile_baseline_diff_count,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        names_b = [str(wb_b.Worksheets(i).Name) for i in range(1, count_b + 1)]
        if names_a != names_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=f"sheet_name_mismatch: original={names_a}, shrunk={names_b}",
                diffs=[],
                diff_count=0,
                volatile_baseline_diff_count=volatile_baseline_diff_count,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        sheet_b_by_name = {
            str(wb_b.Worksheets(i).Name): wb_b.Worksheets(i)
            for i in range(1, count_b + 1)
        }

        total_diff_count = 0
        volatile_repeated_diff_count = 0
        cleanup_diff_count = 0
        hidden_ignored_diff_count = 0
        all_diffs: List[Diff] = []
        visual_result: Optional[Dict[str, Any]] = None

        for i in range(1, count_a + 1):
            sheet_a = wb_a.Worksheets(i)
            name = str(sheet_a.Name)
            sheet_b = sheet_b_by_name[name]

            sheet_t0 = time.perf_counter()
            bounds_a = get_used_bounds(sheet_a)
            bounds_b = get_used_bounds(sheet_b)
            logging.info(
                "SHEET_START values workbook=%s sheet=%d/%d name=%s bounds_original=%s bounds_shrunk=%s",
                filename,
                i,
                count_a,
                name,
                bounds_a,
                bounds_b,
            )
            remaining = max(0, max_diffs_per_file - len(all_diffs))
            (
                diff_count,
                sheet_volatile_count,
                sheet_cleanup_count,
                sheet_hidden_ignored_count,
                diffs,
            ) = (
                compare_sheet_values(
                    sheet_a,
                    sheet_b,
                    name,
                    row_chunk=row_chunk,
                    col_chunk=col_chunk,
                    max_diffs=remaining,
                    volatile_cells=volatile_cells,
                )
            )
            sheet_elapsed = time.perf_counter() - sheet_t0
            performance["sheet_timings"].append(
                {
                    "phase": "values",
                    "sheet_index": i,
                    "sheet": name,
                    "elapsed_sec": round(sheet_elapsed, 6),
                    "diff_count": diff_count,
                    "volatile_repeated_diff_count": sheet_volatile_count,
                    "cleanup_diff_count": sheet_cleanup_count,
                    "hidden_ignored_diff_count": sheet_hidden_ignored_count,
                    "bounds_original": bounds_a,
                    "bounds_shrunk": bounds_b,
                }
            )
            logging.info(
                "SHEET_DONE values workbook=%s sheet=%d/%d name=%s diffs=%d cleanup_diffs=%d volatile_repeated=%d hidden_ignored=%d elapsed=%.3fs",
                filename,
                i,
                count_a,
                name,
                diff_count,
                sheet_cleanup_count,
                sheet_volatile_count,
                sheet_hidden_ignored_count,
                sheet_elapsed,
            )
            total_diff_count += diff_count
            volatile_repeated_diff_count += sheet_volatile_count
            cleanup_diff_count += sheet_cleanup_count
            hidden_ignored_diff_count += sheet_hidden_ignored_count
            if sheet_hidden_ignored_count:
                logging.info(
                    "Ignored %d hidden diffs in %s sheet=%s",
                    sheet_hidden_ignored_count,
                    filename,
                    name,
                )
            if diffs:
                all_diffs.extend(diffs)

            del sheet_a, sheet_b, diffs

        if visual_compare:
            visual_result = compare_workbook_visuals(
                wb_a,
                wb_b,
                filename,
                rows=visual_rows,
                cols=visual_cols,
                max_sheets=visual_max_sheets,
                pixel_tolerance=visual_pixel_tolerance,
                screenshot_dir=visual_screenshot_dir,
                retry_count=visual_retry_count,
                retry_sleep_sec=visual_retry_sleep_sec,
            )

        visual_same = True
        visual_reason: Optional[str] = None
        if visual_result is not None and not bool(visual_result.get("same")):
            visual_same = False
            visual_status = str(visual_result.get("status") or "different")
            visual_reason = f"visual_{visual_status}"

        if total_diff_count == 0:
            if not visual_same:
                return FileResult(
                    filename=filename,
                    status="different",
                    reason=visual_reason or "visual_mismatch",
                    diffs=[],
                    diff_count=0,
                    volatile_baseline_diff_count=volatile_baseline_diff_count,
                    cleanup_diff_count=0,
                    volatile_repeated_diff_count=0,
                    hidden_ignored_diff_count=hidden_ignored_diff_count,
                    visual=visual_result,
                    performance=performance,
                    elapsed_sec=time.time() - t0,
                )

            return FileResult(
                filename=filename,
                status="same",
                reason=None,
                diffs=[],
                diff_count=0,
                volatile_baseline_diff_count=volatile_baseline_diff_count,
                cleanup_diff_count=0,
                volatile_repeated_diff_count=0,
                hidden_ignored_diff_count=hidden_ignored_diff_count,
                visual=visual_result,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        reason = "value_mismatch"
        if cleanup_diff_count == 0 and volatile_repeated_diff_count > 0:
            reason = "volatile_value_mismatch_only"
        if not visual_same:
            reason = f"{reason};{visual_reason or 'visual_mismatch'}"

        return FileResult(
            filename=filename,
            status="different",
            reason=reason,
            diffs=all_diffs,
            diff_count=total_diff_count,
            volatile_baseline_diff_count=volatile_baseline_diff_count,
            cleanup_diff_count=cleanup_diff_count,
            volatile_repeated_diff_count=volatile_repeated_diff_count,
            hidden_ignored_diff_count=hidden_ignored_diff_count,
            visual=visual_result,
            performance=performance,
            elapsed_sec=time.time() - t0,
        )

    except Exception as e:
        return FileResult(
            filename=filename,
            status="error",
            reason=f"{type(e).__name__}: {e}",
            diffs=[],
            diff_count=0,
            performance=performance,
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


def is_visual_recheck_candidate(result: FileResult) -> bool:
    reason = result.reason or ""
    visual_same = bool(result.visual.get("same")) if result.visual else False
    return (
        result.status == "different"
        and int(result.diff_count or 0) == 0
        and "visual_" in reason
        and not visual_same
    )


def compare_workbook_visual_only_in_excel(
    excel_a: Any,
    excel_b: Any,
    original_path: Path,
    shrunk_path: Path,
    temp_dir: Path,
    previous_result: FileResult,
    visual_rows: int = 60,
    visual_cols: int = 20,
    visual_max_sheets: int = 0,
    visual_pixel_tolerance: int = 0,
    visual_screenshot_dir: Optional[Path] = None,
    visual_retry_count: int = 5,
    visual_retry_sleep_sec: float = 0.25,
) -> FileResult:
    t0 = time.time()
    filename = original_path.name
    wb_a = None
    wb_b = None
    performance: Dict[str, Any] = {
        "enabled": True,
        "mode": "visual_only_recheck",
        "timings_sec": {},
        "memory": {},
    }
    try:
        op = repair_workbook_xml_copy_for_excel(original_path, temp_dir)
        sp = repair_workbook_xml_copy_for_excel(shrunk_path, temp_dir)
        op = ensure_short_path_for_excel(op, temp_dir)
        sp = ensure_short_path_for_excel(sp, temp_dir)

        excel_a_pid = excel_process_id(excel_a)
        excel_b_pid = excel_process_id(excel_b)
        performance["excel_pids"] = {
            "original_app": excel_a_pid,
            "comparison_app": excel_b_pid,
        }

        open_t0 = time.perf_counter()
        wb_a = open_workbook(excel_a, op)
        performance["timings_sec"]["open_original_a"] = round(
            time.perf_counter() - open_t0, 6
        )
        performance["memory"]["original_app_after_open_original_a"] = (
            process_memory_snapshot(excel_a_pid)
        )

        open_t0 = time.perf_counter()
        wb_b = open_workbook(excel_b, sp)
        performance["timings_sec"]["open_shrunk_b"] = round(
            time.perf_counter() - open_t0, 6
        )
        performance["memory"]["comparison_app_after_open_shrunk_b"] = (
            process_memory_snapshot(excel_b_pid)
        )

        recalc_t0 = time.perf_counter()
        excel_a.CalculateFull()
        performance["timings_sec"]["recalc_original_a"] = round(
            time.perf_counter() - recalc_t0, 6
        )

        recalc_t0 = time.perf_counter()
        excel_b.CalculateFull()
        performance["timings_sec"]["recalc_shrunk_b"] = round(
            time.perf_counter() - recalc_t0, 6
        )

        count_a = int(wb_a.Worksheets.Count)
        count_b = int(wb_b.Worksheets.Count)
        if count_a != count_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=f"sheet_count_mismatch: original={count_a}, shrunk={count_b}",
                diffs=previous_result.diffs or [],
                diff_count=previous_result.diff_count,
                volatile_baseline_diff_count=previous_result.volatile_baseline_diff_count,
                cleanup_diff_count=previous_result.cleanup_diff_count,
                volatile_repeated_diff_count=previous_result.volatile_repeated_diff_count,
                hidden_ignored_diff_count=previous_result.hidden_ignored_diff_count,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        names_a = [str(wb_a.Worksheets(i).Name) for i in range(1, count_a + 1)]
        names_b = [str(wb_b.Worksheets(i).Name) for i in range(1, count_b + 1)]
        if names_a != names_b:
            return FileResult(
                filename=filename,
                status="different",
                reason=f"sheet_name_mismatch: original={names_a}, shrunk={names_b}",
                diffs=previous_result.diffs or [],
                diff_count=previous_result.diff_count,
                volatile_baseline_diff_count=previous_result.volatile_baseline_diff_count,
                cleanup_diff_count=previous_result.cleanup_diff_count,
                volatile_repeated_diff_count=previous_result.volatile_repeated_diff_count,
                hidden_ignored_diff_count=previous_result.hidden_ignored_diff_count,
                performance=performance,
                elapsed_sec=time.time() - t0,
            )

        visual_result = compare_workbook_visuals(
            wb_a,
            wb_b,
            filename,
            rows=visual_rows,
            cols=visual_cols,
            max_sheets=visual_max_sheets,
            pixel_tolerance=visual_pixel_tolerance,
            screenshot_dir=visual_screenshot_dir,
            retry_count=visual_retry_count,
            retry_sleep_sec=visual_retry_sleep_sec,
        )
        visual_same = bool(visual_result.get("same"))
        status = "same" if visual_same else "different"
        reason = None if visual_same else f"visual_{visual_result.get('status') or 'different'}"
        return FileResult(
            filename=filename,
            status=status,
            reason=reason,
            diffs=previous_result.diffs or [],
            diff_count=previous_result.diff_count,
            volatile_baseline_diff_count=previous_result.volatile_baseline_diff_count,
            cleanup_diff_count=previous_result.cleanup_diff_count,
            volatile_repeated_diff_count=previous_result.volatile_repeated_diff_count,
            hidden_ignored_diff_count=previous_result.hidden_ignored_diff_count,
            visual=visual_result,
            performance=performance,
            elapsed_sec=time.time() - t0,
        )
    except Exception as e:
        return FileResult(
            filename=filename,
            status="error",
            reason=f"visual_only_recheck_failed: {type(e).__name__}: {e}",
            diffs=previous_result.diffs or [],
            diff_count=previous_result.diff_count,
            performance=performance,
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


def run_single_workbook_child(args: argparse.Namespace) -> int:
    child_handlers: List[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if args.log_file is not None:
        args.log_file.parent.mkdir(parents=True, exist_ok=True)
        child_handlers.append(
            logging.FileHandler(args.log_file, mode="a", encoding="utf-8")
        )
    logging.basicConfig(
        level=getattr(logging, str(args.log_level).upper(), logging.INFO),
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=child_handlers,
        force=True,
    )

    pythoncom.CoInitialize()
    excel_a = None
    excel_b = None
    temp_dir: Optional[Path] = None
    result: FileResult
    try:
        temp_dir = Path(tempfile.mkdtemp(prefix="xlsx_compare_child_"))
        excel_a = open_excel_app()
        excel_b = open_excel_app()
        if args.single_workbook_previous_result is not None:
            with args.single_workbook_previous_result.open("r", encoding="utf-8") as f:
                previous_result = file_result_from_dict(json.load(f))
            result = compare_workbook_visual_only_in_excel(
                excel_a,
                excel_b,
                args.single_workbook_original.resolve(),
                args.single_workbook_shrunk.resolve(),
                temp_dir=temp_dir,
                previous_result=previous_result,
                visual_rows=int(args.visual_rows),
                visual_cols=int(args.visual_cols),
                visual_max_sheets=int(args.visual_max_sheets),
                visual_pixel_tolerance=int(args.visual_pixel_tolerance),
                visual_screenshot_dir=args.visual_screenshot_dir,
                visual_retry_count=int(args.visual_retry_count),
                visual_retry_sleep_sec=float(args.visual_retry_sleep_sec),
            )
        else:
            result = compare_workbooks_in_excel(
                excel_a,
                excel_b,
                args.single_workbook_original.resolve(),
                args.single_workbook_shrunk.resolve(),
                temp_dir=temp_dir,
                row_chunk=int(args.row_chunk),
                col_chunk=int(args.col_chunk),
                max_diffs_per_file=int(args.max_diffs_per_file),
                visual_compare=bool(args.visual_compare),
                visual_rows=int(args.visual_rows),
                visual_cols=int(args.visual_cols),
                visual_max_sheets=int(args.visual_max_sheets),
                visual_pixel_tolerance=int(args.visual_pixel_tolerance),
                visual_screenshot_dir=args.visual_screenshot_dir,
                visual_retry_count=int(args.visual_retry_count),
                visual_retry_sleep_sec=float(args.visual_retry_sleep_sec),
            )
    except Exception as e:
        result = FileResult(
            filename=args.single_workbook_original.name,
            status="error",
            reason=f"{type(e).__name__}: {e}",
            diffs=[],
            diff_count=0,
        )
    args.single_workbook_result.parent.mkdir(parents=True, exist_ok=True)
    with args.single_workbook_result.open("w", encoding="utf-8") as f:
        json.dump(file_result_to_dict(result), f, ensure_ascii=False, indent=2)

    try:
        sys.stdout.flush()
        sys.stderr.flush()
    except Exception:
        pass

    try:
        return 0
    finally:
        close_excel_app(excel_a)
        close_excel_app(excel_b)
        pythoncom.CoUninitialize()
        if temp_dir is not None:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass


def compare_workbook_with_timeout(
    original_path: Path,
    shrunk_path: Path,
    *,
    row_chunk: int,
    col_chunk: int,
    max_diffs_per_file: int,
    visual_compare: bool,
    visual_rows: int,
    visual_cols: int,
    visual_max_sheets: int,
    visual_pixel_tolerance: int,
    visual_screenshot_dir: Optional[Path],
    visual_retry_count: int,
    visual_retry_sleep_sec: float,
    timeout_sec: float,
    timeout_max_sec: float,
    timeout_attempts: int,
    kill_hung_excel: bool,
    excel_baseline_pids: set[int],
    previous_result_for_visual_only: Optional[FileResult] = None,
    log_file: Optional[Path] = None,
) -> FileResult:
    filename = original_path.name
    timeout = float(timeout_sec)
    attempts = max(1, int(timeout_attempts))

    for attempt in range(1, attempts + 1):
        child_result_dir = (
            Path(__file__).resolve().parent
            / "_child_results"
            / f"{hashlib.sha256((str(original_path) + str(time.time()) + str(attempt)).encode('utf-8')).hexdigest()}"
        )
        child_result_dir.mkdir(parents=True, exist_ok=True)
        result_path = child_result_dir / "result.json"
        previous_result_path = child_result_dir / "previous_result.json"
        cmd = [
            sys.executable,
            str(Path(__file__).resolve()),
            "--single-workbook-original",
            str(original_path),
            "--single-workbook-shrunk",
            str(shrunk_path),
            "--single-workbook-result",
            str(result_path),
            "--row-chunk",
            str(row_chunk),
            "--col-chunk",
            str(col_chunk),
            "--max-diffs-per-file",
            str(max_diffs_per_file),
            "--visual-rows",
            str(visual_rows),
            "--visual-cols",
            str(visual_cols),
            "--visual-max-sheets",
            str(visual_max_sheets),
            "--visual-pixel-tolerance",
            str(visual_pixel_tolerance),
            "--visual-retry-count",
            str(visual_retry_count),
            "--visual-retry-sleep-sec",
            str(visual_retry_sleep_sec),
            "--workbook-timeout-sec",
            "0",
        ]
        if previous_result_for_visual_only is not None:
            with previous_result_path.open("w", encoding="utf-8") as f:
                json.dump(
                    file_result_to_dict(previous_result_for_visual_only),
                    f,
                    ensure_ascii=False,
                    indent=2,
                )
            cmd.extend(
                [
                    "--single-workbook-previous-result",
                    str(previous_result_path),
                ]
            )
        if visual_compare:
            cmd.append("--visual-compare")
        else:
            cmd.append("--no-visual-compare")
        if visual_screenshot_dir is not None:
            cmd.extend(["--visual-screenshot-dir", str(visual_screenshot_dir)])
        if log_file is not None:
            cmd.extend(["--log-file", str(log_file)])

        logging.info(
            "Workbook timeout attempt %d/%d for %s: %.1fs result=%s",
            attempt,
            attempts,
            filename,
            timeout,
            result_path,
        )
        proc = subprocess.Popen(
            cmd,
            cwd=str(Path(__file__).resolve().parents[2]),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.STDOUT,
            text=True,
        )

        deadline = time.time() + timeout
        retry_after_child_result = False
        while time.time() < deadline:
            if result_path.exists():
                try:
                    with result_path.open("r", encoding="utf-8") as f:
                        result = file_result_from_dict(json.load(f))
                    if proc.poll() is None:
                        try:
                            proc.wait(timeout=75)
                        except subprocess.TimeoutExpired:
                            logging.warning(
                                "Workbook child cleanup exceeded 75s for %s; "
                                "terminating child process",
                                filename,
                            )
                            proc.terminate()
                            try:
                                proc.wait(timeout=5)
                            except subprocess.TimeoutExpired:
                                proc.kill()
                                proc.wait(timeout=5)
                            if kill_hung_excel:
                                kill_hung_excel_processes(
                                    baseline_pids=excel_baseline_pids,
                                    reason=(
                                        "child_cleanup_timeout "
                                        f"filename={filename} attempt={attempt}"
                                    ),
                                )
                        except Exception:
                            proc.kill()
                            proc.wait(timeout=5)
                            if kill_hung_excel:
                                kill_hung_excel_processes(
                                    baseline_pids=excel_baseline_pids,
                                    reason=(
                                        "child_cleanup_exception "
                                        f"filename={filename} attempt={attempt}"
                                    ),
                                )
                    shutil.rmtree(child_result_dir, ignore_errors=True)
                    if (
                        result.status == "error"
                        and is_retryable_excel_error(result.reason)
                        and attempt < attempts
                    ):
                        timeout = min(float(timeout_max_sec), timeout * 2.0)
                        logging.warning(
                            "Retryable Excel error for %s on attempt %d/%d: %s; "
                            "retrying with timeout %.1fs",
                            filename,
                            attempt,
                            attempts,
                            result.reason,
                            timeout,
                        )
                        if kill_hung_excel:
                            kill_hung_excel_processes(
                                baseline_pids=excel_baseline_pids,
                                reason=(
                                    "retryable_child_result "
                                    f"filename={filename} attempt={attempt} "
                                    f"reason={result.reason}"
                                ),
                            )
                        retry_after_child_result = True
                        break
                    return result
                except Exception as e:
                    logging.warning(
                        "Could not read child result for %s yet: %s", filename, e
                    )
            if proc.poll() is not None:
                break
            time.sleep(0.25)

        if retry_after_child_result:
            continue

        if proc.poll() is None and not result_path.exists():
            logging.warning(
                "Workbook timeout after %.1fs for %s on attempt %d/%d",
                timeout,
                filename,
                attempt,
                attempts,
            )
            proc.kill()
            try:
                child_output, _ = proc.communicate(timeout=10)
            except Exception:
                child_output = ""

            if kill_hung_excel:
                kill_hung_excel_processes(
                    baseline_pids=excel_baseline_pids,
                    reason=f"workbook_timeout filename={filename} attempt={attempt}",
                )

            if attempt < attempts:
                timeout = min(float(timeout_max_sec), timeout * 2.0)
                continue

            return FileResult(
                filename=filename,
                status="error",
                reason=(
                    "workbook_timeout: "
                    f"attempts={attempts}, final_timeout_sec={timeout:.1f}"
                ),
                diffs=[],
                diff_count=0,
            )

        child_output = ""
        try:
            child_output, _ = proc.communicate(timeout=5)
        except Exception:
            pass

        if child_output:
            for line in child_output.splitlines():
                logging.debug("child %s: %s", filename, line)

        if proc.returncode != 0:
            if attempt < attempts:
                timeout = min(float(timeout_max_sec), timeout * 2.0)
                logging.warning(
                    "Workbook child process exited with code %s for %s; retrying",
                    proc.returncode,
                    filename,
                )
                if kill_hung_excel:
                    kill_hung_excel_processes(
                        baseline_pids=excel_baseline_pids,
                        reason=(
                            "child_exit_code "
                            f"filename={filename} attempt={attempt} "
                            f"code={proc.returncode}"
                        ),
                    )
                continue
            return FileResult(
                filename=filename,
                status="error",
                reason=f"workbook_child_exit_code: {proc.returncode}",
                diffs=[],
                diff_count=0,
            )

        try:
            with result_path.open("r", encoding="utf-8") as f:
                result = file_result_from_dict(json.load(f))
            if (
                result.status == "error"
                and is_retryable_excel_error(result.reason)
                and attempt < attempts
            ):
                timeout = min(float(timeout_max_sec), timeout * 2.0)
                logging.warning(
                    "Retryable Excel error for %s on attempt %d/%d: %s; "
                    "retrying with timeout %.1fs",
                    filename,
                    attempt,
                    attempts,
                    result.reason,
                    timeout,
                )
                if kill_hung_excel:
                    kill_hung_excel_processes(
                        baseline_pids=excel_baseline_pids,
                        reason=(
                            "retryable_result_after_exit "
                            f"filename={filename} attempt={attempt} "
                            f"reason={result.reason}"
                        ),
                    )
                continue
            return result
        except Exception as e:
            if attempt < attempts:
                timeout = min(float(timeout_max_sec), timeout * 2.0)
                logging.warning(
                    "Workbook child returned no result for %s: %s; retrying",
                    filename,
                    e,
                )
                continue
            return FileResult(
                filename=filename,
                status="error",
                reason=f"workbook_child_no_result: {type(e).__name__}: {e}",
                diffs=[],
                diff_count=0,
            )

    return FileResult(
        filename=filename,
        status="error",
        reason="workbook_timeout_unreachable",
        diffs=[],
        diff_count=0,
    )


# ----------------------------
# CLI / main
# ----------------------------
def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Verify shrunk .xlsx values; skip if worksheet XML sizes match."
    )
    ap.add_argument(
        "originals",
        nargs="?",
        type=Path,
        help="Folder containing original .xlsx files (non-recursive).",
    )
    ap.add_argument(
        "shrunk",
        nargs="?",
        type=Path,
        help=(
            "Folder containing shrunk .xlsx files (non-recursive), or the output "
            "folder when --shrink-before-compare is used."
        ),
    )
    ap.add_argument("--single-workbook-original", type=Path, default=None)
    ap.add_argument("--single-workbook-shrunk", type=Path, default=None)
    ap.add_argument("--single-workbook-result", type=Path, default=None)
    ap.add_argument("--single-workbook-previous-result", type=Path, default=None)
    ap.add_argument(
        "--out", type=Path, default=Path("verify_report.json"), help="Output JSON path."
    )
    ap.add_argument(
        "--log-file",
        type=Path,
        default=None,
        help=(
            "Write logs to this file. If omitted, a timestamped log file is "
            "created next to --out."
        ),
    )
    ap.add_argument(
        "--repair-originals-dir",
        type=Path,
        default=None,
        help=(
            "Create workbook.xml-only repaired copies of originals in this directory "
            "using scripts/excel_shrink/excel_shrink.py --only-clean-workbook, then "
            "use that directory as the verify baseline."
        ),
    )
    ap.add_argument(
        "--shrink-before-compare",
        action="store_true",
        help=(
            "Shrink each workbook with Rust immediately before comparing it. In this "
            "mode the shrunk positional argument is treated as the per-file output "
            "directory, and shrink runtime/RAM metrics are written into the report."
        ),
    )
    ap.add_argument(
        "--rust-shrink-analysis-dir",
        type=Path,
        default=None,
        help=(
            "Directory for per-workbook Rust analysis CSVs when "
            "--shrink-before-compare is enabled. Defaults to <shrunk>/_analysis."
        ),
    )
    ap.add_argument(
        "--rust-shrink-log-dir",
        type=Path,
        default=None,
        help=(
            "Directory for per-workbook Rust shrink logs when "
            "--shrink-before-compare is enabled. Defaults to <shrunk>/_logs."
        ),
    )
    ap.add_argument(
        "--rust-max-active-workbooks",
        type=int,
        default=1,
        help="Rust shrink --max-active-workbooks in --shrink-before-compare mode.",
    )
    ap.add_argument(
        "--rust-writer-workers",
        type=int,
        default=1,
        help="Rust shrink --writer-workers in --shrink-before-compare mode.",
    )
    ap.add_argument(
        "--rust-remove-unreferenced-hidden-content-cells",
        action="store_true",
        default=True,
        help=(
            "Pass --remove-unreferenced-hidden-content-cells to Rust shrink in "
            "--shrink-before-compare mode."
        ),
    )
    ap.add_argument(
        "--no-rust-remove-unreferenced-hidden-content-cells",
        dest="rust_remove_unreferenced_hidden_content_cells",
        action="store_false",
        help="Do not pass --remove-unreferenced-hidden-content-cells to Rust shrink.",
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
        "--workbook-timeout-sec",
        type=float,
        default=300.0,
        help=(
            "Per-workbook timeout in seconds. Uses a child process so hung Excel "
            "COM calls can be terminated. Set to 0 to disable."
        ),
    )
    ap.add_argument(
        "--workbook-timeout-max-sec",
        type=float,
        default=1800.0,
        help="Maximum timeout after exponential backoff.",
    )
    ap.add_argument(
        "--workbook-timeout-attempts",
        type=int,
        default=3,
        help="How many timeout attempts to make per workbook.",
    )
    ap.add_argument(
        "--kill-hung-excel-processes",
        dest="kill_hung_excel_processes",
        action="store_true",
        default=True,
        help=(
            "When a workbook child times out or Excel returns retryable COM/RPC "
            "errors, kill hidden EXCEL.EXE processes spawned after this verify run "
            "started before retrying."
        ),
    )
    ap.add_argument(
        "--no-kill-hung-excel-processes",
        dest="kill_hung_excel_processes",
        action="store_false",
        help="Disable automatic cleanup of hidden post-start EXCEL.EXE processes.",
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
    ap.add_argument(
        "--visual-compare",
        dest="visual_compare",
        action="store_true",
        default=True,
        help=(
            "Also compare Excel-rendered screenshots for each sheet. "
            "Enabled by default. Requires Pillow and disables the worksheet-size skip."
        ),
    )
    ap.add_argument(
        "--no-visual-compare",
        dest="visual_compare",
        action="store_false",
        help="Disable Excel-rendered screenshot comparison.",
    )
    ap.add_argument(
        "--visual-rows",
        type=int,
        default=60,
        help="Number of top-left rows to render for visual comparison.",
    )
    ap.add_argument(
        "--visual-cols",
        type=int,
        default=20,
        help="Number of top-left columns to render for visual comparison.",
    )
    ap.add_argument(
        "--visual-max-sheets",
        type=int,
        default=0,
        help="Max sheets to visually compare per workbook; 0 means all sheets.",
    )
    ap.add_argument(
        "--visual-pixel-tolerance",
        type=int,
        default=0,
        help="Per-channel pixel tolerance for visual comparison.",
    )
    ap.add_argument(
        "--visual-screenshot-dir",
        type=Path,
        default=None,
        help="Optional directory where original/shrunk visual screenshots are saved.",
    )
    ap.add_argument(
        "--visual-retry-count",
        type=int,
        default=5,
        help="Clipboard capture retry count for visual comparison.",
    )
    ap.add_argument(
        "--visual-retry-sleep-sec",
        type=float,
        default=0.25,
        help="Sleep between clipboard capture retries for visual comparison.",
    )

    return ap.parse_args()


def main() -> int:
    args = parse_args()

    if (
        args.single_workbook_original is not None
        or args.single_workbook_shrunk is not None
        or args.single_workbook_result is not None
    ):
        missing = [
            name
            for name in (
                "single_workbook_original",
                "single_workbook_shrunk",
                "single_workbook_result",
            )
            if getattr(args, name) is None
        ]
        if missing:
            print(f"Missing single-workbook args: {missing}", file=sys.stderr)
            return 2
        return run_single_workbook_child(args)

    log_path = args.log_file
    if log_path is None:
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = args.out.with_name(f"{args.out.stem}_{stamp}.log")
    log_path.parent.mkdir(parents=True, exist_ok=True)

    handlers: List[logging.Handler] = [
        logging.StreamHandler(),
        logging.FileHandler(log_path, encoding="utf-8"),
    ]
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=handlers,
        force=True,
    )
    logging.info("Writing verify log to %s", log_path)

    if args.originals is None or not args.originals.is_dir():
        logging.error("Originals folder invalid: %s", args.originals)
        return 2
    if args.shrunk is None:
        logging.error("Shrunk folder invalid: %s", args.shrunk)
        return 2
    if args.shrink_before_compare:
        args.shrunk.mkdir(parents=True, exist_ok=True)
    elif not args.shrunk.is_dir():
        logging.error("Shrunk folder invalid: %s", args.shrunk)
        return 2

    originals_for_compare = args.originals
    if args.repair_originals_dir is not None:
        try:
            repair_originals_with_python_excel_shrink(
                args.originals, args.repair_originals_dir
            )
        except subprocess.CalledProcessError as e:
            logging.error(
                "Repairing originals failed with exit code %s", e.returncode
            )
            return e.returncode or 1
        except Exception as e:
            logging.error("Repairing originals failed: %s", e)
            return 1

        if not args.repair_originals_dir.is_dir():
            logging.error("Repair originals folder invalid: %s", args.repair_originals_dir)
            return 2
        originals_for_compare = args.repair_originals_dir

    if args.visual_compare:
        try:
            import_pillow_for_visual_compare()
        except RuntimeError as e:
            logging.error("%s", e)
            return 2

    progress_path = progress_journal_path(args.out)
    if args.recompare_all:
        try:
            progress_path.unlink(missing_ok=True)
        except Exception as e:
            logging.warning("Could not remove progress journal %s: %s", progress_path, e)
        previous_results_by_name: Dict[str, FileResult] = {}
    else:
        previous_results_by_name = load_previous_completed_results(args.out)
        if previous_results_by_name:
            logging.info(
                "Resume enabled: loaded %d completed workbook results from %s",
                len(previous_results_by_name),
                progress_path if progress_path.exists() else args.out,
            )

    orig_map = list_xlsx_files(originals_for_compare)
    shr_map = list_xlsx_files(args.shrunk)

    if args.shrink_before_compare:
        common_names = sorted(orig_map.keys())
        missing_in_shrunk: List[str] = []
        missing_in_originals = sorted(set(shr_map.keys()) - set(orig_map.keys()))
        rust_shrink_analysis_dir = (
            args.rust_shrink_analysis_dir
            if args.rust_shrink_analysis_dir is not None
            else args.shrunk / "_analysis"
        )
        rust_shrink_log_dir = (
            args.rust_shrink_log_dir
            if args.rust_shrink_log_dir is not None
            else args.shrunk / "_logs"
        )
        logging.info(
            "Rust shrink-before-compare enabled: output=%s analysis=%s logs=%s",
            args.shrunk,
            rust_shrink_analysis_dir,
            rust_shrink_log_dir,
        )
    else:
        common_names = sorted(set(orig_map.keys()) & set(shr_map.keys()))
        missing_in_shrunk = sorted(set(orig_map.keys()) - set(shr_map.keys()))
        missing_in_originals = sorted(set(shr_map.keys()) - set(orig_map.keys()))
        rust_shrink_analysis_dir = None
        rust_shrink_log_dir = None

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

    results: List[FileResult] = [
        previous_results_by_name[name]
        for name in common_names
        if name in previous_results_by_name
        and previous_results_by_name[name].status != "error"
        and not (
            args.visual_compare
            and is_visual_recheck_candidate(previous_results_by_name[name])
        )
    ]
    use_workbook_timeout = float(args.workbook_timeout_sec) > 0
    excel_baseline_pids = list_excel_process_ids()
    if args.kill_hung_excel_processes:
        logging.info(
            "Excel cleanup enabled; protecting baseline EXCEL.EXE PIDs: %s",
            sorted(excel_baseline_pids),
        )
    else:
        logging.info("Excel cleanup disabled")

    # Initialise COM + Excel once in main thread
    pythoncom.CoInitialize()
    excel_a = None
    excel_b = None
    temp_dir: Optional[Path] = None
    try:
        if not use_workbook_timeout:
            excel_a = open_excel_app()
            excel_b = open_excel_app()

        # Cleanup happens after the JSON report is written, so locked temp files
        # cannot prevent the run results from being persisted.
        with contextlib.nullcontext(tempfile.mkdtemp(prefix="xlsx_compare_")) as td:
            temp_dir = Path(td)

            # loop with index
            for idx, name in enumerate(common_names, start=1):
                op = orig_map[name]
                sp = shr_map.get(name, args.shrunk / name)
                remaining = len(common_names) - idx
                shrink_metrics: Optional[Dict[str, Any]] = None

                previous_result = previous_results_by_name.get(name)
                if previous_result is not None:
                    if args.visual_compare and is_visual_recheck_candidate(
                        previous_result
                    ):
                        logging.info(
                            "VISUAL_RECHECK %-60s previous_status=%s reason=%s",
                            name,
                            previous_result.status,
                            previous_result.reason,
                        )
                    elif previous_result.status == "error":
                        logging.info(
                            "RETRY_PREVIOUS_ERROR %-60s reason=%s",
                            name,
                            previous_result.reason,
                        )
                    else:
                        logging.info(
                            "RESUME  %-60s status=%s reason=%s",
                            name,
                            previous_result.status,
                            previous_result.reason,
                        )
                        continue

                logging.info(
                    "Progress %d/%d (%d remaining): %s",
                    idx,
                    len(common_names),
                    remaining,
                    name,
                )

                if args.shrink_before_compare:
                    assert rust_shrink_analysis_dir is not None
                    assert rust_shrink_log_dir is not None
                    try:
                        logging.info("Shrinking workbook before compare: %s", name)
                        shrunk_output, shrink_metrics = shrink_workbook_with_rust_for_verify(
                            op,
                            args.shrunk,
                            analysis_dir=rust_shrink_analysis_dir,
                            log_dir=rust_shrink_log_dir,
                            max_active_workbooks=int(args.rust_max_active_workbooks),
                            writer_workers=int(args.rust_writer_workers),
                            remove_unreferenced_hidden_content_cells=bool(
                                args.rust_remove_unreferenced_hidden_content_cells
                            ),
                        )
                        if shrunk_output is None:
                            res = FileResult(
                                filename=name,
                                status="error",
                                reason=str(
                                    shrink_metrics.get("reason")
                                    or "rust_shrink_output_missing"
                                ),
                                diffs=[],
                                diff_count=0,
                                performance={"enabled": True, "shrink": shrink_metrics},
                                elapsed_sec=float(
                                    shrink_metrics.get("elapsed_sec") or 0.0
                                ),
                            )
                            results.append(res)
                            append_progress_result(progress_path, res)
                            logging.error(
                                "ERROR   %-60s (%.1fs) %s",
                                res.filename,
                                res.elapsed_sec,
                                res.reason,
                            )
                            continue
                        sp = shrunk_output
                        shr_map[name] = sp
                        logging.info(
                            "Shrunk %-60s shrink=%.3fs peak_ws=%sMB peak_private=%sMB",
                            name,
                            float(shrink_metrics.get("elapsed_sec") or 0.0),
                            shrink_metrics.get("peak_working_set_mb"),
                            shrink_metrics.get("peak_private_memory_mb"),
                        )
                    except Exception as e:
                        res = FileResult(
                            filename=name,
                            status="error",
                            reason=f"rust_shrink_failed: {e}",
                            diffs=[],
                            diff_count=0,
                            performance=(
                                {"enabled": True, "shrink": shrink_metrics}
                                if shrink_metrics
                                else None
                            ),
                            elapsed_sec=0.0,
                        )
                        results.append(res)
                        append_progress_result(progress_path, res)
                        logging.error("ERROR   %-60s (%.1fs) %s", name, 0.0, res.reason)
                        continue

                # 1) Pre-check: skip if xl/worksheets/sheetN.xml sizes equal
                skip, skip_reason = should_skip_by_worksheet_sizes(op, sp)
                if skip and not args.visual_compare and not args.shrink_before_compare:
                    performance = None
                    if shrink_metrics:
                        performance = {"enabled": True, "shrink": shrink_metrics}
                    logging.debug("SKIP    %-60s reason=%s", name, skip_reason)
                    res = FileResult(
                        filename=name,
                        status="skipped",
                        reason=skip_reason,
                        diffs=[],
                        diff_count=0,
                        performance=performance,
                        elapsed_sec=0.0,
                    )
                    results.append(res)
                    append_progress_result(progress_path, res)
                    continue

                if skip and args.visual_compare:
                    skip_reason = f"{skip_reason}; visual_compare_enabled"

                logging.debug("NEEDCHK %-60s reason=%s", name, skip_reason)

                # 2) Excel compare (main thread)
                logging.info("Comparing workbook: %s", name)
                res = None
                if use_workbook_timeout:
                    res = compare_workbook_with_timeout(
                        op,
                        sp,
                        row_chunk=int(args.row_chunk),
                        col_chunk=int(args.col_chunk),
                        max_diffs_per_file=int(args.max_diffs_per_file),
                        visual_compare=bool(args.visual_compare),
                        visual_rows=int(args.visual_rows),
                        visual_cols=int(args.visual_cols),
                        visual_max_sheets=int(args.visual_max_sheets),
                        visual_pixel_tolerance=int(args.visual_pixel_tolerance),
                        visual_screenshot_dir=args.visual_screenshot_dir,
                        visual_retry_count=int(args.visual_retry_count),
                        visual_retry_sleep_sec=float(args.visual_retry_sleep_sec),
                        timeout_sec=float(args.workbook_timeout_sec),
                        timeout_max_sec=float(args.workbook_timeout_max_sec),
                        timeout_attempts=int(args.workbook_timeout_attempts),
                        kill_hung_excel=bool(args.kill_hung_excel_processes),
                        excel_baseline_pids=excel_baseline_pids,
                        log_file=log_path,
                        previous_result_for_visual_only=(
                            previous_result
                            if previous_result is not None
                            and args.visual_compare
                            and is_visual_recheck_candidate(previous_result)
                            else None
                        ),
                    )
                else:
                    for attempt in range(1, 4):
                        assert excel_a is not None and excel_b is not None
                        attempt_temp_dir = temp_dir / f"{idx:06d}_{attempt}"
                        res = compare_workbooks_in_excel(
                            excel_a,
                            excel_b,
                            op,
                            sp,
                            temp_dir=attempt_temp_dir,
                            row_chunk=int(args.row_chunk),
                            col_chunk=int(args.col_chunk),
                            max_diffs_per_file=int(args.max_diffs_per_file),
                            visual_compare=bool(args.visual_compare),
                            visual_rows=int(args.visual_rows),
                            visual_cols=int(args.visual_cols),
                            visual_max_sheets=int(args.visual_max_sheets),
                            visual_pixel_tolerance=int(args.visual_pixel_tolerance),
                            visual_screenshot_dir=args.visual_screenshot_dir,
                            visual_retry_count=int(args.visual_retry_count),
                            visual_retry_sleep_sec=float(args.visual_retry_sleep_sec),
                        )

                        if res.status != "error":
                            break

                        rpc_unavailable = is_excel_rpc_unavailable(res.reason)
                        temp_permission_error = is_temp_permission_error(res.reason)
                        if not rpc_unavailable and not temp_permission_error:
                            break

                        if attempt == 3:
                            break

                        if temp_permission_error:
                            logging.warning(
                                "Temp permission error while comparing %s; "
                                "retrying with a fresh temp directory",
                                name,
                            )
                            time.sleep(0.5)
                            continue

                        logging.warning(
                            "Excel COM RPC unavailable while comparing %s; "
                            "restarting Excel and retrying",
                            name,
                        )
                        close_excel_app(excel_a)
                        close_excel_app(excel_b)
                        if args.kill_hung_excel_processes:
                            kill_hung_excel_processes(
                                baseline_pids=excel_baseline_pids,
                                reason=(
                                    "main_thread_excel_rpc_unavailable "
                                    f"filename={name} attempt={attempt}"
                                ),
                            )
                        time.sleep(2.0)
                        excel_a = open_excel_app()
                        excel_b = open_excel_app()

                assert res is not None
                res = attach_shrink_performance(res, shrink_metrics)
                results.append(res)
                append_progress_result(progress_path, res)

                if res.status == "same":
                    hidden_note = (
                        f" hidden_ignored={res.hidden_ignored_diff_count}"
                        if res.hidden_ignored_diff_count
                        else ""
                    )
                    timing_note = ""
                    if res.performance:
                        timings = res.performance.get("timings_sec") or {}
                        derived = res.performance.get("derived_timings_sec") or {}
                        timing_note = (
                            " open_orig={:.3f}s recalc_orig={:.3f}s "
                            "open_shrunk={:.3f}s recalc_shrunk={:.3f}s"
                        ).format(
                            float(timings.get("open_original_a") or 0.0),
                            float(timings.get("recalc_original_a") or 0.0),
                            float(timings.get("open_shrunk_b") or 0.0),
                            float(timings.get("recalc_shrunk_b") or 0.0),
                        )
                        if derived:
                            timing_note += (
                                " shrink={:.3f}s shrink+open_shrunk={:.3f}s "
                                "shrink+open+recalc_shrunk={:.3f}s"
                            ).format(
                                float(derived.get("rust_shrink_only") or 0.0),
                                float(
                                    derived.get(
                                        "rust_shrink_plus_excel_open_shrunk"
                                    )
                                    or 0.0
                                ),
                                float(
                                    derived.get(
                                        "rust_shrink_plus_excel_open_recalc_shrunk"
                                    )
                                    or 0.0
                                ),
                            )
                    logging.info(
                        "SAME    %-60s (%.1fs)%s%s",
                        res.filename,
                        res.elapsed_sec,
                        hidden_note,
                        timing_note,
                    )
                elif res.status == "different":
                    visual_summary = res.visual.get("summary") if res.visual else None
                    logging.warning(
                        "DIFF    %-60s diffs=%d cleanup_diffs=%d "
                        "volatile_repeated_diffs=%d volatile_baseline_diffs=%d "
                        "hidden_ignored_diffs=%d "
                        "visual_diff_pixels=%s visual_diff_ratio=%s "
                        "(%.1fs) reason=%s",
                        res.filename,
                        res.diff_count,
                        res.cleanup_diff_count,
                        res.volatile_repeated_diff_count,
                        res.volatile_baseline_diff_count,
                        res.hidden_ignored_diff_count,
                        (
                            visual_summary.get("different_pixels")
                            if visual_summary
                            else None
                        ),
                        (
                            visual_summary.get("different_pixel_ratio")
                            if visual_summary
                            else None
                        ),
                        res.elapsed_sec,
                        res.reason,
                    )
                    if res.visual:
                        for sheet_visual in res.visual.get("sheets") or []:
                            if sheet_visual.get("same"):
                                continue
                            logging.warning(
                                "VISUAL_MISMATCH %-60s sheet=%s status=%s "
                                "different_pixels=%s ratio=%s bbox=%s reason=%s",
                                res.filename,
                                sheet_visual.get("sheet"),
                                sheet_visual.get("status"),
                                sheet_visual.get("different_pixels"),
                                sheet_visual.get("different_pixel_ratio"),
                                sheet_visual.get("bbox"),
                                sheet_visual.get("reason"),
                            )
                    for d in res.diffs or []:
                        logging.warning(
                            "MISMATCH %-60s category=%s sheet=%s cell=%s "
                            "original=%r shrunk=%r "
                            "original_formula=%r shrunk_formula=%r",
                            res.filename,
                            d.category,
                            d.sheet,
                            d.cell,
                            d.original,
                            d.shrunk,
                            d.original_formula,
                            d.shrunk_formula,
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
            close_excel_app(excel_a)
        except Exception:
            pass
        try:
            close_excel_app(excel_b)
        except Exception:
            pass
        pythoncom.CoUninitialize()

    # JSON report
    same_files = sorted([r.filename for r in results if r.status == "same"])
    skipped_files = sorted([r.filename for r in results if r.status == "skipped"])
    diff_files = sorted([r.filename for r in results if r.status == "different"])
    err_files = sorted([r.filename for r in results if r.status == "error"])
    visual_results = [r for r in results if r.visual is not None]
    cleanup_changed_files = sorted(
        [r.filename for r in results if r.cleanup_diff_count > 0]
    )
    volatile_only_files = sorted(
        [
            r.filename
            for r in results
            if r.diff_count > 0
            and r.cleanup_diff_count == 0
            and r.volatile_repeated_diff_count > 0
        ]
    )
    volatile_baseline_files = sorted(
        [r.filename for r in results if r.volatile_baseline_diff_count > 0]
    )
    hidden_ignored_files = sorted(
        [r.filename for r in results if r.hidden_ignored_diff_count > 0]
    )
    visual_same_files = sorted(
        [r.filename for r in visual_results if r.visual and r.visual.get("same")]
    )
    visual_diff_files = sorted(
        [
            r.filename
            for r in visual_results
            if r.visual
            and not r.visual.get("same")
            and r.visual.get("status") != "error"
        ]
    )
    visual_err_files = sorted(
        [
            r.filename
            for r in visual_results
            if r.visual and r.visual.get("status") == "error"
        ]
    )

    report: Dict[str, Any] = {
        "meta": {
            "originals_folder": str(args.originals),
            "originals_for_compare_folder": str(originals_for_compare),
            "repair_originals_dir": (
                str(args.repair_originals_dir)
                if args.repair_originals_dir is not None
                else None
            ),
            "shrunk_folder": str(args.shrunk),
            "shrink_before_compare": bool(args.shrink_before_compare),
            "rust_shrink_analysis_dir": (
                str(rust_shrink_analysis_dir)
                if rust_shrink_analysis_dir is not None
                else None
            ),
            "rust_shrink_log_dir": (
                str(rust_shrink_log_dir)
                if rust_shrink_log_dir is not None
                else None
            ),
            "rust_max_active_workbooks": int(args.rust_max_active_workbooks),
            "rust_writer_workers": int(args.rust_writer_workers),
            "rust_remove_unreferenced_hidden_content_cells": bool(
                args.rust_remove_unreferenced_hidden_content_cells
            ),
            "row_chunk": int(args.row_chunk),
            "col_chunk": int(args.col_chunk),
            "max_diffs_per_file": int(args.max_diffs_per_file),
            "log_file": str(log_path),
            "workbook_timeout_sec": float(args.workbook_timeout_sec),
            "workbook_timeout_max_sec": float(args.workbook_timeout_max_sec),
            "workbook_timeout_attempts": int(args.workbook_timeout_attempts),
            "kill_hung_excel_processes": bool(args.kill_hung_excel_processes),
            "excel_baseline_pids": sorted(excel_baseline_pids),
            "visual_compare": bool(args.visual_compare),
            "visual_rows": int(args.visual_rows),
            "visual_cols": int(args.visual_cols),
            "visual_max_sheets": int(args.visual_max_sheets),
            "visual_pixel_tolerance": int(args.visual_pixel_tolerance),
            "visual_screenshot_dir": (
                str(args.visual_screenshot_dir)
                if args.visual_screenshot_dir is not None
                else None
            ),
            "common_files": len(common_names),
            "missing_in_shrunk": missing_in_shrunk,
            "missing_in_originals": missing_in_originals,
        },
        "summary": {
            "same": len(same_files),
            "skipped": len(skipped_files),
            "different": len(diff_files),
            "error": len(err_files),
            "cleanup_changed_files": len(cleanup_changed_files),
            "volatile_only_files": len(volatile_only_files),
            "volatile_baseline_files": len(volatile_baseline_files),
            "hidden_ignored_files": len(hidden_ignored_files),
            "cleanup_diff_cells": sum(r.cleanup_diff_count for r in results),
            "volatile_repeated_diff_cells": sum(
                r.volatile_repeated_diff_count for r in results
            ),
            "volatile_baseline_diff_cells": sum(
                r.volatile_baseline_diff_count for r in results
            ),
            "hidden_ignored_diff_cells": sum(
                r.hidden_ignored_diff_count for r in results
            ),
            "visual_same": len(visual_same_files),
            "visual_different": len(visual_diff_files),
            "visual_error": len(visual_err_files),
        },
        "same_files": same_files,
        "skipped_files": [],
        "different_files": [],
        "errors": [],
        "visual_files": [],
        "performance_files": [],
        "value_difference_categories": {
            "cleanup_changed_files": cleanup_changed_files,
            "volatile_only_files": volatile_only_files,
            "volatile_baseline_files": volatile_baseline_files,
            "hidden_ignored_files": hidden_ignored_files,
        },
    }

    for r in results:
        if r.status == "skipped":
            report["skipped_files"].append({"filename": r.filename, "reason": r.reason})
        elif r.status == "same" and r.hidden_ignored_diff_count:
            report.setdefault("same_with_hidden_ignored_files", []).append(
                {
                    "filename": r.filename,
                    "hidden_ignored_diff_count": r.hidden_ignored_diff_count,
                    "elapsed_sec": r.elapsed_sec,
                }
            )
        elif r.status == "different":
            report["different_files"].append(
                {
                    "filename": r.filename,
                    "reason": r.reason,
                    "diff_count": r.diff_count,
                    "cleanup_diff_count": r.cleanup_diff_count,
                    "volatile_repeated_diff_count": r.volatile_repeated_diff_count,
                    "volatile_baseline_diff_count": r.volatile_baseline_diff_count,
                    "hidden_ignored_diff_count": r.hidden_ignored_diff_count,
                    "elapsed_sec": r.elapsed_sec,
                    "diffs": [
                        {
                            "category": d.category,
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

        if r.visual is not None:
            report["visual_files"].append(
                {
                    "filename": r.filename,
                    **r.visual,
                }
            )

        if r.performance is not None:
            report["performance_files"].append(
                {
                    "filename": r.filename,
                    "status": r.status,
                    **r.performance,
                }
            )

    report["skipped_files"].sort(key=lambda x: x["filename"])
    report["different_files"].sort(key=lambda x: x["filename"])
    report["errors"].sort(key=lambda x: x["filename"])
    report["visual_files"].sort(key=lambda x: x["filename"])
    report["performance_files"].sort(key=lambda x: x["filename"])

    args.out.parent.mkdir(parents=True, exist_ok=True)
    with args.out.open("w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    logging.info("Wrote report: %s", args.out)

    if temp_dir is not None:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception as e:
            logging.warning("Temp cleanup failed after report write: %s", e)

    # Drop references (symbolic but fine)
    del results, report, orig_map, shr_map

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
