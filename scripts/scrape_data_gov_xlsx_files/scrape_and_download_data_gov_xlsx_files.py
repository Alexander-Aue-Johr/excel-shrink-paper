#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Download XLSX files from catalog.data.gov via CKAN API and validate them by
opening in headless Microsoft Excel (read-only, no repair).

Behaviours requested:
- Default output base dir: current working directory (where you run the script)
  and download subfolder is ./downloaded/
- If --out is provided, use that as base dir instead (and still use <out>/downloaded/)
- Writes <base>/downloaded_links.csv
- If downloaded_links.csv exists and --scrape is NOT passed:
    -> Do NOT scrape. Instead, read the CSV and try to download missing/failed items.
- If --scrape is passed:
    -> Scrape CKAN and append new URLs even if CSV exists.
- No max-files limit, no size threshold.
- Reject HTML masquerading as xlsx, reject non-zip, reject Excel-open failures.
- Handle duplicate filenames by renaming with (2), (3), ...
- Record invalid downloads in CSV (and delete unless --keep-invalid).

Windows requirement:
  pip install requests pywin32
  Excel must be installed.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import os
import re
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple, Dict, Set, List
from urllib.parse import urlparse, unquote

import requests


CKAN_BASE = "https://catalog.data.gov/api/3/action"
PACKAGE_SEARCH = f"{CKAN_BASE}/package_search"

UA = "xlsx-bulk-downloader/2.1 (+research; contact: none)"


@dataclass(frozen=True)
class ResourceHit:
    dataset_id: str
    dataset_title: str
    dataset_name: str
    resource_id: str
    resource_name: str
    resource_format: str
    url: str
    mimetype: str | None


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def safe_filename(name: str, max_len: int = 180) -> str:
    name = re.sub(r"[^\w\-. ]+", "_", name, flags=re.UNICODE).strip()
    name = re.sub(r"\s+", " ", name)
    if len(name) > max_len:
        root, ext = os.path.splitext(name)
        name = root[: max_len - len(ext) - 1] + "_" + ext
    return name


def looks_like_xlsx_url(url: str) -> bool:
    u = url.lower()
    if u.endswith(".xlsx"):
        return True
    if ".xlsx?" in u or ("xlsx" in u and "download" in u and "format=xlsx" in u):
        return True
    return False


def derive_filename_from_url(url: str) -> str:
    parsed = urlparse(url)
    base = Path(unquote(parsed.path)).name
    if base and base.lower().endswith(".xlsx"):
        return safe_filename(base)
    return safe_filename("download.xlsx")


def derive_filename(hit: ResourceHit) -> str:
    parsed = urlparse(hit.url)
    base = Path(unquote(parsed.path)).name
    if base and base.lower().endswith(".xlsx"):
        return safe_filename(base)
    base = f"{hit.dataset_name}__{hit.resource_id}.xlsx"
    return safe_filename(base)


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            b = f.read(chunk_size)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def validate_xlsx_is_zip(path: Path) -> Tuple[bool, str]:
    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                return False, f"zip_corrupt (first bad member: {bad})"
        return True, "ok"
    except zipfile.BadZipFile:
        return False, "not_a_zip"
    except Exception as e:
        return False, f"zip_error: {type(e).__name__}: {e}"


def uniquify_filename(download_dir: Path, desired_name: str) -> Tuple[str, str]:
    desired = Path(desired_name)
    stem = desired.stem
    suffix = desired.suffix
    if not (download_dir / desired_name).exists():
        return desired_name, ""
    i = 2
    while True:
        candidate = f"{stem} ({i}){suffix}"
        if not (download_dir / candidate).exists():
            return (
                candidate,
                f'renamed_due_to_duplicate_filename (original="{desired_name}")',
            )
        i += 1


def iter_ckan_xlsx_hits(
    session: requests.Session,
    page_size: int = 100,
    max_pages: int = 10_000,
    query: str = "",
) -> Iterable[ResourceHit]:
    start = 0
    fq = 'res_format:("XLSX" OR "xlsx" OR "EXCEL" OR "Excel")'
    for _ in range(max_pages):
        params = {"q": query, "fq": fq, "rows": page_size, "start": start}
        r = session.get(PACKAGE_SEARCH, params=params, timeout=60)
        r.raise_for_status()
        data = r.json()
        if not data.get("success"):
            raise RuntimeError(f"CKAN API returned success=false: {data}")

        result = data["result"]
        results = result.get("results", [])
        if not results:
            break

        for ds in results:
            ds_id = ds.get("id", "")
            ds_title = ds.get("title", "") or ""
            ds_name = ds.get("name", "") or ""
            for res in ds.get("resources", []) or []:
                url = (res.get("url") or "").strip()
                if not url:
                    continue
                res_format = (res.get("format") or "").strip()
                mimetype = res.get("mimetype") or res.get("mimetype_inner")

                if not looks_like_xlsx_url(url):
                    fmt = res_format.lower()
                    mt = (mimetype or "").lower()
                    if fmt != "xlsx" and "spreadsheetml" not in mt:
                        continue

                yield ResourceHit(
                    dataset_id=ds_id,
                    dataset_title=ds_title,
                    dataset_name=ds_name,
                    resource_id=res.get("id", ""),
                    resource_name=(res.get("name") or "").strip(),
                    resource_format=res_format,
                    url=url,
                    mimetype=mimetype,
                )

        start += page_size
        if start >= int(result.get("count", 0)):
            break


def download_file(
    session: requests.Session,
    url: str,
    dest: Path,
    timeout: int = 180,
    retries: int = 4,
    backoff_s: float = 2.0,
) -> Tuple[bool, str, int]:
    last_status = -1
    for attempt in range(1, retries + 1):
        try:
            with session.get(
                url, stream=True, allow_redirects=True, timeout=timeout
            ) as r:
                last_status = r.status_code
                if r.status_code >= 400:
                    return False, f"HTTP {r.status_code}", r.status_code

                dest_tmp = dest.with_suffix(dest.suffix + ".part")
                with dest_tmp.open("wb") as f:
                    for chunk in r.iter_content(chunk_size=1024 * 256):
                        if chunk:
                            f.write(chunk)
                dest_tmp.replace(dest)
            return True, "ok", last_status
        except Exception as e:
            if attempt == retries:
                return False, f"error: {type(e).__name__}: {e}", last_status
            time.sleep(backoff_s * attempt)
    return False, "unknown", last_status


class ExcelValidator:
    def __init__(self) -> None:
        self._excel = None
        self._pythoncom = None
        self._initialised = False

    def start(self) -> None:
        if self._initialised:
            return
        try:
            import pythoncom
            import win32com.client  # type: ignore
        except Exception as e:
            raise RuntimeError(
                "pywin32 is required for Excel validation. Install with: pip install pywin32"
            ) from e

        self._pythoncom = pythoncom
        pythoncom.CoInitialize()

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.AskToUpdateLinks = False
        except Exception:
            pass

        self._excel = excel
        self._initialised = True

    def stop(self) -> None:
        if not self._initialised:
            return
        try:
            if self._excel is not None:
                try:
                    self._excel.Quit()
                except Exception:
                    pass
        finally:
            self._excel = None
            self._initialised = False
            try:
                if self._pythoncom is not None:
                    self._pythoncom.CoUninitialize()
            except Exception:
                pass

    def can_open(self, path: Path) -> Tuple[bool, str]:
        if not self._initialised:
            self.start()
        assert self._excel is not None
        try:
            wb = self._excel.Workbooks.Open(
                str(path),
                UpdateLinks=0,
                ReadOnly=True,
                AddToMru=False,
                CorruptLoad=0,
                Notify=False,
            )
            wb.Close(SaveChanges=False)
            return True, "excel_ok"
        except Exception as e:
            return False, f"excel_open_failed: {type(e).__name__}: {e}"


CSV_FIELDS = [
    "timestamp_utc",
    "status",
    "note",
    "http_status",
    "dataset_title",
    "dataset_name",
    "dataset_id",
    "resource_name",
    "resource_id",
    "resource_format",
    "mimetype",
    "url",
    "original_filename",
    "final_filename",
    "relative_path",
    "bytes",
    "sha256",
]


def load_csv_rows(csv_path: Path) -> List[Dict[str, str]]:
    if not csv_path.exists():
        return []
    with csv_path.open("r", newline="", encoding="utf-8") as f:
        rd = csv.DictReader(f)
        return [dict(row) for row in rd]


def urls_in_csv(csv_rows: List[Dict[str, str]]) -> Set[str]:
    s: Set[str] = set()
    for r in csv_rows:
        u = (r.get("url") or "").strip()
        if u:
            s.add(u)
    return s


def used_filenames_from_csv(csv_rows: List[Dict[str, str]]) -> Set[str]:
    used: Set[str] = set()
    for r in csv_rows:
        status = (r.get("status") or "").strip()
        fn = (r.get("final_filename") or "").strip()
        if fn and status in {"downloaded", "downloaded_renamed"}:
            used.add(fn)
    return used


def append_csv_row(csv_path: Path, row: Dict[str, object]) -> None:
    write_header = not csv_path.exists()
    with csv_path.open("a", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        if write_header:
            wr.writeheader()
        wr.writerow({k: row.get(k, "") for k in CSV_FIELDS})
        f.flush()


def should_retry_row(
    row: Dict[str, str], download_dir: Path, only_missing: bool
) -> bool:
    """
    Decide whether to attempt downloading a row from existing CSV.
    - If status already downloaded and file exists => skip
    - If only_missing: retry only if file missing
    - Else: retry failures/invalids + missing
    """
    status = (row.get("status") or "").strip()
    http_status = (row.get("http_status") or "").strip()
    final_name = (row.get("final_filename") or "").strip()
    url = (row.get("url") or "").strip()

    if not url:
        return False

    if status == "failed" and http_status == "404":
        return False

    if status in {"downloaded", "downloaded_renamed"} and final_name:
        if (download_dir / final_name).exists():
            return False
        # downloaded but missing file -> redo
        return True

    if only_missing:
        # only missing local file (we need a filename to check; if absent, retry)
        if not final_name:
            return True
        return not (download_dir / final_name).exists()

    # default: retry anything that isn't a confirmed existing downloaded file
    return True


def process_one_url(
    *,
    sess: requests.Session,
    excel: ExcelValidator,
    out_dir: Path,
    download_dir: Path,
    csv_path: Path,
    url: str,
    meta: Dict[str, str],
    used_file_names: Set[str],
    keep_invalid: bool,
    sleep_s: float,
    suggested_filename: Optional[str] = None,
) -> Tuple[bool, str]:
    """
    Downloads + validates one URL, appends a CSV row.
    Returns (kept, final_filename_or_empty).
    """
    ts = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

    original_name = (
        safe_filename(suggested_filename)
        if suggested_filename
        else derive_filename_from_url(url)
    )
    if not original_name.lower().endswith(".xlsx"):
        original_name = original_name + ".xlsx"

    final_name = original_name
    comment = ""
    if final_name in used_file_names or (download_dir / final_name).exists():
        final_name, comment = uniquify_filename(download_dir, final_name)

    dest = download_dir / final_name

    ok, dl_note, http_status = download_file(sess, url, dest)
    if not ok:
        if http_status == 404:
            print(f"SKIP HTTP 404: {url}", flush=True)
        else:
            print(f"SKIP download failed ({dl_note}): {url}", flush=True)

        append_csv_row(
            csv_path,
            {
                "timestamp_utc": ts,
                "status": "failed",
                "note": dl_note,
                "http_status": http_status if http_status != -1 else "",
                **meta,
                "url": url,
                "original_filename": original_name,
                "final_filename": "",
                "relative_path": "",
                "bytes": "",
                "sha256": "",
            },
        )
        time.sleep(sleep_s)
        return False, ""

    zip_ok, zip_note = validate_xlsx_is_zip(dest)
    if not zip_ok:
        note = f"not_a_real_xlsx_container: {zip_note}"
        if comment:
            note = f"{note}; {comment}"
        size = dest.stat().st_size
        digest = sha256_file(dest)
        append_csv_row(
            csv_path,
            {
                "timestamp_utc": ts,
                "status": "invalid_zip",
                "note": note,
                "http_status": http_status if http_status != -1 else "",
                **meta,
                "url": url,
                "original_filename": original_name,
                "final_filename": dest.name,
                "relative_path": str(dest.relative_to(out_dir)),
                "bytes": size,
                "sha256": digest,
            },
        )
        if not keep_invalid:
            try:
                dest.unlink(missing_ok=True)
            except Exception:
                pass
        time.sleep(sleep_s)
        return False, ""

    excel_ok, excel_note = excel.can_open(dest)
    if not excel_ok:
        note = f"not_openable_in_excel (treated_as_corrupt_or_not_real_excel): {excel_note}"
        if comment:
            note = f"{note}; {comment}"
        size = dest.stat().st_size
        digest = sha256_file(dest)
        append_csv_row(
            csv_path,
            {
                "timestamp_utc": ts,
                "status": "invalid_excel",
                "note": note,
                "http_status": http_status if http_status != -1 else "",
                **meta,
                "url": url,
                "original_filename": original_name,
                "final_filename": dest.name,
                "relative_path": str(dest.relative_to(out_dir)),
                "bytes": size,
                "sha256": digest,
            },
        )
        if not keep_invalid:
            try:
                dest.unlink(missing_ok=True)
            except Exception:
                pass
        time.sleep(sleep_s)
        return False, ""

    # Passed all checks -> keep
    size = dest.stat().st_size
    digest = sha256_file(dest)

    status = "downloaded"
    note = "ok"
    if comment:
        status = "downloaded_renamed"
        note = comment

    append_csv_row(
        csv_path,
        {
            "timestamp_utc": ts,
            "status": status,
            "note": note,
            "http_status": http_status if http_status != -1 else "",
            **meta,
            "url": url,
            "original_filename": original_name,
            "final_filename": dest.name,
            "relative_path": str(dest.relative_to(out_dir)),
            "bytes": size,
            "sha256": digest,
        },
    )
    used_file_names.add(dest.name)
    time.sleep(sleep_s)
    return True, dest.name


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--out",
        default=".",
        help='Base output directory. Default is "." (current directory). '
        "Files go into <out>/downloaded and CSV into <out>/downloaded_links.csv.",
    )
    ap.add_argument(
        "--scrape",
        action="store_true",
        help="Force scraping CKAN even if downloaded_links.csv already exists.",
    )
    ap.add_argument("--page-size", type=int, default=100, help="CKAN API page size.")
    ap.add_argument(
        "--query", default="", help='Optional CKAN search query, e.g. "education".'
    )
    ap.add_argument(
        "--sleep", type=float, default=0.2, help="Sleep between downloads (seconds)."
    )
    ap.add_argument(
        "--keep-invalid", action="store_true", help="Keep invalid downloads on disk."
    )
    ap.add_argument(
        "--only-missing",
        action="store_true",
        help="When CSV exists and not scraping: only download entries whose local file is missing.",
    )
    args = ap.parse_args()

    out_dir = Path(args.out).resolve()
    ensure_dir(out_dir)

    download_dir = out_dir / "downloaded"
    ensure_dir(download_dir)

    csv_path = out_dir / "downloaded_links.csv"

    sess = requests.Session()
    sess.headers.update({"User-Agent": UA})

    excel = ExcelValidator()

    csv_rows = load_csv_rows(csv_path)
    csv_exists = csv_path.exists()
    urls_seen = urls_in_csv(csv_rows)
    used_file_names: Set[str] = used_filenames_from_csv(csv_rows)

    # Include what's already on disk too
    for p in download_dir.glob("*.xlsx"):
        used_file_names.add(p.name)

    kept = 0
    attempted = 0

    try:
        # Mode A: CSV exists and NOT forced to scrape -> resume from CSV
        if csv_exists and not args.scrape:
            for row in csv_rows:
                if not should_retry_row(
                    row, download_dir, only_missing=args.only_missing
                ):
                    continue

                url = (row.get("url") or "").strip()
                if not url:
                    continue

                # Meta columns to carry through (best effort)
                meta = {
                    "dataset_title": row.get("dataset_title", ""),
                    "dataset_name": row.get("dataset_name", ""),
                    "dataset_id": row.get("dataset_id", ""),
                    "resource_name": row.get("resource_name", ""),
                    "resource_id": row.get("resource_id", ""),
                    "resource_format": row.get("resource_format", ""),
                    "mimetype": row.get("mimetype", ""),
                }

                suggested = (
                    row.get("original_filename") or row.get("final_filename") or None
                )

                attempted += 1
                ok, _final = process_one_url(
                    sess=sess,
                    excel=excel,
                    out_dir=out_dir,
                    download_dir=download_dir,
                    csv_path=csv_path,
                    url=url,
                    meta=meta,
                    used_file_names=used_file_names,
                    keep_invalid=args.keep_invalid,
                    sleep_s=args.sleep,
                    suggested_filename=suggested,
                )
                if ok:
                    kept += 1

            print(
                f"Resume-only mode (CSV exists, no --scrape). Attempted: {attempted}, kept: {kept}"
            )
            print(f"Download dir: {download_dir}")
            print(f"CSV: {csv_path}")
            return 0

        # Mode B: scrape (either CSV missing, or --scrape forced)
        for hit in iter_ckan_xlsx_hits(
            sess, page_size=args.page_size, query=args.query
        ):
            if hit.url in urls_seen:
                continue
            urls_seen.add(hit.url)

            meta = {
                "dataset_title": hit.dataset_title,
                "dataset_name": hit.dataset_name,
                "dataset_id": hit.dataset_id,
                "resource_name": hit.resource_name,
                "resource_id": hit.resource_id,
                "resource_format": hit.resource_format,
                "mimetype": hit.mimetype or "",
            }

            attempted += 1
            ok, _final = process_one_url(
                sess=sess,
                excel=excel,
                out_dir=out_dir,
                download_dir=download_dir,
                csv_path=csv_path,
                url=hit.url,
                meta=meta,
                used_file_names=used_file_names,
                keep_invalid=args.keep_invalid,
                sleep_s=args.sleep,
                suggested_filename=derive_filename(hit),
            )
            if ok:
                kept += 1

        print(
            f"Scrape mode ({'forced' if args.scrape and csv_exists else 'normal'}). Attempted: {attempted}, kept: {kept}"
        )
        print(f"Download dir: {download_dir}")
        print(f"CSV: {csv_path}")
        return 0

    finally:
        excel.stop()


if __name__ == "__main__":
    raise SystemExit(main())
