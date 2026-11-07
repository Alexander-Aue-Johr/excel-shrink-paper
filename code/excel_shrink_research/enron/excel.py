"""Run Excel in an owned process, with deadlines outside COM."""

import ctypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time
from typing import Any, Dict, Optional
import zipfile


class ExcelInfrastructureError(RuntimeError):
    """Excel could not start; retrying every workbook would hide a host problem."""

    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.details = details or {}


def _write_json(path: Path, value: Dict[str, Any]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    for attempt in range(30):
        try:
            os.replace(str(temporary), str(path))
            return
        except PermissionError:
            if attempt == 29:
                raise
            time.sleep(0.01)


def _read_json(path: Path) -> Dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return {}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _is_xlsx(path: Path) -> bool:
    """Use the package type, not the filename extension."""
    try:
        with zipfile.ZipFile(path) as archive:
            content_types = archive.read("[Content_Types].xml")
            return (
                b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                in content_types
                and "xl/workbook.xml" in archive.namelist()
            )
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


def _check_xlsx(path: Path) -> Optional[str]:
    if not _is_xlsx(path):
        return "Excel did not produce an XLSX workbook."
    try:
        with zipfile.ZipFile(path) as archive:
            bad_member = archive.testzip()
        if bad_member:
            return "The XLSX package has an invalid CRC: " + bad_member
    except (OSError, RuntimeError, zipfile.BadZipFile) as error:
        return "The XLSX package could not be read: " + str(error)
    return None


def _create_seed(path: Path) -> None:
    """A blank workbook gives /x an EXCEL7 window before opening untrusted input."""
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, contents in parts.items():
            archive.writestr(name, contents)


def _excel_executable() -> Path:
    if os.name != "nt":
        raise ExcelInfrastructureError("Excel conversion requires Windows and Microsoft Excel.")
    import winreg

    key_name = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, key_name, 0, winreg.KEY_READ | view) as key:
                    path = Path(winreg.QueryValueEx(key, "")[0])
                if path.is_file():
                    return path
            except OSError:
                pass
    raise ExcelInfrastructureError("Microsoft Excel is not registered on this computer.")


def preflight() -> Dict[str, str]:
    executable = _excel_executable()
    powershell = shutil.which("powershell.exe")
    if powershell is None:
        raise ExcelInfrastructureError("Windows PowerShell is not available.")
    return {"excel_executable": str(executable), "powershell_executable": powershell}


def _owned_handle(record: Dict[str, Any], access: int):
    """A PID alone is insufficient: process IDs can be reused."""
    if os.name != "nt" or not record.get("pid") or not record.get("created_filetime"):
        return None
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.OpenProcess(access | 0x1000, False, int(record["pid"]))
    if not handle:
        return None
    times = [wintypes.FILETIME() for _ in range(4)]
    if not kernel.GetProcessTimes(handle, *(ctypes.byref(value) for value in times)):
        kernel.CloseHandle(handle)
        return None
    created = (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime
    if created != int(record["created_filetime"]):
        kernel.CloseHandle(handle)
        return None
    return handle


def _terminate_owned(record: Dict[str, Any]) -> bool:
    handle = _owned_handle(record, 0x0001)
    if not handle:
        return False
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    try:
        return bool(kernel.TerminateProcess(handle, 124))
    finally:
        kernel.CloseHandle(handle)


def _dialog_evidence(record: Dict[str, Any]) -> Dict[str, Any]:
    """Inspect only the owned Excel process; do not capture the user's desktop."""
    handle = _owned_handle(record, 0)
    if not handle:
        return {"windows": []}
    from ctypes import wintypes

    user = ctypes.WinDLL("user32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user.EnumWindows.argtypes = [callback_type, wintypes.LPARAM]
    user.EnumChildWindows.argtypes = [wintypes.HWND, callback_type, wintypes.LPARAM]
    user.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user.IsWindowVisible.argtypes = [wintypes.HWND]
    windows = []

    def describe(window):
        caption = ctypes.create_unicode_buffer(1024)
        class_name = ctypes.create_unicode_buffer(128)
        user.GetWindowTextW(window, caption, len(caption))
        user.GetClassNameW(window, class_name, len(class_name))
        return {"class": class_name.value, "title": caption.value,
                "visible": bool(user.IsWindowVisible(window))}

    @callback_type
    def collect(window, unused):
        process_id = wintypes.DWORD()
        user.GetWindowThreadProcessId(window, ctypes.byref(process_id))
        if process_id.value != int(record["pid"]):
            return True
        details = describe(window)
        if details["class"] == "#32770":
            children = []

            @callback_type
            def child(window_child, unused_child):
                if len(children) < 100:
                    item = describe(window_child)
                    if item["class"].lower() != "edit":
                        children.append(item)
                return True

            user.EnumChildWindows(window, child, 0)
            details["controls"] = children
        windows.append(details)
        return True

    try:
        user.EnumWindows(collect, 0)
        return {"pid": record["pid"], "windows": windows}
    finally:
        kernel.CloseHandle(handle)


def _run_worker(command, directory: Path, stage_timeout: float,
                total_timeout: float) -> Dict[str, Any]:
    """Keep the deadlines outside PowerShell and all COM calls."""
    stage_file = directory / "stage.json"
    result_file = directory / "result.json"
    process_file = directory / "process.json"
    # A retry must not inherit an old result or process identity.
    for previous in (stage_file, result_file, process_file):
        previous.unlink(missing_ok=True)
    start = time.monotonic()
    stage_start = start
    stage = "startup"
    creation_flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    with (directory / "worker.log").open("wb") as log:
        try:
            worker = subprocess.Popen(command, stdout=log, stderr=subprocess.STDOUT,
                                      stdin=subprocess.DEVNULL, creationflags=creation_flags)
        except OSError as error:
            raise ExcelInfrastructureError("The Excel worker could not start: " + str(error)) from error
        timed_out = None
        try:
            while worker.poll() is None:
                state = _read_json(stage_file)
                current_stage = state.get("stage", stage)
                if current_stage != stage:
                    stage, stage_start = current_stage, time.monotonic()
                now = time.monotonic()
                if now - start >= total_timeout:
                    timed_out = "total_timeout"
                elif now - stage_start >= stage_timeout:
                    timed_out = "stage_timeout"
                if timed_out:
                    _write_json(directory / "windows.json", _dialog_evidence(_read_json(process_file)))
                    break
                time.sleep(0.1)
        finally:
            if worker.poll() is None:
                worker.terminate()
                try:
                    worker.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    worker.kill()
                    worker.wait(timeout=5)
            # Killing PowerShell closes the job handle; this birth-checked fallback
            # also covers a worker that exited without its normal cleanup.
            _terminate_owned(_read_json(process_file))
    elapsed = time.monotonic() - start
    state = _read_json(stage_file)
    stage = state.get("stage", stage)
    if timed_out:
        return {"status": "review", "reason": timed_out, "stage": stage,
                "elapsed_seconds": elapsed, "process": _read_json(process_file),
                "error": "Excel exceeded the configured deadline; the cause is not yet known."}
    result = _read_json(result_file)
    if not result:
        result = {"status": "review", "reason": "worker_failed", "stage": stage,
                  "error": "The worker exited without a result. See worker.log."}
    result["elapsed_seconds"] = elapsed
    result["worker_exit_code"] = worker.returncode
    result["process"] = _read_json(process_file)
    return result


def _invoke_phase(mode: str, source: Path, output: Path, directory: Path,
                  executable: Path, stage_timeout: float, total_timeout: float) -> Dict[str, Any]:
    directory.mkdir(parents=True, exist_ok=True)
    config = {
        "mode": mode, "source": str(source.resolve()), "output": str(output.resolve()),
        "excel": str(executable), "directory": str(directory.resolve()),
        "seed": str((directory / "seed.xlsx").resolve()),
    }
    _create_seed(directory / "seed.xlsx")
    _write_json(directory / "config.json", config)
    powershell = shutil.which("powershell.exe")
    if powershell is None:
        raise ExcelInfrastructureError("Windows PowerShell is not available.")
    command = [powershell, "-NoLogo", "-NoProfile", "-NonInteractive", "-STA",
               "-ExecutionPolicy", "Bypass", "-File", str(Path(__file__).with_name("excel_worker.ps1")),
               "-ConfigPath", str(directory / "config.json")]
    try:
        result = _run_worker(command, directory, stage_timeout, total_timeout)
    finally:
        (directory / "seed.xlsx").unlink(missing_ok=True)
    if result.get("status") != "accepted" and result.get("stage") in ("startup", "configure"):
        raise ExcelInfrastructureError("Excel failed during " + result.get("stage", "startup") +
                                       ". See " + str(directory / "worker.log"), result)
    return result


def _failure_details(phase):
    details = {key: phase[key] for key in ("reason", "stage", "error") if key in phase}
    message = str(phase.get("error", "")).lower()
    if any(word in message for word in ("password", "kennwort", "passwort")):
        details["reason"] = "password_rejected"
    return details


def convert_file(source: Path, destination: Path, attempt_dir: Path,
                 stage_timeout: float = 120, total_timeout: float = 360,
                 *, force_save: bool = False,
                 reconstruction_destination: Optional[Path] = None) -> Dict[str, Any]:
    """Convert without changing the source, then reopen in a fresh Excel process.

    Accepted means that this Excel installation opened the output without a
    blocking dialog under the recorded settings. It is not a fidelity claim.
    Keep attempt_dir on the data drive, with a short path for legacy Excel.
    """
    source, destination, attempt_dir = Path(source), Path(destination), Path(attempt_dir)
    if stage_timeout <= 0 or total_timeout <= 0:
        raise ValueError("Timeouts must be positive.")
    executable = _excel_executable()
    if not source.is_file():
        raise FileNotFoundError(source)
    if len(str(attempt_dir.resolve())) > 170:
        raise ValueError("Use an attempt directory shorter than 170 characters for Excel staging.")
    attempt_dir.mkdir(parents=True, exist_ok=True)
    start = time.monotonic()
    candidate = attempt_dir / "converted.xlsx"
    candidate.unlink(missing_ok=True)
    original_xlsx = _is_xlsx(source) and not force_save
    # Excel recognizes old formats from content. A short local name also avoids
    # long paths and archive characters that its file dialogs cannot handle.
    suffix = source.suffix if 0 < len(source.suffix) <= 8 else ".xls"
    local_source = attempt_dir / ("input" + suffix)
    result: Dict[str, Any] = {
        "status": "review", "source_sha256": _sha256(source),
        "method": "copy_xlsx" if original_xlsx else "excel_save_as_xlsx",
        "options": {"update_links": 0, "read_only": True, "calculation": "manual",
                    "macros": "disabled", "events": False, "repair": False,
                    "open_password": "explicit_test_value", "write_password": "explicit_test_value",
                    "conversion_local": True, "verification_open": "minimal_read_only",
                    "verify_display_alerts": True},
        "phases": {}, "visual_validation": "not_performed", "normalizations": [],
    }
    try:
        if original_xlsx:
            shutil.copyfile(source, candidate)
        else:
            shutil.copyfile(source, local_source)
            conversion = _invoke_phase("convert", local_source, candidate, attempt_dir / "convert",
                                       executable, stage_timeout, max(0.1, total_timeout - (time.monotonic() - start)))
            result["phases"]["convert"] = conversion
            if (conversion["status"] != "accepted" and conversion.get("stage") == "save"
                    and conversion.get("shared_workbook") and conversion.get("reason") == "excel_error"):
                # Retry only a shared-workbook save failure, within the same file deadline.
                remaining = total_timeout - (time.monotonic() - start)
                if remaining > 0:
                    candidate.unlink(missing_ok=True)
                    conversion = _invoke_phase("convert_exclusive", local_source, candidate,
                                               attempt_dir / "convert_exclusive", executable, stage_timeout, remaining)
                    result["phases"]["convert_exclusive"] = conversion
                    if conversion["status"] == "accepted":
                        result["normalizations"].append("legacy_shared_revision_history_removed")
            if conversion["status"] != "accepted":
                if (reconstruction_destination is not None and conversion.get('stage') == 'save'
                        and conversion.get('reason') == 'excel_error'):
                    recovery_target = Path(reconstruction_destination)
                    if recovery_target.resolve() in (source.resolve(), destination.resolve()):
                        raise ValueError('Reconstruction must use a separate output path.')
                    remaining = total_timeout - (time.monotonic() - start)
                    if remaining > 0:
                        candidate.unlink(missing_ok=True)
                        reconstruction = _invoke_phase('reconstruct', local_source, candidate,
                                                       attempt_dir / 'reconstruct', executable, stage_timeout, remaining)
                        result['phases']['reconstruct'] = reconstruction
                        result['reconstruction_trigger'] = _failure_details(conversion)
                        if reconstruction['status'] == 'accepted':
                            package_error = _check_xlsx(candidate)
                            remaining = total_timeout - (time.monotonic() - start)
                            if package_error:
                                result.update(reason='invalid_reconstructed_xlsx', stage='package_check', error=package_error)
                                return result
                            if remaining <= 0:
                                result.update(reason='total_timeout', stage='reconstruction_verify')
                                return result
                            verification = _invoke_phase('verify', candidate, candidate,
                                                         attempt_dir / 'reconstruction_verify', executable, stage_timeout, remaining)
                            result['phases']['reconstruction_verify'] = verification
                            if verification['status'] != 'accepted':
                                result.update(_failure_details(verification)); return result
                            if reconstruction.get('sheet_names') != verification.get('sheet_names'):
                                result.update(reason='reconstructed_sheet_list_changed', stage='reconstruction_verify'); return result
                            if _sha256(source) != result['source_sha256']:
                                result.update(reason='source_changed', stage='reconstruction_verify'); return result
                            recovery_target.parent.mkdir(parents=True, exist_ok=True)
                            pending = recovery_target.with_name(recovery_target.name + '.pending')
                            shutil.copyfile(candidate, pending); os.replace(str(pending), str(recovery_target))
                            result.update(status='reconstructed', method='excel_all_cells_copy_paste',
                                          reason='reconstructed_workbook_requires_review', stage='complete',
                                          output=str(recovery_target), output_sha256=_sha256(recovery_target),
                                          output_bytes=recovery_target.stat().st_size,
                                          reconstruction=reconstruction.get('reconstruction', {}),
                                          prompt_check='fresh_excel_instance_with_display_alerts',
                                          visual_validation='not_performed', eligible_for_regular_corpus=False)
                            return result
                        result.update(_failure_details(reconstruction)); return result
                result.update(_failure_details(conversion))
                return result
            result["macros_removed"] = conversion.get("macros_removed", False)
            result["excel"] = conversion.get("excel", {})
        container_error = _check_xlsx(candidate)
        if container_error:
            result.update(reason="invalid_xlsx", stage="package_check", error=container_error)
            return result
        remaining = total_timeout - (time.monotonic() - start)
        if remaining <= 0:
            result.update(reason="total_timeout", stage="verify", error="No time remains for the reopen check.")
            return result
        verification = _invoke_phase("verify", candidate, candidate, attempt_dir / "verify",
                                     executable, stage_timeout, remaining)
        result["phases"]["verify"] = verification
        if verification["status"] != "accepted":
            result.update(_failure_details(verification))
            return result
        if not original_xlsx and conversion.get("sheet_names") != verification.get("sheet_names"):
            result.update(reason="sheet_list_changed", stage="verify", error="Sheet names or order changed during reopening.")
            return result
        if "legacy_shared_revision_history_removed" in result["normalizations"] and verification.get("shared_workbook"):
            result.update(reason="sharing_still_enabled", stage="verify", error="The exclusive-mode copy is still shared.")
            return result
        if _sha256(source) != result["source_sha256"]:
            result.update(reason="source_changed", stage="verify", error="The source changed during conversion.")
            return result
        destination.parent.mkdir(parents=True, exist_ok=True)
        pending = destination.with_name(destination.name + ".pending")
        shutil.copyfile(candidate, pending)
        os.replace(str(pending), str(destination))
        result.update(status="accepted", reason="opened_without_blocking_dialog", stage="complete",
                      output_sha256=_sha256(destination), output_bytes=destination.stat().st_size,
                      excel=verification.get("excel", {}),
                      prompt_check="fresh_excel_instance_with_display_alerts")
        return result
    except ExcelInfrastructureError as error:
        result.update(reason="infrastructure_error", stage=error.details.get("stage", "startup"),
                      error=str(error), infrastructure=error.details)
        raise
    finally:
        result["elapsed_seconds"] = time.monotonic() - start
        _write_json(attempt_dir / "result.json", result)
        local_source.unlink(missing_ok=True)
        if result["status"] in ("accepted", "reconstructed"):
            candidate.unlink(missing_ok=True)
