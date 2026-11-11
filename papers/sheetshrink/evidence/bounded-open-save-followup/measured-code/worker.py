"""Execute one bounded library round trip and publish phase checkpoints."""
import ctypes
import json
import os
from pathlib import Path
import sys
import time
import traceback


STATE = Path(sys.argv[3])


def publish(phase, **fields):
    data = dict(phase=phase, monotonic=time.perf_counter(), **fields)
    temporary = STATE.with_suffix(".pending")
    temporary.write_text(json.dumps(data), encoding="utf-8")
    for attempt in range(50):
        try:
            temporary.replace(STATE)
            return
        except PermissionError:
            if attempt == 49:
                raise
            time.sleep(0.01)


def main():
    library, source, _, destination = sys.argv[1:5]
    source, destination = Path(source), Path(destination)
    publish("importing")
    if library == "allocation-probe":
        publish("opening")
        chunks = []
        while True:
            chunks.append(bytearray(8 * 1024 * 1024))
    elif library == "openpyxl":
        import openpyxl
        run_round_trip(lambda: openpyxl.load_workbook(source), lambda book: book.save(destination))
    elif library in ("pandas", "pandas-csv"):
        import pandas as pd
        def save_frames(frames):
            if library == "pandas-csv":
                destination.mkdir(exist_ok=True)
                for index, frame in enumerate(frames.values()):
                    frame.to_csv(destination / f"sheet-{index + 1}.csv", index=False)
            else:
                with pd.ExcelWriter(destination, engine="openpyxl") as writer:
                    for name, frame in frames.items():
                        frame.to_excel(writer, sheet_name=name, index=False)
        run_round_trip(lambda: pd.read_excel(source, sheet_name=None), save_frames)
    elif library == "excel":
        run_excel(source, destination)
    else:
        raise ValueError(f"Unknown worker library: {library}")


def run_round_trip(opener, saver):
    publish("opening")
    started = time.perf_counter()
    value = opener()
    opened = time.perf_counter() - started
    publish("saving", open_seconds=opened)
    started = time.perf_counter()
    saver(value)
    saved = time.perf_counter() - started
    publish("completed", open_seconds=opened, save_seconds=saved)


def run_excel(source, destination):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    started = time.perf_counter()
    application = win32com.client.DispatchEx("Excel.Application")
    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(application.Hwnd, ctypes.byref(pid))
    publish("excel_binding", excel_pid=pid.value)
    permit = STATE.with_suffix(".permit")
    while not permit.exists():
        if time.perf_counter() - started > 30:
            raise TimeoutError("Excel ownership handshake timed out before opening")
        time.sleep(0.02)
    application.Visible = False
    application.DisplayAlerts = False
    application.AutomationSecurity = 3
    publish("opening", excel_pid=pid.value)
    workbook = application.Workbooks.Open(str(source.resolve()), UpdateLinks=0, ReadOnly=True,
                                         IgnoreReadOnlyRecommended=True, AddToMru=False)
    opened = time.perf_counter() - started
    publish("saving", excel_pid=pid.value, open_seconds=opened)
    started = time.perf_counter()
    workbook.SaveAs(str(destination.resolve()), FileFormat=51)
    saved = time.perf_counter() - started
    started = time.perf_counter()
    workbook.Close(False)
    application.Quit()
    workbook = None
    application = None
    pythoncom.CoUninitialize()
    publish("completed", open_seconds=opened, save_seconds=saved,
            close_quit_seconds=time.perf_counter() - started)


if __name__ == "__main__":
    package_root = Path(os.environ["BENCHMARK_PACKAGES"])
    sys.path[:0] = [str(package_root), str(package_root / "win32"), str(package_root / "win32/lib")]
    dll_path = os.add_dll_directory(str(package_root / "pywin32_system32"))
    try:
        main()
    except BaseException as error:
        previous = json.loads(STATE.read_text()) if STATE.exists() else {}
        publish("failed", failed_phase=previous.get("phase"), open_seconds=previous.get("open_seconds"),
                error_type=type(error).__name__, error=str(error), traceback=traceback.format_exc())
        raise
