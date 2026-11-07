import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from unittest import mock

from excel_shrink_research.enron import excel


class ExcelSupervisorTests(unittest.TestCase):
    def test_xlsx_detection_uses_content(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "workbook.unusual"
            excel._create_seed(source)
            self.assertTrue(excel._is_xlsx(source))
            self.assertIsNone(excel._check_xlsx(source))

    def test_stage_timeout_ends_worker_and_records_review(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "worker.py"
            script.write_text(
                "import json, pathlib, time\n"
                "p = pathlib.Path(__file__).parent\n"
                "(p / 'stage.json').write_text(json.dumps({'stage':'open'}))\n"
                "time.sleep(30)\n", encoding="utf-8")
            start = time.monotonic()
            result = excel._run_worker([sys.executable, str(script)], root, 0.4, 5)
            self.assertEqual("review", result["status"])
            self.assertEqual("stage_timeout", result["reason"])
            self.assertEqual("open", result["stage"])
            self.assertLess(time.monotonic() - start, 5)
            self.assertTrue((root / "windows.json").is_file())

    def test_absolute_deadline_survives_stage_changes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "worker.py"
            script.write_text(
                "import json, pathlib, time, os\n"
                "from excel_shrink_research.enron.excel import _write_json\n"
                "p = pathlib.Path(__file__).parent\n"
                "for i in range(200):\n"
                " _write_json(p / 'stage.json', {'stage':str(i)})\n"
                " time.sleep(.05)\n", encoding="utf-8")
            result = excel._run_worker([sys.executable, str(script)], root, 2, 0.5)
            self.assertEqual("total_timeout", result["reason"])

    def test_existing_xlsx_is_copied_without_conversion(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, destination = root / "input.other", root / "output.xlsx"
            excel._create_seed(source)
            original = source.read_bytes()
            with mock.patch.object(excel, "_excel_executable", return_value=Path("EXCEL.EXE")), \
                    mock.patch.object(excel, "_invoke_phase", return_value={"status": "accepted", "excel": {}}) as phase:
                result = excel.convert_file(source, destination, root / "attempt")
            self.assertEqual("accepted", result["status"])
            self.assertEqual("copy_xlsx", result["method"])
            self.assertEqual(original, destination.read_bytes())
            self.assertEqual(original, source.read_bytes())
            phase.assert_called_once()
            self.assertEqual("verify", phase.call_args[0][0])

    def test_rejected_output_is_not_published(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, destination = root / "input.xlsx", root / "output.xlsx"
            excel._create_seed(source)
            with mock.patch.object(excel, "_excel_executable", return_value=Path("EXCEL.EXE")), \
                    mock.patch.object(excel, "_invoke_phase", return_value={
                        "status": "review", "reason": "stage_timeout", "stage": "open"}):
                result = excel.convert_file(source, destination, root / "attempt")
            self.assertEqual("review", result["status"])
            self.assertFalse(destination.exists())
            self.assertTrue(source.exists())

    @unittest.skipUnless(os.name == "nt", "Windows process identity test")
    def test_cleanup_ignores_reused_pid(self):
        process = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"],
                                   creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            self.assertFalse(excel._terminate_owned({"pid": process.pid, "created_filetime": "1"}))
            self.assertIsNone(process.poll())
        finally:
            process.terminate()
            process.wait(timeout=5)

    @unittest.skipUnless(os.name == "nt", "Windows process identity test")
    def test_timeout_ends_only_the_owned_child(self):
        import ctypes
        from ctypes import wintypes

        unrelated = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"],
                                     creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            with tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                script = root / "worker.py"
                script.write_text(
                    "import ctypes, json, pathlib, subprocess, sys, time\n"
                    "from ctypes import wintypes\n"
                    "p = pathlib.Path(__file__).parent\n"
                    "child = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(30)'])\n"
                    "k = ctypes.WinDLL('kernel32', use_last_error=True)\n"
                    "k.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4\n"
                    "times = [wintypes.FILETIME() for _ in range(4)]\n"
                    "assert k.GetProcessTimes(int(child._handle), *(ctypes.byref(t) for t in times))\n"
                    "created = (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime\n"
                    "(p / 'process.json').write_text(json.dumps({'pid':child.pid, 'created_filetime':str(created)}))\n"
                    "(p / 'stage.json').write_text(json.dumps({'stage':'open'}))\n"
                    "time.sleep(30)\n", encoding="utf-8")
                result = excel._run_worker([sys.executable, str(script)], root, 0.5, 5)
                self.assertEqual("stage_timeout", result["reason"])
                self.assertIsNone(unrelated.poll())
                # GetExitCodeProcess also works while a terminated process handle
                # remains open, so this verifies cleanup without relying on PID reuse.
                handle = excel._owned_handle(result["process"], 0x00100000)
                if handle:
                    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
                    kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
                    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
                    try:
                        self.assertEqual(0, kernel.WaitForSingleObject(handle, 3000))
                    finally:
                        kernel.CloseHandle(handle)
        finally:
            unrelated.terminate()
            unrelated.wait(timeout=5)


if __name__ == "__main__":
    unittest.main()
