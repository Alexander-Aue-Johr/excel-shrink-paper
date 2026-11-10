import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from excel_shrink_research import preparation_queue as queue
from excel_shrink_research.enron.common import write_json


class QueueTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.receipt = {"Pid": 123, "StartedUtc": "2026-10-26T19:19:41.9018239Z", "OutputRoot": str(self.root)}
        self.run_id = "current-run"
        self.finished()

    def tearDown(self):
        self.temporary.cleanup()

    def finished(self, **overrides):
        room = self.root / "control-room"
        value = {"runner_pid": 123, "runner_started_unix_ms": queue.started_ms(self.receipt["StartedUtc"]),
                 "running": False, "phase": "finished", "run_id": self.run_id, "groups": []}
        value.update(overrides)
        write_json(room / "queue.json", value)
        write_json(room / "progress.json", {"run_id": self.run_id,
                   "scope": {"corpora": ["Enron"], "size_part": {"part": 1, "parts": 4}},
                   "counts": {"prepared": 2, "skipped": 1, "total": 3,
                              "remaining": 3, "in_progress": 0, "preparation_remaining": 0}})

    def test_preparation_outcomes_do_not_require_visual_completions(self):
        result = queue.completion(self.root, self.receipt, "Enron", 1)
        self.assertEqual(2, result["counts"]["prepared"])
        self.assertEqual(3, result["counts"]["remaining"])

    def test_stopped_or_replaced_runner_cannot_advance_the_queue(self):
        for changes in ({"phase": "stopped"}, {"runner_pid": 999},
                        {"runner_started_unix_ms": 0}, {"running": True},
                        {"groups": [{"active": True}]}):
            with self.subTest(changes=changes):
                self.finished(**changes)
                with self.assertRaises(RuntimeError):
                    queue.completion(self.root, self.receipt, "Enron", 1)

    def test_incomplete_preparation_or_foreign_progress_cannot_advance(self):
        path = self.root / "control-room/progress.json"
        for change in ("pending", "foreign", "selection"):
            self.finished()
            value = json.loads(path.read_text())
            if change == "pending":
                value["counts"]["preparation_remaining"] = 1
            elif change == "foreign":
                value["run_id"] = "old-run"
            else:
                value["scope"]["size_part"] = None
            write_json(path, value)
            with self.assertRaises(RuntimeError):
                queue.completion(self.root, self.receipt, "Enron", 1)

    def test_a_locked_study_cannot_advance(self):
        with patch.object(queue, "output_idle", return_value=False):
            with self.assertRaises(RuntimeError):
                queue.completion(self.root, self.receipt, "Enron", 1)

    def test_preflight_blocks_reconversion_of_compatible_saved_work(self):
        producer = "a" * 64
        profiles = {"default": "hash"}
        record = {"identity": "enron:1:a.xls", "prepared": True, "prepared_by": producer,
                  "source_fingerprint": "source", "complete": False,
                  "results": [{"profile": "default", "config_hash": "hash"}]}
        write_json(self.root / "records/one.json", record)
        (self.root / "prepared").mkdir()
        (self.root / "prepared/one.zip").write_bytes(b"already validated bundle")
        check = lambda: queue.check_existing_preparations(self.root, "enron", "source", profiles, {producer}, producer)
        self.assertEqual(1, check())
        record["results"][0]["config_hash"] = "changed"
        write_json(self.root / "records/one.json", record)
        with self.assertRaises(RuntimeError):
            check()

    def test_existing_receipt_is_reused_instead_of_starting_a_duplicate(self):
        write_json(self.root / "jobs/B/receipt.json", self.receipt)
        with patch.object(queue.subprocess, "run", side_effect=AssertionError("Duplicate launch")):
            self.assertEqual(self.receipt, queue.launch_job({}, self.root, "B", "Enron", "output", 8771, set()))

    def test_dependency_order_and_saved_checkpoints(self):
        config = {"queue_root": str(self.root / "queue"), "initial_run": self.receipt,
                  "transfer_status": str(self.root / "copy-status.json"),
                  "staged_transfer": str(self.root / "transfer"), "enron_output": "enron-output",
                  "source": {"id": "enron", "source_fingerprint": "source"}, "profiles": {},
                  "approved_builds": [], "executable_sha256": "a" * 64,
                  "euses_root": "euses-data", "euses_output": "euses-output"}
        write_json(self.root / "copy-status.json", {"state": "staged"})
        write_json(self.root / "transfer/manifest.json", {"preparation_builds": []})
        write_json(self.root / "config.json", config)
        calls = []
        def wait(receipt, corpus, partition, report, stop):
            calls.append("A" if partition else "B" if corpus == "Enron" else "D")
            return {"counts": {"preparation_remaining": 0}}
        def merge(*args, **kwargs):
            calls.append("import")
            return {"conflicts": []}
        def launch(config, root, name, *args, **kwargs):
            calls.append("launch_" + name)
            return self.receipt
        with patch.object(queue, "wait_run", side_effect=wait), \
             patch.object(queue, "wait_download", side_effect=lambda *args: calls.append("wait_C")), \
             patch.object(queue, "check_existing_preparations", return_value=1), \
             patch.object(queue, "launch_job", side_effect=launch), \
             patch("excel_shrink_research.preparation_transfer.merge_transfer", side_effect=merge), \
             patch("excel_shrink_research.euses.download_euses", side_effect=lambda *args, **kwargs: calls.append("C")), \
             patch("excel_shrink_research.euses.prepare_runner_archive", return_value={"path": "inputs.zip"}):
            self.assertEqual(0, queue.run_queue(self.root / "config.json"))
            self.assertEqual(["A", "import", "launch_B", "B", "wait_C", "C", "launch_D", "D"], calls)
            calls.clear()
            self.assertEqual(0, queue.run_queue(self.root / "config.json"))
            self.assertEqual([], calls)


if __name__ == "__main__":
    unittest.main()
