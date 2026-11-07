import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from excel_shrink_research.enron.common import RunLock, sha256, write_json
from excel_shrink_research.enron.release import accepted_outputs, package
from excel_shrink_research.enron.workflow import Progress, convert_inventory


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.config = {"stage_timeout_seconds": 2, "file_timeout_seconds": 5}

    def inventory(self, entries):
        records = []
        for index, (name, data, classification) in enumerate(entries):
            path = self.root / "extracted" / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)
            records.append({"id": str(index), "archive": "fixture.zip", "member": name,
                            "index": index, "size": len(data), "sha256": sha256(path),
                            "classification": classification, "kind": "xls",
                            "extracted_path": path.relative_to(self.root).as_posix()})
        path = self.root / "inventory.jsonl"
        path.write_text("".join(json.dumps(row) + "\n" for row in records), encoding="utf-8")
        return path

    @staticmethod
    def convert(source, destination, attempt, **kwargs):
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(b"converted:" + source.read_bytes())
        return {"status": "accepted"}

    def test_duplicate_contents_keep_both_source_occurrences_and_resume(self):
        inventory = self.inventory([("a/report.xls", b"one", "candidate"), ("b/report.xls", b"one", "candidate")])
        with patch("excel_shrink_research.enron.workflow.convert_file", side_effect=self.convert) as convert:
            summary = convert_inventory(self.root, inventory, self.config, Progress(self.root))
            self.assertEqual(convert.call_count, 1)
            self.assertEqual(summary["accepted_occurrences"], 2)
            self.assertEqual(summary["unique_accepted"], 1)
            convert_inventory(self.root, inventory, self.config, Progress(self.root))
            self.assertEqual(convert.call_count, 1)
        mappings = (self.root / "conversions.jsonl").read_text().splitlines()
        self.assertEqual(len(mappings), 2)
        self.assertEqual(json.loads(mappings[0])["result"]["output_path"], json.loads(mappings[1])["result"]["output_path"])

    def test_problem_duplicates_are_copied_to_review_at_each_source_path(self):
        inventory = self.inventory([("a/report.xls", b"bad", "candidate"), ("b/report.xls", b"bad", "candidate")])
        with patch("excel_shrink_research.enron.workflow.convert_file", return_value={"status": "review", "reason": "timeout"}):
            summary = convert_inventory(self.root, inventory, self.config, Progress(self.root))
        self.assertEqual(summary["review_occurrences"], 2)
        self.assertEqual((self.root / "review/files/a/report.xls").read_bytes(), b"bad")
        self.assertEqual((self.root / "review/files/b/report.xls").read_bytes(), b"bad")
        self.assertTrue((self.root / "extracted/a/report.xls").exists())

    def test_changed_output_prevents_resume_and_release(self):
        inventory = self.inventory([("one.xls", b"one", "candidate")])
        with patch("excel_shrink_research.enron.workflow.convert_file", side_effect=self.convert):
            summary = convert_inventory(self.root, inventory, self.config, Progress(self.root))
        write_json(self.root / "sources.json", {"archives": []})
        summary.update(inventory_sha256=sha256(inventory),
                       conversion_manifest_sha256=sha256(self.root / "conversions.jsonl"),
                       source_manifest_sha256=sha256(self.root / "sources.json"))
        write_json(self.root / "summary.json", summary)
        (self.root / "xlsx/one.xls.xlsx").write_bytes(b"changed")
        with self.assertRaisesRegex(ValueError, "changed"):
            convert_inventory(self.root, inventory, self.config, Progress(self.root))
        with self.assertRaisesRegex(ValueError, "changed"):
            accepted_outputs(self.root)

    def test_pilot_does_not_publish_a_partial_manifest(self):
        inventory = self.inventory([("one.xls", b"one", "candidate"), ("two.xls", b"two", "candidate")])
        with patch("excel_shrink_research.enron.workflow.convert_file", side_effect=self.convert):
            summary = convert_inventory(self.root, inventory, self.config, Progress(self.root), limit=1)
        self.assertTrue(summary["limited"])
        self.assertFalse((self.root / "conversions.jsonl").exists())

    def test_lock_prevents_concurrent_writers(self):
        with RunLock(self.root):
            with self.assertRaisesRegex(RuntimeError, "Another"):
                with RunLock(self.root):
                    pass

    def test_exact_source_reuse_skips_excel_and_keeps_all_profiles(self):
        inventory = self.inventory([("a/report.xls", b"one", "candidate"), ("b/other.xls", b"one", "candidate")])

        def reuse(digest, destination, variants_root, attempt, caches, **kwargs):
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(b"saved baseline")
            variants = {}
            for profile in ("default", "maximum-cleaning", "too-aggressive-cleaning"):
                path = variants_root / profile / kwargs["relative_path"]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(profile.encode())
                variants[profile] = {"path": str(path), "sha256": sha256(path)}
            return {"status": "accepted", "method": "reuse_study_preparation", "variants": variants,
                    "prompt_check": "historical_study_baseline_not_freshly_reopened"}

        with patch("excel_shrink_research.enron.workflow.reuse_file", side_effect=reuse) as saved:
            with patch("excel_shrink_research.enron.workflow.convert_file") as excel:
                summary = convert_inventory(self.root, inventory, self.config, Progress(self.root))
                self.assertEqual(saved.call_count, 1)
                excel.assert_not_called()
        self.assertEqual(summary["unique_reused"], 1)
        self.assertEqual(summary["reused_occurrences"], 2)
        self.assertEqual(summary["conversion_attempts"], 0)
        with patch("excel_shrink_research.enron.workflow.reuse_file") as saved:
            with patch("excel_shrink_research.enron.workflow.convert_file") as excel:
                convert_inventory(self.root, inventory, self.config, Progress(self.root))
                saved.assert_not_called()
                excel.assert_not_called()
        (self.root / "variants/default/a/report.xls.xlsx").write_bytes(b"changed")
        with self.assertRaisesRegex(ValueError, "variant changed"):
            convert_inventory(self.root, inventory, self.config, Progress(self.root))

    def test_broken_reuse_does_not_trigger_a_new_conversion(self):
        inventory = self.inventory([("one.xls", b"one", "candidate")])
        with patch("excel_shrink_research.enron.workflow.reuse_file", return_value={"status": "review", "reason": "reuse_invalid"}):
            with patch("excel_shrink_research.enron.workflow.convert_file") as excel:
                summary = convert_inventory(self.root, inventory, self.config, Progress(self.root))
                excel.assert_not_called()
        self.assertEqual(summary["review_occurrences"], 1)


if __name__ == "__main__":
    unittest.main()
