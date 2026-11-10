import copy
import json
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from excel_shrink_research.enron.benchmark import _validate_batch_report
from excel_shrink_research.enron.common import long_path, sha256, write_json
from excel_shrink_research.enron.release import accepted_outputs, package


class ReleaseTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name).resolve()
        output = self.root / "xlsx/archive/a.xls.xlsx"
        output.parent.mkdir(parents=True)
        output.write_bytes(b"converted workbook")
        self.relative = output.relative_to(self.root).as_posix()
        self.digest = sha256(output)
        self.inventory = [
            {"id": "first", "sha256": "source-one", "classification": "candidate", "member": "a.xls"},
            {"id": "duplicate", "sha256": "source-one", "classification": "candidate", "member": "b.xls"},
            {"id": "review", "sha256": "source-two", "classification": "ambiguous", "member": "unknown"},
            {"id": "mail", "classification": "other", "member": "message.txt"},
        ]
        accepted = {"status": "accepted", "output_path": self.relative, "output_sha256": self.digest}
        self.conversions = [
            {"source": self.inventory[0], "result": accepted},
            {"source": self.inventory[1], "result": accepted},
            {"source": self.inventory[2], "result": {"status": "review", "reason": "format_review"}},
        ]
        write_json(self.root / "sources.json", {"archives": ["fixture.zip"]})
        self.write_rows("inventory.jsonl", self.inventory)
        self.write_rows("conversions.jsonl", self.conversions)
        self.summary = {
            "limited": False, "inventory_entries": 4, "candidate_occurrences": 3,
            "accepted_occurrences": 2, "review_occurrences": 1, "unique_accepted": 1,
            "scope": "Test source collection",
        }
        self.seal_summary()

    def write_rows(self, name, records):
        (self.root / name).write_text("".join(json.dumps(row) + "\n" for row in records), encoding="utf-8")

    def seal_summary(self):
        for name, key in (
            ("sources.json", "source_manifest_sha256"),
            ("inventory.jsonl", "inventory_sha256"),
            ("conversions.jsonl", "conversion_manifest_sha256"),
        ):
            self.summary[key] = sha256(self.root / name)
        write_json(self.root / "summary.json", self.summary)

    def test_duplicate_contents_keep_all_occurrences(self):
        self.assertEqual(accepted_outputs(self.root), {self.relative: self.digest})

    def test_changed_manifest_is_rejected(self):
        self.write_rows("conversions.jsonl", self.conversions[:1])
        with self.assertRaisesRegex(ValueError, "metadata changed"):
            accepted_outputs(self.root)

    def test_missing_duplicate_occurrence_is_rejected_even_when_file_count_matches(self):
        self.write_rows("conversions.jsonl", [self.conversions[0], self.conversions[2]])
        self.seal_summary()
        with self.assertRaisesRegex(ValueError, "missing inventory occurrences"):
            accepted_outputs(self.root)

    def test_repeated_occurrence_cannot_hide_a_missing_one(self):
        self.write_rows("conversions.jsonl", [self.conversions[0], self.conversions[0], self.conversions[2]])
        self.seal_summary()
        with self.assertRaisesRegex(ValueError, "repeats or adds"):
            accepted_outputs(self.root)

    def test_changed_source_mapping_is_rejected(self):
        changed = copy.deepcopy(self.conversions)
        changed[0]["source"]["member"] = "different.xls"
        self.write_rows("conversions.jsonl", changed)
        self.seal_summary()
        with self.assertRaisesRegex(ValueError, "source differs"):
            accepted_outputs(self.root)

    def test_summary_cannot_hide_review_occurrences(self):
        self.summary["review_occurrences"] = 0
        self.seal_summary()
        with self.assertRaisesRegex(ValueError, "counts"):
            accepted_outputs(self.root)

    def test_package_contains_code_and_run_provenance(self):
        write_json(self.root / "state/code/version.json", {"excel.py": "code hash"})
        write_json(self.root / "runs/run-1/run.json", {"commit": "commit hash"})
        (self.root / "runs/run-1/worker.log").write_text("verbose output", encoding="utf-8")
        destination = self.root / "release.zip"
        package(self.root, destination)
        with zipfile.ZipFile(destination) as archive:
            names = set(archive.namelist())
            self.assertIn("metadata/code/version.json", names)
            self.assertIn("metadata/runs/run-1/run.json", names)
            self.assertFalse(any(name.endswith("worker.log") for name in names))

    def batch_report(self):
        run = self.root / "benchmark/01-default"
        output = run / self.relative
        long_path(output.parent).mkdir(parents=True, exist_ok=True)
        long_path(output).write_bytes(b"shrunk workbook")
        report = {
            "schema_version": 1, "elapsed_ms": 10,
            "summary": {"total": 1, "processed": 1, "preserved": 0, "skipped": 0, "failed": 0},
            "files": [{"input": str(self.root / self.relative), "output": str(output), "status": "processed"}],
        }
        return run, report

    def test_long_source_paths_survive_release_and_batch_accounting(self):
        self.assertEqual(self.root.parent, Path(tempfile.gettempdir()).resolve())
        self.addCleanup(shutil.rmtree, long_path(self.root))
        previous = self.root / self.relative
        self.relative = "xlsx/" + "/".join(["archive-directory-" + "x" * 30] * 6) + "/book.xlsx"
        destination = long_path(self.root / self.relative)
        destination.parent.mkdir(parents=True)
        destination.write_bytes(previous.read_bytes())
        previous.unlink()
        for row in self.conversions:
            if row["result"]["status"] == "accepted":
                row["result"]["output_path"] = self.relative
        self.write_rows("conversions.jsonl", self.conversions)
        self.seal_summary()
        self.assertEqual(accepted_outputs(self.root), {self.relative: self.digest})
        archive = self.root / "long-paths.zip"
        package(self.root, archive)
        with zipfile.ZipFile(archive) as package_file:
            self.assertIn(self.relative, package_file.namelist())
        run, report = self.batch_report()
        _validate_batch_report(report, self.root, run, {self.relative: self.digest})

    def test_batch_report_requires_exact_input_output_mapping(self):
        run, report = self.batch_report()
        inputs = {self.relative: self.digest}
        _validate_batch_report(report, self.root, run, inputs)
        report["files"][0]["output"] = str(run / "xlsx/wrong.xlsx")
        with self.assertRaisesRegex(RuntimeError, "unexpected input/output mapping"):
            _validate_batch_report(report, self.root, run, inputs)

    def test_batch_report_rejects_skipped_files_and_extra_outputs(self):
        run, report = self.batch_report()
        inputs = {self.relative: self.digest}
        report["files"][0]["status"] = "skipped"
        with self.assertRaisesRegex(RuntimeError, "skipped or failed"):
            _validate_batch_report(report, self.root, run, inputs)
        report["files"][0]["status"] = "processed"
        (run / "xlsx/extra.xlsx").write_bytes(b"unaccounted output")
        with self.assertRaisesRegex(RuntimeError, "output directory differs"):
            _validate_batch_report(report, self.root, run, inputs)


if __name__ == "__main__":
    unittest.main()
