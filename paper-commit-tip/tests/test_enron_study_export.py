import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile


MODULE_PATH = Path(__file__).parents[1] / "src/excel_shrink_research/enron/study_export.py"
SPEC = importlib.util.spec_from_file_location("study_export", MODULE_PATH)
study_export = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(study_export)


class StudyExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.study = self.root / "study"
        self.cache = self.root / "cache"
        (self.study / "prepared").mkdir(parents=True)
        (self.study / "records").mkdir()

    def prepare(self, name="one", source="a" * 64, *, repaired=False, suffix=b""):
        original = b"original" + suffix
        payload = self.study / "prepared" / (name + ".zip")
        with zipfile.ZipFile(payload, "w") as bundle:
            for member in study_export._REQUIRED:
                bundle.writestr(member, original if member == "original.xlsx" else member.encode() + suffix)
            bundle.writestr("measurements.json", "[]")
        record = {
            "prepared": True,
            "source_sha256": source,
            "baseline_sha256": hashlib.sha256(original).hexdigest(),
            "baseline_repaired": repaired,
            "provenance": {"schema_version": 1, "preparation": {"run_id": "test-run"}},
        }
        (self.study / "records" / (name + ".json")).write_text(json.dumps(record), encoding="utf-8")
        return payload, record

    def entries(self):
        return list((self.cache / "entries").glob("*/*.json"))

    def test_exports_hashes_record_and_payload_without_modifying_source(self):
        payload, record = self.prepare(repaired=True)
        original = payload.read_bytes()
        summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(summary["exported"], 1)
        entry = json.loads(self.entries()[0].read_text(encoding="utf-8"))
        self.assertEqual(entry["record"], record)
        self.assertTrue(entry["baseline_repaired"])
        self.assertEqual(entry["source_sha256"], record["source_sha256"])
        self.assertEqual((self.cache / entry["bundle_path"]).read_bytes(), original)
        self.assertEqual(entry["bundle_sha256"], hashlib.sha256(original).hexdigest())
        self.assertEqual(len(entry["files"]), 4)
        self.assertEqual(payload.read_bytes(), original)

    def test_resume_does_not_hash_or_copy_unchanged_payloads(self):
        self.prepare()
        study_export.export_study(self.study, self.cache)
        with patch.object(study_export, "_hash_file", side_effect=AssertionError("rehashed")):
            summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(summary["existing"], 1)
        self.assertEqual(summary["errors"], [])

    def test_same_source_keeps_distinct_preparations(self):
        self.prepare("first")
        self.prepare("second", suffix=b"changed")
        summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(summary["exported"], 2)
        self.assertEqual(len(self.entries()), 2)

    def test_rejects_baseline_mismatch_without_publishing(self):
        self.prepare()
        path = self.study / "records/one.json"
        record = json.loads(path.read_text())
        record["baseline_sha256"] = "b" * 64
        path.write_text(json.dumps(record), encoding="utf-8")
        summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(len(summary["errors"]), 1)
        self.assertEqual(self.entries(), [])
        self.assertEqual(list((self.cache / ".pending").glob("*")), [])

    def test_incomplete_record_is_retried(self):
        self.prepare()
        path = self.study / "records/one.json"
        record = json.loads(path.read_text())
        record["prepared"] = False
        path.write_text(json.dumps(record), encoding="utf-8")
        summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(summary["pending"], 1)
        self.assertEqual(self.entries(), [])
        record["prepared"] = True
        path.write_text(json.dumps(record), encoding="utf-8")
        self.assertEqual(study_export.export_study(self.study, self.cache)["exported"], 1)

    def test_missing_cached_payload_is_restored(self):
        self.prepare()
        study_export.export_study(self.study, self.cache)
        entry = json.loads(self.entries()[0].read_text(encoding="utf-8"))
        (self.cache / entry["bundle_path"]).unlink()
        self.assertEqual(study_export.export_study(self.study, self.cache)["exported"], 1)
        self.assertTrue((self.cache / entry["bundle_path"]).is_file())

    def test_optional_conversion_is_preserved(self):
        payload, _ = self.prepare()
        with zipfile.ZipFile(payload, "a") as bundle:
            bundle.writestr("converted.xlsx", b"pure conversion")
            bundle.writestr("preparation.json", "{}")
        study_export.export_study(self.study, self.cache)
        entry = json.loads(self.entries()[0].read_text(encoding="utf-8"))
        names = {item["name"] for item in entry["files"]}
        self.assertIn("converted.xlsx", names)
        self.assertIn("preparation.json", names)

    def test_incomplete_bundle_never_produces_entry(self):
        payload, _ = self.prepare()
        payload.write_bytes(b"partial zip")
        summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(len(summary["errors"]), 1)
        self.assertEqual(self.entries(), [])

    def test_bundle_replacement_during_copy_is_retried(self):
        payload, _ = self.prepare()
        copyfile = study_export.shutil.copyfile

        def replace_after_copy(source, destination):
            result = copyfile(source, destination)
            payload.write_bytes(b"new bundle")
            return result

        with patch.object(study_export.shutil, "copyfile", side_effect=replace_after_copy):
            summary = study_export.export_study(self.study, self.cache)
        self.assertEqual(summary["pending"], 1)
        self.assertEqual(self.entries(), [])

    def test_stop_file_prevents_watch_scan(self):
        stop_file = self.root / "stop"
        stop_file.touch()
        with patch.object(study_export, "export_study", side_effect=AssertionError("scanned")):
            study_export.watch_study(self.study, self.cache, stop_file)

    def test_snapshot_survives_original_bundle_cleanup(self):
        payload, record = self.prepare()
        snapshot = self.root / "snapshot"
        result = study_export.snapshot_study(self.study, snapshot)
        self.assertEqual(result["captured"], 1)
        payload.unlink()
        (self.study / "records/one.json").unlink()
        summary = study_export.export_study(snapshot, self.cache)
        self.assertEqual(summary["exported"], 1)
        entry = json.loads(self.entries()[0].read_text(encoding="utf-8"))
        self.assertEqual(entry["record"], record)
        self.assertEqual(entry["provenance"]["study_root"], str(self.study.resolve()))

    def test_snapshot_keeps_replaced_bundle_versions(self):
        payload, _ = self.prepare()
        snapshot = self.root / "snapshot"
        study_export.snapshot_study(self.study, snapshot)
        first_bytes = next((snapshot / "prepared").glob("*.zip")).read_bytes()
        payload.unlink()  # The study publishes a fresh file, never edits in place.
        self.prepare(suffix=b"new preparation")
        result = study_export.snapshot_study(self.study, snapshot)
        self.assertEqual(result["captured"], 1)
        self.assertEqual(result["existing"], 0)
        bundles = list((snapshot / "prepared").glob("*.zip"))
        self.assertEqual(len(bundles), 2)
        self.assertIn(first_bytes, [path.read_bytes() for path in bundles])
        self.assertEqual(study_export.export_study(snapshot, self.cache)["exported"], 2)

    def test_snapshot_falls_back_to_copy_and_preserves_metadata(self):
        self.prepare()
        manifests = self.study / "metadata/runs"
        manifests.mkdir(parents=True)
        (manifests / "test-run.json").write_text('{"run_id":"test-run"}', encoding="utf-8")
        snapshot = self.root / "snapshot"
        with patch.object(study_export.os, "link", side_effect=OSError("not supported")):
            result = study_export.snapshot_study(self.study, snapshot)
        self.assertEqual(result["captured"], 1)
        self.assertEqual(result["errors"], [])
        self.assertTrue((snapshot / "metadata/runs/test-run.json").is_file())
        with patch.object(study_export.shutil, "copy2", side_effect=AssertionError("recopied")):
            result = study_export.snapshot_study(self.study, snapshot)
        self.assertEqual(result["existing"], 1)

    def test_watch_snapshots_before_indexing(self):
        stop_file = self.root / "stop"
        calls = []

        def snapshot(source, destination):
            calls.append("snapshot")
            return {"captured": 1, "errors": []}

        def export(source, destination, on_progress=None):
            calls.append("export")
            stop_file.touch()
            return {"exported": 0, "errors": []}

        with patch.object(study_export, "snapshot_study", side_effect=snapshot), \
             patch.object(study_export, "export_study", side_effect=export), \
             patch.object(study_export.time, "sleep"):
            study_export.watch_study(self.study, self.cache, stop_file)
        self.assertEqual(calls, ["snapshot", "export"])


if __name__ == "__main__":
    unittest.main()
