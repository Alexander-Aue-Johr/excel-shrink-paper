import hashlib
import io
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from excel_shrink_research import preparation_transfer as transfer


def digest(data):
    return hashlib.sha256(data).hexdigest()


def workbook(label):
    result = io.BytesIO()
    with zipfile.ZipFile(result, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<Types><Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>')
        archive.writestr("xl/workbook.xml", '<workbook name="' + label + '"/>')
    return result.getvalue()


class PreparationTransferTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.base = Path(temporary.name)
        self.source = self.base / "transfer"
        self.stage = self.base / "staged"
        self.study = self.base / "study"
        self.source.mkdir()
        self.profiles = {name: "hash-" + name for name in transfer.PROFILES}
        self.corpus = {"id": "c" * 64, "source_fingerprint": "f" * 64,
                       "catalog_sha256": "d" * 64, "archive_path": "corpus/source.zip", "total_workbooks": 2}
        self.build = "b" * 64
        source_archive = self.source / self.corpus["archive_path"]
        source_archive.parent.mkdir()
        self.originals = [b"first original source", b"second original source"]
        with zipfile.ZipFile(source_archive, "w") as archive:
            archive.writestr("first.xls", self.originals[0])
            archive.writestr("second.xls", self.originals[1])
        self.expected = {**self.corpus, "archive_sha256": digest(source_archive.read_bytes())}
        for directory in ("study/prepared", "study/records", "study/metadata/runs", "study/metadata/machines", "failures/originals"):
            (self.source / directory).mkdir(parents=True, exist_ok=True)
        (self.source / "study/metadata/runs/run.json").write_text('{"machine_id":"other"}')
        (self.source / "study/metadata/machines/machine.json").write_text('{"name":"other machine"}')
        self.records = []
        for index, name in enumerate(("first.xls", "second.xls")):
            identity = self.corpus["id"] + ":" + str(index) + ":" + name
            key = digest(identity.encode())[:20]
            record = {"identity": identity, "name": name, "corpus_id": self.corpus["id"],
                      "source_fingerprint": self.corpus["source_fingerprint"], "source_sha256": digest(self.originals[index]),
                      "prepared_by": self.build, "prepared": index == 0, "complete": index == 1,
                      "status": "prepared" if index == 0 else "skipped",
                      "run_selection": {**{field: self.corpus[field] for field in ("source_fingerprint", "catalog_sha256")},
                                        "corpus_id": self.corpus["id"], "archive_index": index,
                                        "size_part": {"part": 1, "parts": 1}, "rank": index + 1,
                                        "bytes": len(self.originals[index])},
                      "provenance": {"preparation": {"run_id": "run", "machine_manifest": "metadata/machines/machine.json"}},
                      "results": []}
            if index == 0:
                originals = {"original.xlsx": workbook("baseline")}
                originals.update({name + ".xlsx": workbook(name) for name in transfer.PROFILES})
                record["baseline_sha256"] = digest(originals["original.xlsx"])
                record["results"] = [{"profile": name, "config_hash": self.profiles[name],
                                      "status": "not compared", "checks_complete": False,
                                      "shrink_seconds": 0.25, "original_bytes": len(originals["original.xlsx"]),
                                      "shrunk_bytes": len(originals[name + ".xlsx"])} for name in transfer.PROFILES]
                with zipfile.ZipFile(self.source / ("study/prepared/" + key + ".zip"), "w") as bundle:
                    for member, data in originals.items():
                        bundle.writestr(member, data)
                    bundle.writestr("measurements.json", json.dumps(record["results"]))
            else:
                record["error"] = "Excel automation failed on the source machine."
                (self.source / "failures/originals/second.xls").write_bytes(self.originals[index])
            self.records.append((key, record))
            (self.source / ("study/records/" + key + ".json")).write_text(json.dumps(record))
        self.manifest = {"format": "excel-shrink-preparation-transfer", "schema_version": 1, "complete": True,
                         "contains_completed_comparison_claims": False, "corpus": self.corpus,
                         "size_part": {"part": 1, "parts": 1}, "profiles": self.profiles,
                         "preparation_builds": [self.build],
                         "counts": {"prepared": 1, "failed_preparation": 1, "records": 2}}
        self.publish_manifest()

    def publish_manifest(self):
        files = []
        for path in self.source.rglob("*"):
            if path.is_file() and path.name not in ("manifest.json", "manifest.sha256", "COPY_VERIFIED.txt"):
                data = path.read_bytes()
                files.append({"path": path.relative_to(self.source).as_posix(), "bytes": len(data), "sha256": digest(data)})
        self.manifest.update(files=files, total_payload_bytes=sum(row["bytes"] for row in files))
        self.write_manifest()

    def write_manifest(self):
        data = json.dumps(self.manifest).encode()
        (self.source / "manifest.json").write_bytes(data)
        (self.source / "manifest.sha256").write_text(digest(data) + "  manifest.json\n")
        (self.source / "COPY_VERIFIED.txt").write_text("The source machine completed the transfer.\n")

    def merge(self):
        return transfer.merge_transfer(self.stage, self.study, expected_source=self.expected, expected_profiles=self.profiles)

    def test_copies_all_payloads_and_reuses_verified_existing_files(self):
        first = transfer.stage_transfer(self.source, self.stage)
        self.assertEqual("staged", first["state"])
        self.assertEqual(len(self.manifest["files"]), first["files_copied"])
        second = transfer.stage_transfer(self.source, self.stage)
        self.assertEqual(0, second["files_copied"])
        for row in self.manifest["files"]:
            self.assertEqual(row["sha256"], digest((self.stage / row["path"]).read_bytes()))

    def test_corrupt_payload_never_gets_published_as_a_complete_staged_transfer(self):
        row = self.manifest["files"][0]
        (self.source / row["path"]).write_bytes(b"damaged")
        with self.assertRaisesRegex(ValueError, "manifest|recorded size"):
            transfer.stage_transfer(self.source, self.stage)
        status = json.loads((self.stage / "stage-status.json").read_text())
        self.assertEqual("failed", status["state"])
        self.assertFalse((self.stage / "COPY_VERIFIED.txt").exists())

    def test_rejects_parent_traversal_even_with_a_matching_manifest_checksum(self):
        self.manifest["files"][0]["path"] = "../outside.xls"
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, "Unsafe transfer path"):
            transfer.stage_transfer(self.source, self.stage)
        self.assertFalse((self.base / "outside.xls").exists())

    def test_rejects_incomplete_or_changed_manifest(self):
        (self.source / "TRANSFER_INCOMPLETE.txt").touch()
        with self.assertRaisesRegex(ValueError, "completely copied"):
            transfer.stage_transfer(self.source, self.stage)
        (self.source / "TRANSFER_INCOMPLETE.txt").unlink()
        (self.source / "manifest.json").write_text("{}")
        with self.assertRaisesRegex(ValueError, "manifest checksum"):
            transfer.stage_transfer(self.source, self.stage)

    def test_merge_uses_the_rust_output_lock_and_does_not_modify_active_study(self):
        transfer.stage_transfer(self.source, self.stage)
        with transfer._lock(self.study / ".run.lock", "test lock"):
            with self.assertRaises(transfer.StudyBusy):
                self.merge()
            self.assertFalse((self.study / "records").exists())

    def test_merges_paired_preparations_failures_and_original_provenance(self):
        transfer.stage_transfer(self.source, self.stage)
        result = self.merge()
        self.assertEqual((1, 1), (result["imported_prepared"], result["imported_failed"]))
        self.assertEqual([self.build], result["approved_builds"])
        for key, record in self.records:
            self.assertEqual(record, json.loads((self.study / ("records/" + key + ".json")).read_text()))
            self.assertEqual(record["prepared"], (self.study / ("prepared/" + key + ".zip")).exists())
        again = self.merge()
        self.assertEqual(0, again["imported_prepared"])
        self.assertEqual(2, again["retained_existing"])

    def test_existing_completed_visual_result_is_not_replaced(self):
        transfer.stage_transfer(self.source, self.stage)
        key, source_record = self.records[0]
        target = self.study / ("records/" + key + ".json")
        target.parent.mkdir(parents=True)
        record = {**source_record, "complete": True, "prepared": False, "status": "unchanged"}
        target.write_text(json.dumps(record))
        result = self.merge()
        self.assertEqual(1, result["retained_existing"])
        self.assertEqual(record, json.loads(target.read_text()))

    def test_failure_cannot_replace_an_existing_valid_preparation(self):
        transfer.stage_transfer(self.source, self.stage)
        self.merge()
        key, failed = self.records[1]
        existing = {**failed, "complete": True, "status": "unchanged"}
        target = self.study / ("records/" + key + ".json")
        target.write_text(json.dumps(existing))
        self.merge()
        self.assertEqual(existing, json.loads(target.read_text()))

    def test_metadata_conflict_stops_before_publishing_incoming_records(self):
        transfer.stage_transfer(self.source, self.stage)
        path = self.study / "metadata/runs/run.json"
        path.parent.mkdir(parents=True)
        path.write_text('{"different":"existing provenance"}')
        with self.assertRaisesRegex(ValueError, "immutable local metadata"):
            self.merge()
        self.assertFalse((self.study / "records").exists())

    def test_changed_staged_file_invalidates_cached_copy_verification(self):
        transfer.stage_transfer(self.source, self.stage)
        row = next(row for row in self.manifest["files"] if row["path"].startswith("study/records/"))
        (self.stage / row["path"]).write_text("damaged")
        with self.assertRaisesRegex(ValueError, "Payload differs"):
            self.merge()

    def test_profile_or_source_mismatch_prevents_merge(self):
        transfer.stage_transfer(self.source, self.stage)
        with self.assertRaisesRegex(ValueError, "profile configuration"):
            transfer.merge_transfer(self.stage, self.study, expected_source=self.expected, expected_profiles={})
        wrong = {**self.expected, "source_fingerprint": "a" * 64}
        with self.assertRaisesRegex(ValueError, "source_fingerprint"):
            transfer.merge_transfer(self.stage, self.study, expected_source=wrong, expected_profiles=self.profiles)

    def test_false_partition_rank_is_rejected_before_import(self):
        key, record = self.records[0]
        record["run_selection"]["rank"] = 2
        (self.source / ("study/records/" + key + ".json")).write_text(json.dumps(record))
        self.publish_manifest()
        transfer.stage_transfer(self.source, self.stage)
        with self.assertRaisesRegex(ValueError, "claimed size partition"):
            self.merge()
        self.assertFalse((self.study / "records").exists())

    def test_named_xlsx_that_is_not_a_workbook_is_rejected(self):
        key, record = self.records[0]
        path = self.source / ("study/prepared/" + key + ".zip")
        with zipfile.ZipFile(path) as bundle:
            contents = {name: bundle.read(name) for name in bundle.namelist()}
        contents["default.xlsx"] = b"not an XLSX package"
        with zipfile.ZipFile(path, "w") as bundle:
            for name, data in contents.items():
                bundle.writestr(name, data)
        self.publish_manifest()
        transfer.stage_transfer(self.source, self.stage)
        with self.assertRaises(zipfile.BadZipFile):
            self.merge()
        self.assertFalse((self.study / "records").exists())


if __name__ == "__main__":
    unittest.main()
