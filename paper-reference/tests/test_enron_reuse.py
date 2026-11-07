import hashlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import zipfile

from excel_shrink_research.enron.common import sha256, write_json
from excel_shrink_research.enron.reuse import PROFILES, reuse_file


def workbook(marker):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '</Types>')
        archive.writestr("xl/workbook.xml", '<workbook marker="' + marker + '"/>')
    return stream.getvalue()


class ReuseTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.cache = self.root / "cache"
        self.source_sha256 = hashlib.sha256(b"original legacy bytes").hexdigest()
        self.destination = self.root / "output/xlsx/archive/renamed.xls.xlsx"
        self.variants_root = self.root / "output/variants"

    def preparation(self, marker="one", missing=(), converted=False, embedded_record=False):
        payloads = {name: workbook(marker + "-" + name) for name in
                    ("original.xlsx",) + tuple(profile + ".xlsx" for profile in PROFILES)}
        if converted:
            payloads["converted.xlsx"] = workbook(marker + "-conversion")
        if embedded_record:
            payloads["preparation.json"] = b'{"producer": "historical-run"}'
        for name in missing:
            payloads.pop(name)
        pending = self.root / (marker + ".zip")
        with zipfile.ZipFile(pending, "w") as archive:
            for name, data in payloads.items():
                archive.writestr(name, data)
            archive.writestr("measurements.json", "[]")
        digest = sha256(pending)
        bundle = self.cache / "bundles" / (digest + ".zip")
        bundle.parent.mkdir(parents=True, exist_ok=True)
        pending.replace(bundle)
        record = {
            "source_sha256": self.source_sha256, "prepared": True,
            "baseline_sha256": hashlib.sha256(payloads["original.xlsx"]).hexdigest(),
            "baseline_repaired": True,
            "member": "unrelated/source_name.xls",
            "results": [{"profile": profile, "config_hash": marker + "-" + profile}
                        for profile in PROFILES],
        }
        entry = {
            "schema_version": 1, "source_sha256": self.source_sha256,
            "bundle_sha256": digest, "bundle_path": bundle.relative_to(self.cache).as_posix(),
            "record": record, "provenance": {"producer": "historical-run-" + marker},
            "files": [{"name": name, "bytes": len(data),
                       "sha256": hashlib.sha256(data).hexdigest()} for name, data in payloads.items()],
        }
        entry_path = self.cache / "entries" / self.source_sha256 / (digest + ".json")
        write_json(entry_path, entry)
        return entry_path, entry, payloads

    def reuse(self, source_sha256=None, **kwargs):
        return reuse_file(source_sha256 or self.source_sha256, self.destination,
            self.variants_root, self.root / "attempt", [self.cache],
            relative_path="archive/renamed.xls.xlsx", **kwargs)

    def test_exact_source_match_reuses_all_variants_without_excel(self):
        _, entry, payloads = self.preparation()
        with patch("excel_shrink_research.enron.excel.convert_file",
                   side_effect=AssertionError("Reuse must not invoke Excel")):
            result = self.reuse()
        self.assertEqual(result["status"], "accepted")
        self.assertEqual(self.destination.read_bytes(), payloads["original.xlsx"])
        self.assertTrue(result["baseline_repaired"])
        self.assertEqual(result["prompt_check"], "historical_study_baseline_not_freshly_reopened")
        self.assertEqual(result["reuse"]["record"], entry["record"])
        for profile in PROFILES:
            output = result["variants"][profile]
            self.assertEqual(Path(output["path"]).read_bytes(), payloads[profile + ".xlsx"])
            self.assertEqual(output["sha256"], sha256(Path(output["path"])))
            self.assertEqual(output["config_hash"], "one-" + profile)
        self.assertTrue((self.root / "attempt/reuse.json").is_file())

    def test_different_source_bytes_are_a_cache_miss(self):
        self.preparation()
        self.assertIsNone(self.reuse(hashlib.sha256(b"different").hexdigest()))
        self.assertFalse(self.destination.exists())

    def test_bundle_tampering_requires_review(self):
        _, entry, _ = self.preparation()
        bundle = self.cache / entry["bundle_path"]
        with bundle.open("ab") as stream:
            stream.write(b"changed")
        result = self.reuse()
        self.assertEqual(result["status"], "review")
        self.assertIn("bundle", result["error"])
        self.assertFalse(self.destination.exists())

    def test_workbook_metadata_tampering_requires_review(self):
        path, entry, _ = self.preparation()
        entry["files"][1]["sha256"] = "0" * 64
        write_json(path, entry)
        result = self.reuse()
        self.assertEqual(result["status"], "review")
        self.assertIn("checksum", result["error"])
        self.assertFalse(self.destination.exists())

    def test_different_recorded_baseline_requires_review(self):
        path, entry, _ = self.preparation()
        entry["record"]["baseline_sha256"] = "0" * 64
        write_json(path, entry)
        result = self.reuse()
        self.assertEqual(result["status"], "review")
        self.assertIn("baseline", result["error"])

    def test_missing_profile_is_not_silently_regenerated(self):
        self.preparation(missing=("maximum-cleaning.xlsx",))
        result = self.reuse()
        self.assertEqual(result["status"], "review")
        self.assertIn("all three", result["error"])
        self.assertFalse(self.destination.exists())

    def test_conflicts_choose_one_complete_bundle_deterministically(self):
        _, first, first_payloads = self.preparation("first")
        _, second, second_payloads = self.preparation("second")
        expected = min((first, first_payloads), (second, second_payloads),
                       key=lambda item: item[0]["bundle_sha256"])
        result = self.reuse()
        self.assertEqual(result["status"], "accepted")
        self.assertEqual(result["reuse"]["candidate_count"], 2)
        self.assertEqual(result["reuse"]["bundle_sha256"], expected[0]["bundle_sha256"])
        for profile, output in result["variants"].items():
            self.assertEqual(Path(output["path"]).read_bytes(), expected[1][profile + ".xlsx"])

    def test_conversion_before_baseline_repair_is_kept_separately(self):
        _, _, payloads = self.preparation(converted=True)
        result = self.reuse()
        self.assertEqual(result["status"], "accepted")
        self.assertEqual(self.destination.read_bytes(), payloads["original.xlsx"])
        conversion = result["conversion_artifact"]
        self.assertEqual(Path(conversion["path"]).read_bytes(), payloads["converted.xlsx"])
        self.assertNotEqual(result["output_sha256"], conversion["sha256"])

    def test_embedded_preparation_record_is_accepted(self):
        self.preparation(embedded_record=True)
        self.assertEqual(self.reuse()["status"], "accepted")

    def test_unknown_historical_repair_state_stays_unknown(self):
        path, entry, _ = self.preparation()
        del entry["record"]["baseline_repaired"]
        write_json(path, entry)
        self.assertIsNone(self.reuse()["baseline_repaired"])

    def test_unsafe_bundle_path_requires_review(self):
        path, entry, _ = self.preparation()
        entry["bundle_path"] = "../outside.zip"
        write_json(path, entry)
        self.assertEqual(self.reuse()["status"], "review")

    def test_unsafe_output_path_is_rejected_before_import(self):
        self.preparation()
        with self.assertRaisesRegex(ValueError, "within"):
            reuse_file(self.source_sha256, self.destination, self.variants_root,
                self.root / "attempt", [self.cache], relative_path="../escape.xlsx")

    def test_changed_existing_variant_is_preserved_for_review(self):
        self.preparation()
        target = self.variants_root / "maximum-cleaning/archive/renamed.xls.xlsx"
        target.parent.mkdir(parents=True)
        target.write_bytes(b"changed")
        result = self.reuse()
        self.assertEqual(result["status"], "review")
        self.assertEqual(target.read_bytes(), b"changed")
        self.assertFalse(self.destination.exists())


if __name__ == "__main__":
    unittest.main()
