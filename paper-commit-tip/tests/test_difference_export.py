import csv
import hashlib
import io
import json
from pathlib import Path
import tempfile
import unittest
import warnings
from unittest.mock import patch
import zipfile

from excel_shrink_research.enron import difference_export as export


def archive_bytes(entries):
    stream = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, content in entries:
                archive.writestr(name, content)
    return stream.getvalue()


def row(content, archive, names, indices):
    return {"sha256": hashlib.sha256(content).hexdigest(), "size": len(content), "kind": "xls",
            "classification": "candidate", "archive": archive,
            "member_chain": json.dumps(names), "index_chain": json.dumps(indices)}


class DifferenceExportTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.base = Path(self.temporary.name)
        self.root, self.destination, self.work = (self.base / name for name in ("data", "usb", "work"))
        (self.root / "comparison").mkdir(parents=True)
        (self.root / "raw").mkdir()
        self.old = self.root / "old.zip"
        self.old_data = b"old original bytes"
        self.old.write_bytes(archive_bytes([("old/workbook.xls", self.old_data)]))
        self.old_rows = [row(self.old_data, "old.zip", ["old/workbook.xls"], [0])]
        self.new_rows = []

    def tearDown(self):
        self.temporary.cleanup()

    def sources(self, entries):
        (self.root / "raw/new.zip").write_bytes(archive_bytes(entries))
        for name, values in (("new-only", self.new_rows), ("old-not-found", self.old_rows)):
            with (self.root / "comparison" / (name + ".csv")).open("w", encoding="utf-8", newline="") as stream:
                writer = csv.DictWriter(stream, fieldnames=list(self.old_rows[0]))
                writer.writeheader()
                writer.writerows(values)

    def run_export(self):
        return export.export_differences(self.root, self.old, self.destination, self.work,
                                         self.base / "status.json")

    def test_nested_files_are_copied_once_and_all_aliases_are_mapped(self):
        first, second = b"first original", b"second original"
        nested = archive_bytes([("book.doc", first), ("other.xls", second), ("alias.xls", first)])
        self.new_rows = [row(first, "new.zip", ["samples.zip", "book.doc"], [0, 0]),
                         row(second, "new.zip", ["samples.zip", "other.xls"], [0, 1]),
                         row(first, "new.zip", ["samples.zip", "alias.xls"], [0, 2])]
        self.sources([("samples.zip", nested)])
        with patch.object(export.tempfile, "NamedTemporaryFile", wraps=export.tempfile.NamedTemporaryFile) as staging:
            result = self.run_export()
        self.assertEqual(1, staging.call_count)
        self.assertEqual({"new-only": 2, "old-not-found": 1}, result["counts"])
        self.assertEqual(3, len(result["files"]))
        for item in result["files"]:
            self.assertEqual(item["sha256"], hashlib.sha256((self.destination / item["path"]).read_bytes()).hexdigest())
        copied = next(item["path"] for item in result["files"] if item["sha256"] == hashlib.sha256(first).hexdigest())
        self.assertTrue(copied.endswith("book.doc"))
        self.assertIn("samples.zip.contents/", copied)
        with (self.destination / "new-only/files.csv").open(encoding="utf-8-sig", newline="") as stream:
            aliases = list(csv.DictReader(stream))
        self.assertEqual(3, len(aliases))
        self.assertEqual(aliases[0]["export_path"], aliases[2]["export_path"])
        self.assertTrue((self.destination / "COPY_COMPLETE.txt").is_file())

    def test_source_hash_mismatch_does_not_publish_a_file(self):
        self.new_rows = [row(b"expected original", "new.zip", ["book.xls"], [0])]
        self.sources([("book.xls", b"different original")])
        with self.assertRaises(ValueError):
            self.run_export()
        self.assertFalse((self.destination / "COPY_COMPLETE.txt").exists())
        self.assertEqual("failed", json.loads((self.base / "status.json").read_text())["state"])
        self.assertFalse(list(self.destination.rglob("*.pending")))

    def test_output_paths_cannot_escape_the_destination(self):
        data = b"spreadsheet bytes"
        name = "../../outside.xls"
        self.new_rows = [row(data, "new.zip", [name], [0])]
        self.sources([(name, data)])
        result = self.run_export()
        for item in result["files"]:
            (self.destination / item["path"]).resolve().relative_to(self.destination.resolve())
        self.assertFalse((self.base / "outside.xls").exists())

    def test_case_collisions_keep_different_contents(self):
        self.new_rows = [row(b"first", "new.zip", ["A.xls"], [0]),
                         row(b"second", "new.zip", ["a.XLS"], [1])]
        self.sources([("A.xls", b"first"), ("a.XLS", b"second")])
        result = self.run_export()
        paths = [item["path"].lower() for item in result["files"]]
        self.assertEqual(3, len(set(paths)))

    def test_resume_reuses_verified_outputs_and_retains_changed_files(self):
        self.new_rows = [row(b"first", "new.zip", ["book.xls"], [0])]
        self.sources([("book.xls", b"first")])
        result = self.run_export()
        with patch.object(export, "_extract_tree", side_effect=AssertionError("Existing bytes were recopied")):
            self.run_export()
        path = self.destination / result["files"][0]["path"]
        path.write_bytes(b"changed locally")
        with self.assertRaises(ValueError):
            self.run_export()
        self.assertEqual(b"changed locally", path.read_bytes())
        self.assertFalse((self.destination / "COPY_COMPLETE.txt").exists())

    def test_archive_indices_are_checked_against_original_member_names(self):
        self.new_rows = [row(b"first", "new.zip", ["wrong-name.xls"], [0])]
        self.sources([("book.xls", b"first")])
        with self.assertRaisesRegex(ValueError, "index or member name"):
            self.run_export()

    def test_changed_comparison_lists_prevent_the_completion_marker(self):
        self.new_rows = [row(b"first", "new.zip", ["book.xls"], [0])]
        self.sources([("book.xls", b"first")])
        extract = export._extract_tree
        def change_list(*args):
            extract(*args)
            with (self.root / "comparison/new-only.csv").open("a", encoding="utf-8") as stream:
                stream.write("\n")
        with patch.object(export, "_extract_tree", side_effect=change_list):
            with self.assertRaisesRegex(ValueError, "lists changed"):
                self.run_export()
        self.assertFalse((self.destination / "COPY_COMPLETE.txt").exists())


if __name__ == "__main__":
    unittest.main()
