import csv
import hashlib
import io
import json
from pathlib import Path
import struct
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from excel_shrink_research.enron import compare, ingest
from excel_shrink_research.enron.common import sha256, write_json


WORKBOOK = b"\x09\x08\x08\x00original workbook bytes"


def archive_bytes(entries, compression=zipfile.ZIP_STORED):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=compression) as archive:
        for name, data in entries:
            archive.writestr(name, data)
    return output.getvalue()


class CompareTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.base = Path(self.temporary.name)
        self.root = self.base / "new"
        self.root.mkdir()
        self.destination = self.base / "comparison"
        self.sources = []

    def old_index(self, entries):
        archive = self.base / "old.zip"
        archive.write_bytes(archive_bytes(entries))
        return compare.index_old_archive(archive, self.destination)

    def source(self, entries, name="source.zip"):
        data = archive_bytes(entries)
        source = {"name": name, "size": len(data), "md5": hashlib.md5(data).hexdigest(),
                  "sha1": hashlib.sha1(data).hexdigest(), "url": "https://example.invalid/" + name}
        raw = self.root / "raw" / name
        raw.parent.mkdir(exist_ok=True)
        raw.write_bytes(data)
        stat = raw.stat()
        write_json(raw.parent / ".verified" / (name + ".json"),
                   {"source": ingest._fingerprint(source), "size": stat.st_size,
                    "mtime_ns": stat.st_mtime_ns, "sha256": sha256(raw)})
        self.sources.append(source)
        write_json(self.root / "sources.json", {"archives": self.sources})
        return source

    def report(self, name):
        with (self.destination / name).open(encoding="utf-8", newline="") as stream:
            return list(csv.DictReader(stream))

    def test_renames_duplicates_and_both_difference_directions(self):
        old = self.old_index([("old-name.xls", WORKBOOK), ("old-only.xls", WORKBOOK + b" old")])
        self.source([("renamed.data", WORKBOOK), ("twice.xls", WORKBOOK),
                     ("new.xls", WORKBOOK + b" new"), ("manual.bin", b"\x00unknown")])
        summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, summary["old_found_unique_sha256"])
        self.assertEqual(1, summary["old_not_found_unique_sha256"])
        self.assertEqual(1, summary["new_only_unique_sha256"])
        self.assertEqual(2, summary["counts"]["matching_new_occurrences"])
        self.assertEqual(2, len(self.report("matches.csv")))
        self.assertEqual("old-only.xls", json.loads(self.report("old-not-found.csv")[0]["member_chain"])[0])
        self.assertEqual(1, len(self.report("review-only.csv")))
        self.assertTrue(summary["new_only_definitive"])
        self.assertTrue(summary["old_not_found_definitive"])

    def test_duplicate_names_remain_distinct_occurrences_inside_nested_zip(self):
        old = self.old_index([("original.xls", WORKBOOK)])
        with self.assertWarns(UserWarning):
            nested = archive_bytes([("same.xls", WORKBOOK), ("same.xls", WORKBOOK)])
        self.source([("attachments.zip", nested)])
        summary = compare.compare_corpora(self.root, old, self.destination)
        matches = self.report("matches.csv")
        self.assertEqual(2, summary["counts"]["matching_new_occurrences"])
        self.assertEqual({"[0, 0]", "[0, 1]"}, {row["index_chain"] for row in matches})
        self.assertEqual(["attachments.zip", "same.xls"], json.loads(matches[0]["member_chain"]))
        self.assertTrue(summary["coverage_complete"])

    def test_bad_nested_member_does_not_hide_readable_siblings(self):
        for failure in ("header", "crc", "deflate"):
            with self.subTest(failure=failure):
                old = self.old_index([("original.xls", WORKBOOK), ("missing.xls", WORKBOOK + b" missing")])
                compression = zipfile.ZIP_DEFLATED if failure == "deflate" else zipfile.ZIP_STORED
                nested = bytearray(archive_bytes([("before.xls", WORKBOOK), ("broken.JPG", b"some bad bytes"),
                                                  ("after.xls", WORKBOOK + b" new")], compression))
                with zipfile.ZipFile(io.BytesIO(nested)) as archive:
                    info = archive.infolist()[1]
                    offset = info.header_offset
                    if failure == "header":
                        nested[offset:offset + 4] = b"BAD!"
                    else:
                        name_size, extra_size = struct.unpack_from("<HH", nested, offset + 26)
                        data_offset = offset + 30 + name_size + extra_size
                        if failure == "deflate":
                            nested[data_offset] |= 6  # Reserved DEFLATE block type.
                        else:
                            nested[data_offset] ^= 1
                self.sources = []
                self.source([("attachments.zip", bytes(nested))], name=failure + ".zip")
                summary = compare.compare_corpora(self.root, old, self.destination)
                self.assertEqual(1, summary["old_found_unique_sha256"])
                self.assertEqual(1, summary["new_only_unique_sha256"])
                self.assertTrue(summary["new_only_definitive"])
                self.assertFalse(summary["old_not_found_definitive"])
                issues = self.report("scan-issues.csv")
                self.assertEqual(1, len(issues))
                self.assertEqual(["attachments.zip", "broken.JPG"], json.loads(issues[0]["member_chain"]))

    def test_complete_inventory_is_reused_and_partial_inventory_is_not(self):
        old = self.old_index([("old.xls", WORKBOOK)])
        self.source([("shared.xls", WORKBOOK), ("new.xls", WORKBOOK + b" new")])
        ingest.inventory_sources(self.root, {"archives": self.sources})
        with patch.object(compare, "_scan_source", side_effect=AssertionError("Unexpected source scan")):
            summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, summary["new_only_unique_sha256"])
        marker = self.destination / "sources" / "source.zip.complete.json"
        self.assertEqual("inventory", json.loads(marker.read_text())["method"])
        inventory = self.root / "inventory" / "source.zip.jsonl"
        inventory.rename(inventory.with_suffix(".jsonl.partial"))
        with patch.object(compare, "_scan_source", wraps=compare._scan_source) as scan:
            compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, scan.call_count)

    def test_truncated_inventory_falls_back_to_the_source(self):
        old = self.old_index([("old.xls", WORKBOOK)])
        self.source([("shared.xls", WORKBOOK), ("new.xls", WORKBOOK + b" new")])
        ingest.inventory_sources(self.root, {"archives": self.sources})
        inventory = self.root / "inventory" / "source.zip.jsonl"
        inventory.write_text(inventory.read_text().splitlines()[0] + "\n")
        summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, summary["new_only_unique_sha256"])
        self.assertEqual(1, summary["old_found_unique_sha256"])

    def test_size_prefilter_finds_legacy_bytes_despite_other_classification(self):
        data = b"<?xml original bytes" + b"x" * 100000
        old = self.old_index([("legacy.xls", data)])
        self.source([("renamed.xml", data), ("unrelated.xml", data + b" other")])
        summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, summary["old_found_unique_sha256"])
        self.assertEqual(1, summary["counts"]["prefix_only"])
        self.assertEqual(1, len(self.report("matches.csv")))

    def test_corrupt_crc_metadata_cannot_establish_absence_of_other_content(self):
        data = b"<?xml original bytes" + b"x" * 100000
        old = self.old_index([("legacy.xls", data)])
        source = self.source([("renamed.xml", data)])
        raw = self.root / "raw" / source["name"]
        damaged = bytearray(raw.read_bytes())
        central = damaged.index(b"PK\x01\x02")
        struct.pack_into("<I", damaged, central + 16, struct.unpack_from("<I", damaged, central + 16)[0] ^ 1)
        raw.write_bytes(damaged)
        source.update(md5=hashlib.md5(damaged).hexdigest(), sha1=hashlib.sha1(damaged).hexdigest())
        write_json(self.root / "sources.json", {"archives": self.sources})
        stat = raw.stat()
        write_json(raw.parent / ".verified" / (source["name"] + ".json"),
                   {"source": ingest._fingerprint(source), "size": stat.st_size,
                    "mtime_ns": stat.st_mtime_ns, "sha256": sha256(raw)})
        summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertFalse(summary["old_not_found_definitive"])
        self.assertFalse(summary["coverage_complete"])
        self.assertEqual(0, summary["counts"].get("prefix_only", 0))
        self.assertEqual("unreadable_member", self.report("scan-issues.csv")[0]["reason"])
        self.assertEqual("not_found_in_scanned_archives", self.report("old-not-found.csv")[0]["status"])

    def test_missing_hash_in_completed_inventory_cannot_prove_old_absence(self):
        old = self.old_index([("old.xls", WORKBOOK)])
        self.source([("shared.xls", WORKBOOK)])
        ingest.inventory_sources(self.root, {"archives": self.sources})
        inventory = self.root / "inventory" / "source.zip.jsonl"
        row = json.loads(inventory.read_text().splitlines()[0])
        row["sha256"] = None
        inventory.write_text(json.dumps(row) + "\n")
        summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, summary["old_found_unique_sha256"])
        self.assertEqual(0, summary["old_not_found_unique_sha256"])

    def test_cache_resumes_and_changes_invalidate_only_the_affected_source(self):
        old = self.old_index([("old.xls", WORKBOOK)])
        self.source([("same.xls", WORKBOOK)], "first.zip")
        self.source([("new.xls", WORKBOOK + b" new")], "second.zip")
        compare.compare_corpora(self.root, old, self.destination)
        with patch.object(compare, "_scan_source", side_effect=AssertionError("A completed source was rescanned")):
            compare.compare_corpora(self.root, old, self.destination)
        raw = self.root / "raw" / "second.zip"
        raw.touch()
        with patch.object(compare, "_scan_source", wraps=compare._scan_source) as scan:
            summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, scan.call_count)
        self.assertFalse(summary["coverage_complete"])
        self.assertIn("source_not_verified", {row["reason"] for row in self.report("scan-issues.csv")})

    def test_unreadable_old_content_blocks_definitive_new_only_claims(self):
        old = self.old_index([("old.xls", WORKBOOK)])
        row = json.loads(old.read_text().splitlines()[0])
        row.update(sha256=None, status="unreadable")
        old.write_text(json.dumps(row) + "\n")
        self.source([("new.xls", WORKBOOK + b" new")])
        summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertEqual(1, summary["new_only_unique_sha256"])
        self.assertFalse(summary["new_only_definitive"])
        self.assertFalse(summary["old_not_found_definitive"])
        self.assertEqual(1, summary["old_unreadable_entries"])
        self.assertEqual("old_file_unreadable", self.report("old-not-found.csv")[0]["status"])

    def test_depth_limit_reports_a_gap_and_preserves_definitive_positive_difference(self):
        old = self.old_index([("old.xls", WORKBOOK)])
        self.source([("nested.zip", archive_bytes([("hidden.xls", WORKBOOK)])),
                     ("new.xls", WORKBOOK + b" new")])
        with patch.object(ingest, "MAX_NESTING", 0):
            summary = compare.compare_corpora(self.root, old, self.destination)
        self.assertFalse(summary["old_not_found_definitive"])
        self.assertTrue(summary["new_only_definitive"])
        self.assertEqual(1, summary["new_only_unique_sha256"])
        self.assertEqual(1, len(self.report("scan-issues.csv")))
        self.assertEqual("not_found_in_scanned_archives", self.report("old-not-found.csv")[0]["status"])

    def test_old_index_reuses_a_verified_cache_and_detects_source_changes(self):
        old = self.old_index([("old.xls", WORKBOOK), ("readme.txt", b"Some notes")])
        with patch.object(compare.zipfile, "ZipFile", side_effect=AssertionError("Cached source was reopened")):
            self.assertEqual(old, compare.index_old_archive(self.base / "old.zip", self.destination))
        (self.base / "old.zip").write_bytes(archive_bytes([("changed.xls", WORKBOOK + b" changed")]))
        compare.index_old_archive(self.base / "old.zip", self.destination)
        indexed = [json.loads(line) for line in old.read_text().splitlines()]
        self.assertEqual("changed.xls", indexed[0]["member"])


if __name__ == "__main__":
    unittest.main()
