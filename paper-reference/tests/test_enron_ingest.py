import hashlib
import io
import json
from pathlib import Path
import shutil
import struct
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from excel_shrink_research.enron import ingest


class Response(io.BytesIO):
    def __init__(self, data, status=200, headers=None):
        super().__init__(data)
        self.status = status
        self.headers = headers or {}


def source(data, name="source.zip"):
    return {
        "name": name, "size": len(data), "md5": hashlib.md5(data).hexdigest(),
        "sha1": hashlib.sha1(data).hexdigest(), "url": "https://example.invalid/" + name,
    }


def archive_bytes(entries):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_STORED) as archive:
        for name, data in entries:
            archive.writestr(name, data)
    return output.getvalue()


def ole_bytes(stream_name):
    header = bytearray(512)
    header[:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<H", header, 30, 9)
    struct.pack_into("<I", header, 48, 0)
    struct.pack_into("<I", header, 68, 0xFFFFFFFE)
    struct.pack_into("<109I", header, 76, 1, *([0xFFFFFFFF] * 108))
    directory = bytearray(512)
    encoded = (stream_name + "\0").encode("utf-16-le")
    directory[:len(encoded)] = encoded
    struct.pack_into("<H", directory, 64, len(encoded))
    directory[66] = 2
    fat = struct.pack("<128I", 0xFFFFFFFE, 0xFFFFFFFD, *([0xFFFFFFFF] * 126))
    return bytes(header + directory) + fat


class IngestTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self):
        shutil.rmtree(ingest._long_path(self.root))
        self.temporary.cleanup()

    def inventory(self, entries):
        data = archive_bytes(entries)
        raw = self.root / "raw" / "source.zip"
        raw.parent.mkdir(exist_ok=True)
        raw.write_bytes(data)
        manifest = {"archives": [source(data)]}
        path = ingest.inventory_sources(self.root, manifest)
        return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()], manifest

    def test_freeze_pins_metadata_and_reuses_it(self):
        files = [{**source(b"zip", "source-%03d.zip" % index), "size": "3"} for index in range(159)]
        metadata = json.dumps({"files": files}).encode("utf-8")
        with patch.object(ingest, "_request", return_value=Response(metadata)) as request:
            first = ingest.freeze_sources(self.root)
            second = ingest.freeze_sources(self.root)
        self.assertEqual(1, request.call_count)
        self.assertEqual(first, second)
        self.assertEqual(hashlib.sha256((self.root / "source-metadata.json").read_bytes()).hexdigest(), first["metadata_sha256"])
        self.assertEqual(159, len(first["archives"]))

    def test_freeze_rejects_an_incomplete_source_set(self):
        with patch.object(ingest, "_request", return_value=Response(b'{"files": []}')):
            with self.assertRaisesRegex(ingest.SourceError, "Expected 159"):
                ingest.freeze_sources(self.root)
        self.assertFalse((self.root / "sources.json").exists())

    def test_download_resumes_at_the_partial_length_and_verifies_hashes(self):
        data = b"0123456789"
        destination = self.root / "source.zip"
        destination.with_suffix(".zip.partial").write_bytes(data[:4])
        response = Response(data[4:], 206, {"Content-Range": "bytes 4-9/10"})
        with patch.object(ingest, "_request", return_value=response) as request:
            hashes = ingest._download(source(data), destination, None)
        self.assertEqual({"Range": "bytes=4-"}, request.call_args.args[1])
        self.assertEqual(data, destination.read_bytes())
        self.assertEqual(hashlib.sha256(data).hexdigest(), hashes["sha256"])

    def test_download_restarts_when_server_ignores_range(self):
        data = b"0123456789"
        destination = self.root / "source.zip"
        destination.with_suffix(".zip.partial").write_bytes(data[:4])
        with patch.object(ingest, "_request", return_value=Response(data)):
            ingest._download(source(data), destination, None)
        self.assertEqual(data, destination.read_bytes())

    def test_checksum_failure_never_becomes_a_completed_archive(self):
        destination = self.root / "source.zip"
        with patch.object(ingest, "_request", side_effect=lambda *args: Response(b"bad")), patch.object(ingest.time, "sleep"):
            with self.assertRaisesRegex(ingest.SourceError, "checksum mismatch"):
                ingest._download(source(b"zip"), destination, None, retries=2)
        self.assertFalse(destination.exists())
        self.assertFalse(destination.with_suffix(".zip.partial").exists())

    def test_incorrect_range_is_rejected_without_appending(self):
        destination = self.root / "source.zip"
        partial = destination.with_suffix(".zip.partial")
        partial.write_bytes(b"0123")
        with patch.object(ingest, "_request", return_value=Response(b"789", 206, {"Content-Range": "bytes 7-9/10"})):
            with self.assertRaisesRegex(ingest.SourceError, "unexpected byte range"):
                ingest._download(source(b"0123456789"), destination, None, retries=1)
        self.assertEqual(b"0123", partial.read_bytes())

    def test_extraction_preserves_occurrences_and_handles_windows_paths(self):
        workbook = b"\x09\x08\x08\x00example"
        rows, _ = self.inventory([
            ("../outside.xls", workbook), ("folder/Book.xls", workbook),
            ("folder/book.xls", workbook), ("CON.xls", workbook),
            ("folder/percent%.xls", workbook), ("C:\\escape.xls", workbook),
            ("report.pdf", b"%PDF-1.4 unrelated"),
        ])
        candidates = [row for row in rows if row["classification"] == "candidate"]
        paths = [row["extracted_path"] for row in candidates]
        self.assertEqual(6, len(candidates))
        self.assertEqual(6, len({path.casefold() for path in paths}))
        self.assertTrue(any("%2E%2E" in path for path in paths))
        self.assertTrue(any("%43ON.xls" in path for path in paths))
        self.assertTrue(any("percent%25.xls" in path for path in paths))
        self.assertTrue(any("%3A" in path for path in paths))
        self.assertEqual("C%3A%5Cescape.xls", ingest._escape_component("C:\\escape.xls"))
        self.assertEqual(5, sum("duplicate_of" in row for row in candidates))
        self.assertEqual("other", rows[-1]["classification"])
        self.assertIsNone(rows[-1]["extracted_path"])
        for row in candidates:
            self.assertTrue(row["attempt_excel"])
            self.assertEqual(workbook, (self.root / row["extracted_path"]).read_bytes())

    def test_content_detection_finds_workbooks_with_unrelated_suffixes(self):
        xlsx = archive_bytes([("[Content_Types].xml", b"types"), ("xl/workbook.xml", b"<workbook/>")])
        rows, _ = self.inventory([
            ("attachment.bin", ole_bytes("Workbook")),
            ("attachment.dat", xlsx),
            ("encrypted.dat", ole_bytes("EncryptedPackage")),
            ("word.xls", ole_bytes("WordDocument")),
            ("plain.xls", b"one;two\n1;2\n"),
            ("xml.dat", b'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/>'),
        ])
        self.assertEqual(["xls", "xlsx", "encrypted-office", "other-office", "text", "spreadsheetml"], [row["kind"] for row in rows])
        self.assertTrue(rows[2]["attempt_excel"])
        self.assertEqual("ambiguous", rows[2]["classification"])
        self.assertIsNone(rows[3]["extracted_path"])
        self.assertTrue(rows[4]["attempt_excel"])

    def test_nested_zip_members_keep_the_source_chain(self):
        nested = archive_bytes([("inner/book.data", b"ID;P\nC;X1;Y1;K1\n")])
        rows, _ = self.inventory([("attachments.zip", nested)])
        self.assertEqual("scanned", rows[0]["nested_status"])
        self.assertEqual(["attachments.zip", "inner/book.data"], rows[1]["member_chain"])
        self.assertEqual([0, 0], rows[1]["index_chain"])
        self.assertIn("attachments.zip/inner/book.data", rows[1]["extracted_path"])
        self.assertTrue(rows[1]["attempt_excel"])

    def test_nested_zip_limit_retains_bytes_for_review(self):
        nested = archive_bytes([("book.xls", b"\x09\x08\x08\x00example")])
        with patch.object(ingest, "MAX_NESTING", 0):
            rows, _ = self.inventory([("attachments.zip", nested)])
        self.assertEqual(1, len(rows))
        self.assertEqual("ambiguous", rows[0]["classification"])
        self.assertFalse(rows[0]["attempt_excel"])
        self.assertEqual(nested, (self.root / rows[0]["extracted_path"]).read_bytes())

    def test_encrypted_nested_zip_is_retained_without_blocking_other_workbooks(self):
        nested = bytearray(archive_bytes([("secret.xls", b"encrypted bytes")]))
        local_header = nested.index(b"PK\x03\x04")
        central_header = nested.index(b"PK\x01\x02")
        struct.pack_into("<H", nested, local_header + 6, 1)
        struct.pack_into("<H", nested, central_header + 8, 1)
        rows, _ = self.inventory([
            ("encrypted.zip", nested), ("normal.xls", b"\x09\x08\x08\x00example"),
        ])
        self.assertEqual("encrypted-archive", rows[0]["kind"])
        self.assertEqual("ambiguous", rows[0]["classification"])
        self.assertFalse(rows[0]["attempt_excel"])
        self.assertEqual(bytes(nested), (self.root / rows[0]["extracted_path"]).read_bytes())
        self.assertTrue(rows[1]["attempt_excel"])

    def test_bad_nested_member_is_recorded_and_readable_siblings_continue(self):
        workbook = b"\x09\x08\x08\x00example"
        nested = bytearray(archive_bytes([("photo.jpg", b"photo"), ("inner/book.xls", workbook)]))
        central = nested.index(b"PK\x01\x02")
        struct.pack_into("<I", nested, central + 42, 1)  # Point the first member away from its local header.
        rows, manifest = self.inventory([("attachments.zip", nested), ("normal.xls", workbook)])
        container = rows[0]
        self.assertEqual("partial", container["nested_status"])
        self.assertEqual(1, container["nested_read_errors"])
        self.assertFalse(container["attempt_excel"])
        self.assertEqual(bytes(nested), (self.root / container["extracted_path"]).read_bytes())
        unreadable = rows[1]
        self.assertEqual("unreadable", unreadable["classification"])
        self.assertEqual(["attachments.zip", "photo.jpg"], unreadable["member_chain"])
        self.assertEqual([0, 0], unreadable["index_chain"])
        self.assertIsNone(unreadable["sha256"])
        self.assertFalse(unreadable["attempt_excel"])
        self.assertEqual("BadZipFile", unreadable["read_error"]["type"])
        self.assertEqual(2, sum(row["attempt_excel"] for row in rows))
        marker = json.loads((self.root / "inventory/source.zip.complete.json").read_text())
        self.assertEqual("completed_with_read_errors", marker["scan_status"])
        self.assertEqual(1, marker["counts"]["read_errors"])
        with patch.object(ingest, "_inventory_archive", side_effect=AssertionError("Unexpected rescan")):
            ingest.inventory_sources(self.root, manifest)

    def test_nested_crc_failure_retains_original_bytes_and_continues(self):
        workbook = b"\x09\x08\x08\x00example"
        nested = bytearray(archive_bytes([("bad.xls", workbook), ("good.xls", workbook)]))
        nested[nested.index(b"example")] = ord("X")
        rows, _ = self.inventory([("attachments.zip", nested)])
        self.assertEqual("partial", rows[0]["nested_status"])
        self.assertEqual(bytes(nested), (self.root / rows[0]["extracted_path"]).read_bytes())
        self.assertEqual("unreadable", rows[1]["classification"])
        self.assertIn("CRC", rows[1]["read_error"]["message"])
        self.assertTrue(rows[2]["attempt_excel"])
        self.assertEqual(workbook, (self.root / rows[2]["extracted_path"]).read_bytes())

    def test_nested_output_failure_still_stops_inventory_publication(self):
        nested = archive_bytes([("book.xls", b"\x09\x08\x08\x00example")])
        data = archive_bytes([("attachments.zip", nested)])
        raw = self.root / "raw" / "source.zip"
        raw.parent.mkdir()
        raw.write_bytes(data)
        with patch.object(Path, "replace", side_effect=OSError("Disk full")):
            with self.assertRaisesRegex(ingest.SourceError, "Disk full"):
                ingest.inventory_sources(self.root, {"archives": [source(data)]})
        self.assertFalse((self.root / "inventory.jsonl").exists())
        self.assertFalse((self.root / "inventory/source.zip.complete.json").exists())

    def test_bad_nested_deflate_stream_is_recorded(self):
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as nested:
            nested.writestr("bad.xls", b"\x09\x08\x08\x00example" * 20)
        damaged = bytearray(output.getvalue())
        damaged[30 + len("bad.xls")] = 7  # Invalid DEFLATE block type.
        rows, _ = self.inventory([("attachments.zip", damaged)])
        self.assertEqual("partial", rows[0]["nested_status"])
        self.assertEqual("unreadable", rows[1]["classification"])
        self.assertIn("invalid block type", rows[1]["read_error"]["message"])
        self.assertEqual(bytes(damaged), (self.root / rows[0]["extracted_path"]).read_bytes())

    @unittest.skipUnless(ingest.os.name == "nt", "Windows path behavior")
    def test_deep_source_tree_works_without_windows_long_path_policy(self):
        name = "/".join(["long-directory-name" * 3] * 6) + "/book.xls"
        rows, _ = self.inventory([(name, b"\x09\x08\x08\x00example")])
        path = self.root / rows[0]["extracted_path"]
        self.assertGreater(len(str(path)), 260)
        self.assertTrue(ingest._long_path(path).is_file())

    def test_zip_crc_failure_stops_inventory_publication(self):
        data = bytearray(archive_bytes([("book.xls", b"\x09\x08\x08\x00example")]))
        payload = data.index(b"example")
        data[payload] = ord("X")
        raw = self.root / "raw" / "source.zip"
        raw.parent.mkdir()
        raw.write_bytes(data)
        with self.assertRaisesRegex(ingest.SourceError, "CRC"):
            ingest.inventory_sources(self.root, {"archives": [source(data)]})
        self.assertFalse((self.root / "inventory.jsonl").exists())
        self.assertFalse((self.root / "inventory" / "source.zip.complete.json").exists())

    def test_checkpoint_is_reused_but_missing_outputs_are_rebuilt(self):
        rows, manifest = self.inventory([("book.xls", b"\x09\x08\x08\x00example")])
        with patch.object(ingest, "_inventory_archive", side_effect=AssertionError("Unexpected rescan")):
            ingest.inventory_sources(self.root, manifest)
        extracted = self.root / rows[0]["extracted_path"]
        extracted.unlink()
        ingest.inventory_sources(self.root, manifest)
        self.assertTrue(extracted.exists())


if __name__ == "__main__":
    unittest.main()
