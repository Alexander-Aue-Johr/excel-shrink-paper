import struct
import unittest
import zlib

from excel_shrink_research.enron.container_embeddings import ppt_objects, native_payload


class PowerPointObjectTests(unittest.TestCase):
    def test_native_package_preserves_payload(self):
        payload = b'unchanged file bytes'
        wrapper = struct.pack('<IH', 0, 2) + b'file.xls\0C:\\source\0' + b'\0' * 8 + b'C:\\temp\0'
        wrapper += struct.pack('<I', len(payload))
        self.assertEqual(payload, native_payload(wrapper + payload))
        with self.assertRaises(ValueError):
            native_payload(wrapper + payload[:-1])

    def test_compressed_object_in_container_keeps_exact_bytes(self):
        payload = b'embedded object bytes' * 100
        body = struct.pack('<I', len(payload)) + zlib.compress(payload)
        record = struct.pack('<HHI', 0x10, 0x1011, len(body)) + body
        container = struct.pack('<HHI', 0xf, 0x3e8, len(record)) + record
        self.assertEqual([(8, payload)], list(ppt_objects(container)))

    def test_wrong_decompressed_size_is_rejected(self):
        body = struct.pack('<I', 1) + zlib.compress(b'too many bytes')
        record = struct.pack('<HHI', 0x10, 0x1011, len(body)) + body
        with self.assertRaises(ValueError):
            list(ppt_objects(record))

    def test_truncated_record_is_rejected(self):
        with self.assertRaises(ValueError):
            list(ppt_objects(struct.pack('<HHI', 0, 0x1011, 100) + b'short'))

    def test_uncompressed_object_keeps_exact_bytes(self):
        payload = b'unchanged bytes'
        record = struct.pack('<HHI', 0, 0x1011, len(payload)) + payload
        self.assertEqual([(0, payload)], list(ppt_objects(record)))


if __name__ == '__main__':
    unittest.main()
