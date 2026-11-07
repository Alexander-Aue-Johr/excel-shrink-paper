import struct
import unittest

from excel_shrink_research.enron.ole_recovery import recover


def compound_fixture():
    data = bytearray(5632)
    data[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHH', data, 26, 3, 65534, 9)
    struct.pack_into('<I', data, 48, 1)
    struct.pack_into('<109I', data, 76, 0, *([0xffffffff] * 108))
    struct.pack_into('<128I', data, 512, *([0xffffffff] * 128))
    struct.pack_into('<II', data, 512, 0xfffffffd, 0xfffffffe)
    for sector in range(2, 10):
        struct.pack_into('<I', data, 512 + sector * 4, sector + 1 if sector < 9 else 0xfffffffe)
    for index, name, kind in ((0, 'Root Entry', 5), (1, 'Workbook', 2)):
        offset = 1024 + index * 128
        encoded = (name + '\0').encode('utf-16le')
        data[offset:offset + len(encoded)] = encoded
        struct.pack_into('<H', data, offset + 64, len(encoded))
        data[offset + 66] = kind
        struct.pack_into('<III', data, offset + 68, 0xffffffff, 0xffffffff, 1 if index == 0 else 0xffffffff)
    struct.pack_into('<IQ', data, 1024 + 128 + 116, 2, 4096)
    data[1536:] = b'x' * 4096
    return data


class OleRecoveryTests(unittest.TestCase):
    def test_complete_chain_preserves_every_byte(self):
        report, stream = recover(compound_fixture())
        self.assertTrue(report['complete'])
        self.assertEqual(stream, b'x' * 4096)

    def test_truncated_sector_is_not_replaced_with_zeros(self):
        report, stream = recover(compound_fixture()[:-300])
        self.assertFalse(report['complete'])
        self.assertEqual(len(stream), 3584)
        self.assertNotIn('workbook_sha256', report)

    def test_chain_cycle_is_not_reported_as_complete(self):
        data = compound_fixture()
        struct.pack_into('<I', data, 512 + 3 * 4, 2)
        report, stream = recover(data)
        self.assertFalse(report['complete'])
        self.assertIn('repeated', report['error'])


if __name__ == '__main__':
    unittest.main()
