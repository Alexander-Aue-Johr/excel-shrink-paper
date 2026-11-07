import hashlib
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from excel_shrink_research.enron.pst_probe import download


class VerifiedDownloadTests(unittest.TestCase):
    def test_same_size_cache_is_not_trusted_without_matching_hash(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target = root / 'source', root / 'target'
            source.write_bytes(b'valid')
            target.write_bytes(b'wrong')
            expected = hashlib.sha256(source.read_bytes()).hexdigest()
            download(source.as_uri(), target, expected, 5)
            self.assertEqual(b'valid', target.read_bytes())

    def test_wrong_source_hash_never_becomes_verified_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, target = root / 'source', root / 'target'
            source.write_bytes(b'wrong')
            with patch('excel_shrink_research.enron.pst_probe.time.sleep'):
                with self.assertRaises(ValueError):
                    download(source.as_uri(), target, '0' * 64, 5)
            self.assertFalse(target.exists())


if __name__ == '__main__':
    unittest.main()
