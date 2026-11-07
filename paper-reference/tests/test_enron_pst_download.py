import hashlib
import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from excel_shrink_research.enron.pst_download import resumable_download


class Response(io.BytesIO):
    def __init__(self, data, status=200, headers=None):
        super().__init__(data)
        self.status = status
        self.headers = headers or {}


class ResumablePstTests(unittest.TestCase):
    def test_verified_range_resumes_without_repeating_prefix(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'source.pst'
            payload = b'prefix-and-rest'
            target.with_suffix('.pst.partial').write_bytes(payload[:6])
            response = Response(payload[6:], 206, {'Content-Range': 'bytes 6-14/15'})
            with patch('urllib.request.urlopen', return_value=response) as request:
                resumable_download('https://example.test/source', target, len(payload), hashlib.sha256(payload).hexdigest())
            self.assertEqual('bytes=6-', request.call_args.args[0].get_header('Range'))
            self.assertEqual(payload, target.read_bytes())

    def test_server_ignoring_range_restarts_instead_of_appending(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'source.pst'
            payload = b'complete payload'
            target.with_suffix('.pst.partial').write_bytes(payload[:4])
            with patch('urllib.request.urlopen', return_value=Response(payload)):
                resumable_download('https://example.test/source', target, len(payload), hashlib.sha256(payload).hexdigest())
            self.assertEqual(payload, target.read_bytes())

    def test_wrong_hash_is_never_published_as_a_pst(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'source.pst'
            with patch('urllib.request.urlopen', side_effect=lambda *a, **k: Response(b'bad')), \
                 patch('excel_shrink_research.enron.pst_download.time.sleep'):
                with self.assertRaises(ValueError):
                    resumable_download('https://example.test/source', target, 3, '0' * 64)
            self.assertFalse(target.exists())
            self.assertTrue(list(target.parent.glob('*.bad-sha256-*')))

    def test_existing_bad_final_file_is_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'source.pst'
            target.write_bytes(b'bad')
            with patch('urllib.request.urlopen') as request:
                with self.assertRaises(ValueError):
                    resumable_download('https://example.test/source', target, 3, '0' * 64)
            self.assertEqual(b'bad', target.read_bytes())
            request.assert_not_called()


if __name__ == '__main__':
    unittest.main()
