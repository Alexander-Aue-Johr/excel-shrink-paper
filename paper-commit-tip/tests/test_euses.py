import hashlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from excel_shrink_research import euses
from excel_shrink_research.enron import ingest


def zip_bytes(entries):
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_STORED) as archive:
        for name, content in entries:
            archive.writestr(name, content)
    return output.getvalue()


class Response(io.BytesIO):
    def __init__(self, data, status=200, headers=None):
        super().__init__(data)
        self.status = status
        self.headers = headers or {}


class EusesTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.blobs = {name: zip_bytes([(name[:-4] + '/processed/book.xls', b'\x09\x08\x04\x00 spreadsheet ' + name.encode())])
                      for name in euses.ARCHIVE_NAMES}
        self.blobs['README.txt'] = b'Original release description'
        self.write_metadata()

    def tearDown(self):
        self.directory.cleanup()

    def write_metadata(self):
        data = {'id': euses.RECORD_ID, 'doi': euses.DOI, 'files': [
            {'key': name, 'size': len(value), 'checksum': 'md5:' + hashlib.md5(value).hexdigest()}
            for name, value in self.blobs.items()]}
        (self.root / 'source-metadata.json').write_text(json.dumps(data), encoding='utf-8')

    def install_raw(self):
        (self.root / 'raw').mkdir(exist_ok=True)
        for name, value in self.blobs.items():
            (self.root / 'raw' / name).write_bytes(value)
        euses.freeze_sources(self.root)

    def request(self, url, headers=None):
        name = url.split('/')[-2]
        data = self.blobs[name]
        if headers and headers.get('Range'):
            offset = int(headers['Range'].split('=')[1].split('-')[0])
            return Response(data[offset:], 206, {'Content-Range': 'bytes %d-%d/%d' % (offset, len(data) - 1, len(data))})
        return Response(data)

    def test_metadata_pins_original_release_and_rejects_changed_source_set(self):
        result = euses.freeze_sources(self.root)
        self.assertEqual(euses.DOI, result['doi'])
        self.assertEqual(12, len(result['files']))
        self.assertEqual(11, len([row for row in result['files'] if row['name'].endswith('.zip')]))
        data = json.loads((self.root / 'source-metadata.json').read_text())
        data['id'] = 123
        with self.assertRaisesRegex(ValueError, 'original EUSES'):
            euses._manifest_from_metadata(json.dumps(data).encode())
        data['id'] = euses.RECORD_ID
        data['files'][0]['checksum'] = 'md5:' + 'z' * 32
        with self.assertRaisesRegex(ValueError, 'checksum'):
            euses._manifest_from_metadata(json.dumps(data).encode())

    def test_download_verifies_each_source_and_reuses_unchanged_receipts(self):
        with patch.object(ingest, '_request', side_effect=self.request) as request:
            euses.download_euses(self.root)
            self.assertEqual(12, request.call_count)
        with patch.object(ingest, '_request', side_effect=AssertionError('An unchanged file was downloaded again')):
            euses.download_euses(self.root)
        status = json.loads((self.root / 'download-status.json').read_text())
        self.assertEqual('complete', status['state'])
        receipt = json.loads((self.root / 'raw/.verified/database.zip.json').read_text())
        self.assertEqual(hashlib.sha256(self.blobs['database.zip']).hexdigest(), receipt['sha256'])

    def test_interrupted_download_resumes_from_the_recorded_partial_bytes(self):
        (self.root / 'raw').mkdir()
        (self.root / 'raw/database.zip.partial').write_bytes(self.blobs['database.zip'][:13])
        seen = []
        def request(url, headers=None):
            if 'database.zip/' in url:
                seen.append(headers)
            return self.request(url, headers)
        with patch.object(ingest, '_request', side_effect=request):
            euses.download_euses(self.root)
        self.assertEqual({'Range': 'bytes=13-'}, seen[0])
        self.assertEqual(self.blobs['database.zip'], (self.root / 'raw/database.zip').read_bytes())

    def test_changed_download_is_not_silently_accepted(self):
        self.install_raw()
        (self.root / 'raw/database.zip').write_bytes(b'wrong contents')
        with self.assertRaisesRegex(RuntimeError, 'download size'):
            euses.download_euses(self.root)
        self.assertEqual('failed', json.loads((self.root / 'download-status.json').read_text())['state'])

    def test_archive_preserves_all_subsets_bytes_and_nonstandard_names(self):
        self.blobs['database.zip'] = zip_bytes([
            ('database/processed/book.xls.1', b'first workbook'),
            ('database/bad/error_page', b'bad original'),
            ('database/duplicates/book.XLS', b'first workbook'),
            ('database/processed/Case.xls', b'case one'),
            ('database/processed/case.xls', b'case two'),
            ('../escape.xls', b'archived bytes'),
        ])
        self.write_metadata()
        self.install_raw()
        result = euses.prepare_runner_archive(self.root)
        mapping = [json.loads(line) for line in (self.root / 'input-manifest.jsonl').read_text().splitlines()]
        self.assertEqual(15, result['workbooks'])
        self.assertEqual(16, result['total_files'])
        self.assertEqual({'processed': 13, 'bad': 1, 'duplicates': 1, 'other': 1, 'renamed_for_runner': 2,
                          'unique_inputs': 15, 'duplicate_occurrences': 1}, result['counts'])
        self.assertEqual(15, len({row['runner_member'].lower() for row in mapping}))
        with zipfile.ZipFile(result['path']) as archive:
            for row in mapping:
                self.assertTrue(row['runner_member'].lower().endswith(('.xls', '.xlsx')))
                self.assertNotIn('..', row['runner_member'].split('/'))
                content = archive.read(row['runner_member'])
                self.assertEqual(row['sha256'], hashlib.sha256(content).hexdigest())
        legacy = next(row for row in mapping if row['source_member'].endswith('book.xls.1'))
        self.assertTrue(legacy['runner_member'].endswith('book.xls.1.xls'))
        duplicate = next(row for row in mapping if row['subset'] == 'duplicates')
        self.assertEqual(legacy['id'], duplicate['duplicate_of'])
        self.assertEqual(legacy['runner_member'], duplicate['runner_member'])
        self.assertEqual(legacy['runner_identity'], duplicate['runner_identity'])
        self.assertEqual(2, len({row['runner_member'] for row in mapping if row['source_member'].lower().endswith('case.xls')}))
        self.assertEqual(euses.CORPUS_ID, result['corpus_id'])

    def test_completed_archive_is_reused_and_a_tampered_source_is_rejected(self):
        self.install_raw()
        first = euses.prepare_runner_archive(self.root)
        with patch.object(euses.zipfile, 'ZipFile', side_effect=AssertionError('A valid archive was rebuilt')):
            second = euses.prepare_runner_archive(self.root)
        self.assertEqual(first['sha256'], second['sha256'])
        (self.root / 'raw/database.zip').write_bytes(b'bad bytes')
        with self.assertRaisesRegex(RuntimeError, 'download size'):
            euses.prepare_runner_archive(self.root)
        self.assertEqual('failed', json.loads((self.root / 'archive-status.json').read_text())['state'])

    def test_unreadable_source_member_does_not_publish_a_partial_archive(self):
        data = bytearray(zip_bytes([('database/processed/book.xls', b'original bytes')]))
        offset = data.index(b'original bytes')
        data[offset] ^= 1
        self.blobs['database.zip'] = bytes(data)
        self.write_metadata()
        self.install_raw()
        with self.assertRaises(zipfile.BadZipFile):
            euses.prepare_runner_archive(self.root)
        self.assertFalse((self.root / 'runner-archive.json').exists())
        self.assertEqual('failed', json.loads((self.root / 'archive-status.json').read_text())['state'])


if __name__ == '__main__':
    unittest.main()
