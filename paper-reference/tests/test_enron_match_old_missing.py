import argparse
import contextlib
import csv
import io
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from excel_shrink_research.enron import embedded_workbooks as e
from excel_shrink_research.enron.match_old_missing import match


class MissingOriginalTests(unittest.TestCase):
    def test_same_workbook_with_changed_other_stream_is_not_full_storage_match(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for folder in ('comparison', 'embedded-workbooks-v2/recovery', 'container-embeddings'):
                (root / folder).mkdir(parents=True, exist_ok=True)
            def workbook(path, other):
                storage = e.pythoncom.StgCreateDocfile(str(path), e.CREATE, 0)
                for name, payload in [('Workbook', b'workbook data'), ('Other', other)]:
                    stream = storage.CreateStream(name, e.CREATE, 0, 0)
                    stream.Write(payload)
                    stream = None
                storage.Commit(0)
                storage = None
            old = root / 'old.xls'
            candidate = root / 'container-embeddings/extracted.xls'
            workbook(old, b'old metadata')
            workbook(candidate, b'changed metadata')
            reader = e.open_storage(candidate)
            try:
                book_hash, tree_hash = e.workbook_hash(reader), e.fingerprint(reader)
            finally:
                reader.close()
            with zipfile.ZipFile(root / 'old.zip', 'w') as archive:
                archive.writestr('old.xls', old.read_bytes())
                archive.writestr('plain.xls', b'plain text')
            with (root / 'comparison/old-not-found.csv').open('w', newline='') as stream:
                writer = csv.DictWriter(stream, fieldnames=['sha256', 'kind', 'member_chain', 'index_chain'])
                writer.writeheader()
                for index, name, data, kind in [(0, 'old.xls', old.read_bytes(), 'ole'), (1, 'plain.xls', b'plain text', 'text')]:
                    writer.writerow(dict(sha256=e.digest(data), kind=kind, member_chain=json.dumps([name]), index_chain=json.dumps([index])))
            for folder in ('embedded-workbooks-v2', 'container-embeddings'):
                rows = [] if folder == 'embedded-workbooks-v2' else [dict(source_sha256='source', object_path=['object'],
                        path='extracted.xls', file_sha256=e.digest(candidate.read_bytes()), workbook_sha256=book_hash, storage_sha256=tree_hash)]
                e.write_json(root / folder / 'comparison.json', rows)
                e.write_json(root / folder / 'source-documents.json', {})
            e.write_json(root / 'embedded-workbooks-v2/recovery/report.json', [])
            with contextlib.redirect_stdout(io.StringIO()):
                match(argparse.Namespace(root=str(root), old_zip=str(root / 'old.zip'), output=str(root / 'result')))
            findings = json.loads((root / 'result/matches.json').read_text())
            self.assertEqual(['identical_workbook_stream'], [r['evidence'] for r in findings])
            rows = json.loads((root / 'result/results.json').read_text())
            self.assertEqual(['match_found', 'no_file_match_no_workbook_stream'], [r['comparison'] for r in rows])
            self.assertNotEqual(e.digest(old.read_bytes()), e.digest(candidate.read_bytes()))


if __name__ == '__main__':
    unittest.main()
