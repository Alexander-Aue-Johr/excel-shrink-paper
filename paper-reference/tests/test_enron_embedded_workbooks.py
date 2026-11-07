import tempfile
import unittest
import struct
from pathlib import Path
from excel_shrink_research.enron import embedded_workbooks as e


class EmbeddedWorkbookTests(unittest.TestCase):
    def test_invalid_directory_dates_do_not_change_streams(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'invalid-dates.xls'
            root = e.pythoncom.StgCreateDocfile(str(path), e.CREATE, 0)
            stream = root.CreateStream('Workbook', e.CREATE, 0, 0)
            stream.Write(b'workbook bytes')
            stream = None
            root.Commit(0)
            root = None
            data = bytearray(path.read_bytes())
            sector_size = 1 << struct.unpack_from('<H', data, 30)[0]
            start = (struct.unpack_from('<I', data, 48)[0] + 1) * sector_size
            for index in (0, 1):
                data[start + index * 128 + 100:start + index * 128 + 116] = b'\xff' * 16
            path.write_bytes(data)
            before = path.read_bytes()
            reader = e.open_storage(path)
            try:
                self.assertEqual(e.workbook_hash(reader), e.digest(b'workbook bytes'))
                expected = e.fingerprint(reader)
                self.assertEqual(e.extract(reader, path.with_name('extracted.xls')), expected)
            finally:
                reader.close()
            self.assertEqual(path.read_bytes(), before)

    def test_extract_preserves_streams_and_distinguishes_other_content(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)
            root = e.pythoncom.StgCreateDocfile(str(path / 'word.doc'), e.CREATE, 0)
            pool = root.CreateStorage('ObjectPool', e.CREATE, 0, 0)
            book = pool.CreateStorage('_123', e.CREATE, 0, 0)
            stream = book.CreateStream('Workbook', e.CREATE, 0, 0)
            stream.Write(b'unchanged workbook data')
            stream = None
            stream = book.CreateStream('Other', e.CREATE, 0, 0)
            stream.Write(b'additional content')
            stream = None
            before = e.workbook_hash(book)
            original_tree = e.fingerprint(book)
            book = None
            pool = None
            root.Commit(0)
            root = None
            root = e.pythoncom.StgOpenStorage(str(path / 'word.doc'), None, e.READ)
            paths = []
            for object_path, child in e.embedded(root):
                paths.append(object_path)
                self.assertEqual(e.extract(child, path / 'book.xls'), original_tree)
            child = None
            self.assertEqual(paths, [('ObjectPool', '_123')])
            copied = e.pythoncom.StgOpenStorage(str(path / 'book.xls'), None, e.READ)
            self.assertEqual(e.workbook_hash(copied), before)
            copied = None
            changed = e.pythoncom.StgOpenStorage(str(path / 'book.xls'), None, 0x12)
            stream = changed.OpenStream('Other', None, 0x12, 0)
            stream.Write(b'changed content')
            stream = None
            self.assertEqual(e.workbook_hash(changed), before)
            self.assertNotEqual(e.fingerprint(changed), original_tree)
            changed = None
            root = None


if __name__ == '__main__':
    unittest.main()



