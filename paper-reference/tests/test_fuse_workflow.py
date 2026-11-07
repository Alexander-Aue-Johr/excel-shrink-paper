import tempfile
from pathlib import Path
import unittest
import zipfile

from excel_shrink_research.fuse_workflow import classify, safe_member, sheet_reductions, MAXIMUM, nested_archives
from excel_shrink_research.enron.excel import _create_seed


class FuseWorkflowTests(unittest.TestCase):
    def test_split_archive_is_extracted_once_from_first_volume(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in ('FUSE.7z.001', 'FUSE.7z.002', 'FUSE.7z.141', 'other.7z', 'notes.txt'):
                (root / name).touch()
            self.assertEqual(['FUSE.7z.001', 'other.7z'], [path.name for path in nested_archives(root)])

    def test_content_identifies_xlsx_without_extension(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ('a' * 40)
            _create_seed(path)
            self.assertEqual('native_xlsx', classify(path))

    def test_archive_path_cannot_leave_destination(self):
        for path in ('../file', 'nested/../../file', 'C:\\file', '/root', '\\server\\share'):
            with self.subTest(path=path), self.assertRaises(ValueError):
                safe_member(path)
        safe_member('corpus/a/b.xls')

    def test_xml_ratio_is_uncompressed_and_empty_sheet_is_not_a_needle(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            before, after = root / 'before.xlsx', root / 'after.xlsx'
            for path, xml in [(before, b'x' * 100000), (after, b'<worksheet><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>')]:
                with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr('xl/worksheets/sheet1.xml', xml)
            row = sheet_reductions(before, after, 1)[0]
            self.assertEqual(1, row['meaningful_cells'])
            self.assertAlmostEqual(100 * row['shrunk_xml_bytes'] / 100000, row['remaining_xml_pct'])
            with zipfile.ZipFile(after, 'w') as archive:
                archive.writestr('xl/worksheets/sheet1.xml', '<worksheet><sheetData/></worksheet>')
            self.assertEqual(0, sheet_reductions(before, after, 1)[0]['meaningful_cells'])

    def test_maximum_excludes_formula_references_only_mode(self):
        self.assertIn('--remove-unreferenced-whitespace-cells', MAXIMUM)
        self.assertNotIn('--cell-formula-references-only', MAXIMUM)


if __name__ == '__main__':
    unittest.main()
