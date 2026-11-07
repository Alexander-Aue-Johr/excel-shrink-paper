import json
from argparse import Namespace
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from excel_shrink_research.enron import excel
from excel_shrink_research.convert_workbooks import safe_name
from excel_shrink_research import convert_workbooks as batch


class ConversionRecoveryTests(unittest.TestCase):
    def test_shared_save_failure_retries_exclusive_and_keeps_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source.xls'
            source.write_bytes(b'legacy source')
            def phase(mode, local, output, *args):
                if mode == 'convert':
                    return {'status': 'review', 'stage': 'save', 'reason': 'excel_error', 'shared_workbook': True}
                if mode == 'convert_exclusive':
                    excel._create_seed(output)
                return {'status': 'accepted', 'sheet_names': ['Sheet1'], 'shared_workbook': False}
            with mock.patch.object(excel, '_excel_executable', return_value=Path('excel.exe')), mock.patch.object(excel, '_invoke_phase', side_effect=phase) as call:
                result = excel.convert_file(source, root / 'output.xlsx', root / 'attempt')
            self.assertEqual('accepted', result['status'])
            self.assertIn('legacy_shared_revision_history_removed', result['normalizations'])
            self.assertEqual(['convert', 'convert_exclusive', 'verify'], [c.args[0] for c in call.call_args_list])
            self.assertEqual(b'legacy source', source.read_bytes())

    def test_password_failure_is_explicit_and_not_retried(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source.xls'
            source.write_bytes(b'legacy')
            failure = {'status': 'review', 'stage': 'open', 'reason': 'excel_error', 'error': 'Das Kennwort ist ungültig.'}
            with mock.patch.object(excel, '_excel_executable', return_value=Path('excel.exe')), mock.patch.object(excel, '_invoke_phase', return_value=failure) as call:
                result = excel.convert_file(source, root / 'output.xlsx', root / 'attempt')
            self.assertEqual('password_rejected', result['reason'])
            self.assertEqual(1, call.call_count)
            self.assertFalse((root / 'output.xlsx').exists())

    def test_changed_sheet_list_is_not_published(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source.xls'
            source.write_bytes(b'legacy')
            def phase(mode, local, output, *args):
                if mode == 'convert':
                    excel._create_seed(output)
                return {'status': 'accepted', 'sheet_names': ['Before' if mode == 'convert' else 'After']}
            with mock.patch.object(excel, '_excel_executable', return_value=Path('excel.exe')), mock.patch.object(excel, '_invoke_phase', side_effect=phase):
                result = excel.convert_file(source, root / 'output.xlsx', root / 'attempt')
            self.assertEqual('sheet_list_changed', result['reason'])
            self.assertFalse((root / 'output.xlsx').exists())

    def test_archive_paths_cannot_escape_output(self):
        for name in ('../a.xls', '/a.xls', 'C:/a.xls', 'folder/../../a.xls', '\\\\server\\file.xls'):
            with self.subTest(name=name), self.assertRaises(ValueError):
                safe_name(name)

    def test_resume_reuses_verified_output_but_rejects_changed_bytes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'inputs'
            source.mkdir()
            excel._create_seed(source / 'a.xlsx')
            args = Namespace(input=source, output=root / 'output', members_file=None,
                             stage_timeout=10, file_timeout=20)
            def convert(local, target, *rest):
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(local.read_bytes())
                return {'status': 'accepted', 'source_sha256': batch.digest(local), 'output_sha256': batch.digest(target)}
            with mock.patch.object(batch, 'preflight'), mock.patch.object(batch, 'convert_file', side_effect=convert) as call:
                batch.run(args)
                batch.run(args)
                self.assertEqual(1, call.call_count)
                target = next((args.output / 'xlsx').rglob('*.xlsx'))
                target.write_bytes(b'changed')
                batch.run(args)
                self.assertEqual(2, call.call_count)


if __name__ == '__main__':
    unittest.main()
