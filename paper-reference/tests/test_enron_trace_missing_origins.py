import argparse
import contextlib
import csv
import hashlib
import io
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from excel_shrink_research.enron.trace_missing_origins import mailbox_identity, run


class OriginTests(unittest.TestCase):
    def test_split_mailbox_and_export_number_are_distinct(self):
        row = mailbox_identity('richard_sanders_000_1_1_1.pst.123.xls')
        self.assertEqual('richard_sanders_000_1_1_1.pst', row['source_pst'])
        self.assertEqual('123', row['export_suffix'])
        self.assertEqual('sanders-r', row['edrm_custodian_code'])
        self.assertEqual('mclaughlin-e', mailbox_identity('errol_mclaughlin_jr_001_1_1.pst.1.xls')['edrm_custodian_code'])

    def test_size_lead_is_not_promoted_to_identity(self):
        self.check_metadata_identity(False)

    def test_metadata_hash_match_is_verified_against_actual_bytes(self):
        self.check_metadata_identity(True)

    def check_metadata_identity(self, identical):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'raw').mkdir()
            (root / 'comparison').mkdir()
            old = b'old bytes'
            new = old if identical else b'new bytes'
            md5 = hashlib.md5(new).hexdigest()
            name = 'export/nuix/eric_saibi_000_1_1.pst.1.xls'
            with zipfile.ZipFile(root / 'old.zip', 'w') as archive:
                archive.writestr('export/README.md', 'Nuix PST extraction description')
                archive.writestr(name, old)
                archive.writestr('export/artifacts/md5.lst', '')
            with (root / 'comparison/old-not-found.csv').open('w', newline='') as stream:
                writer = csv.DictWriter(stream, fieldnames=['sha256', 'kind', 'index_chain', 'member_chain'])
                writer.writeheader()
                writer.writerow(dict(sha256=hashlib.sha256(old).hexdigest(), kind='text', index_chain='[1]', member_chain=json.dumps([name])))
            xml = ('<Root><Document DocID="3.1.A" DocType="Message"><Tags><Tag TagName="#Subject" TagValue="Test"/></Tags></Document>'
                   '<Document DocID="3.1.A.1" DocType="File"><Tags><Tag TagName="#FileName" TagValue="original.tsv"/></Tags><Files>'
                   '<File FileType="Native"><ExternalFile FilePath="native" FileName="file.tsv" FileSize="9" Hash="%s"/></File>'
                   '</Files></Document><Relationship Type="Attachment" ParentDocId="3.1.A" ChildDocId="3.1.A.1"/></Root>') % md5
            with zipfile.ZipFile(root / 'raw/edrm-enron-v2_saibi-e_xml.zip', 'w') as archive:
                archive.writestr('metadata.xml', xml)
                archive.writestr('native/file.tsv', new)
            with contextlib.redirect_stdout(io.StringIO()):
                run(argparse.Namespace(root=str(root), old_zip=str(root / 'old.zip'), output=str(root / 'result')))
            leads = json.loads((root / 'result/metadata-leads.json').read_text())
            hits = json.loads((root / 'result/verified-metadata-matches.json').read_text())
            findings = hits if identical else leads
            self.assertTrue(findings[0]['parent_relationship_verified'])
            self.assertEqual('Test', findings[0]['email']['subject'])
            self.assertEqual(int(identical), len(hits))
            self.assertEqual(int(identical), json.loads((root / 'result/origins.json').read_text())[0]['verified_metadata_matches'])


if __name__ == '__main__':
    unittest.main()
