"""Exercise writer isolation and resumable metadata with real SQLite connections."""

from concurrent.futures import ThreadPoolExecutor
import io
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import ledger
from ledger import Ledger
from supervisor import StatusSampler


class LedgerTransactionTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.first = Ledger(self.root, reserve=0)
        self.second = Ledger(self.root, reserve=0)
        self.second.db.execute('PRAGMA busy_timeout=100')

    def tearDown(self):
        self.first.db.close()
        self.second.db.close()
        self.temporary.cleanup()

    def source(self, name):
        path = self.root / name
        path.write_bytes(name.encode())
        return path

    def test_archive_streaming_does_not_hold_previous_member_transaction(self):
        parent = self.first.add(self.source('archive.zip'), corpus='fuse')
        self.first.store(parent, 1, 'first.xls', io.BytesIO(b'first'), 'zip', 1)

        class NextMember(io.BytesIO):
            def read(stream, size=-1):
                self.assertFalse(self.first.db.in_transaction)
                self.second.issue(parent, 'concurrent_worker', 'Writer progressed during stream read')
                return super().read(size)

        self.first.store(parent, 2, 'second.xls', NextMember(b'second'), 'zip', 1)
        self.assertEqual(self.second.db.execute('SELECT COUNT(*) FROM edges').fetchone()[0], 2)

    def test_hashing_and_issue_reporting_release_the_writer(self):
        self.first.issue('prior', 'reuse_unavailable', 'Continue with Excel')
        original_digest = ledger.digest

        def concurrent_digest(path):
            self.assertFalse(self.first.db.in_transaction)
            self.second.issue('peer', 'progress', 'Hashing does not block metadata')
            return original_digest(path)

        with patch.object(ledger, 'digest', side_effect=concurrent_digest):
            sha = self.first.add(self.source('workbook.xls'), corpus='enron-pst')
        self.assertIsNotNone(self.second.db.execute('SELECT sha FROM objects WHERE sha=?', (sha,)).fetchone())

    def test_failed_registration_rolls_back_object_and_source_cache(self):
        with patch.object(self.first, 'membership', side_effect=RuntimeError('Registration failed')):
            with self.assertRaisesRegex(RuntimeError, 'Registration failed'):
                self.first.add(self.source('failed.xls'), corpus='fuse')
        self.assertFalse(self.first.db.in_transaction)
        self.assertEqual(self.second.db.execute('SELECT COUNT(*) FROM objects').fetchone()[0], 0)
        self.assertEqual(self.second.db.execute('SELECT COUNT(*) FROM source_cache').fetchone()[0], 0)
        self.second.issue('peer', 'progress', 'Rollback released the writer')

    def test_repeated_issue_reopens_a_resolved_infrastructure_error(self):
        self.first.issue('archive', 'zip_member', 'book.xls: database is locked')
        self.first.db.execute('UPDATE issues SET resolved=1')
        self.second.issue('archive', 'zip_member', 'book.xls: database is locked')
        row = self.first.db.execute('SELECT resolved FROM issues').fetchone()
        self.assertEqual(row[0], 0)

    def test_two_workers_claim_different_objects(self):
        expected = {self.first.add(self.source(name), corpus='fuse') for name in ('one.xls', 'two.xls')}

        def claim_one(_index):
            worker = Ledger(self.root, reserve=0)
            try:
                return worker.claim('inspect', 'fuse')['sha']
            finally:
                worker.db.close()

        with ThreadPoolExecutor(max_workers=2) as pool:
            claimed = set(pool.map(claim_one, range(2)))
        self.assertEqual(claimed, expected)

    def test_pst_work_precedes_other_archives_when_fuse_queue_is_empty(self):
        self.first.add(self.source('large-archive.zip'), corpus='enron-xml')
        pst = self.first.add(self.source('attachment.xls'), corpus='enron-pst')
        self.assertEqual(self.first.claim('inspect', 'fuse')['sha'], pst)

    def test_replayed_member_preserves_progress_and_provenance(self):
        parent = self.first.add(self.source('archive.zip'), corpus='fuse')
        sha = self.first.store(parent, 1, 'book.xls', io.BytesIO(b'workbook'), 'zip', 1)
        self.first.db.execute("UPDATE objects SET conversion='accepted',output='existing.xlsx' WHERE sha=?", (sha,))
        self.first.store(parent, 1, 'book.xls', io.BytesIO(b'workbook'), 'zip', 1)
        row = self.second.db.execute('SELECT conversion,output FROM objects WHERE sha=?', (sha,)).fetchone()
        self.assertEqual(tuple(row), ('accepted', 'existing.xlsx'))
        self.assertEqual(self.second.db.execute('SELECT COUNT(*) FROM edges').fetchone()[0], 1)


class StatusSamplerTests(unittest.TestCase):
    def test_restart_reuses_statistics_without_claiming_a_fresh_scan(self):
        previous = {'updated_utc': 'heartbeat', 'statistics_updated_utc': 'statistics', 'corpora': {}}
        sampler = StatusSampler(interval=1800, initial=previous)
        with patch.object(Ledger, 'status', side_effect=AssertionError('Unexpected full scan')):
            status = sampler.collect(object(), sampler.last)
        self.assertEqual(status['statistics_updated_utc'], 'statistics')


if __name__ == '__main__':
    unittest.main()
