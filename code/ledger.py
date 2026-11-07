"""A resumable, content-addressed extraction and conversion ledger."""
import csv
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import sqlite3
import time

SSD_RESERVE_BYTES = 6 * 1024**3


class InspectionPaused(Exception):
    """An archive stopped between committed members at the user's request."""


def inspection_lane_bounds(lane, lanes):
    """Partition SHA-256 prefixes into disjoint, exhaustive worker ranges."""
    if not 1 <= lanes <= 16 or not 0 <= lane < lanes:
        raise ValueError('Invalid inspection lane')
    alphabet = '0123456789abcdefg'
    return alphabet[16 * lane // lanes], alphabet[16 * (lane + 1) // lanes]


def inspection_priority(path, *, member=None, candidate=False, kind='unknown'):
    """Prefer tables and containers, then mail, then plain text and media.

    Extension hints affect scheduling only. Classification always reads bytes.
    """
    if candidate:
        return 0
    suffix = Path(member or path).suffix.lower()
    if kind in ('ole', 'zip', 'tar', 'gzip', 'bzip2', 'xz', 'seven', 'rtf', 'xml_container', 'legacy_excel'):
        return 10
    if suffix in {'.xls', '.xlsx', '.xlsm', '.xlsb', '.csv', '.tsv', '.slk', '.sylk', '.ods',
                  '.doc', '.docx', '.ppt', '.pptx', '.msg', '.zip', '.tar', '.gz', '.7z', '.rar', '.rtf', '.xml'}:
        return 10
    if suffix in ('.eml', '.mht', '.mhtml') or kind == 'mime':
        return 30
    if suffix in ('.txt', '.text'):
        return 40
    if suffix in ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.wav', '.mp3'):
        return 50
    return 20


def utc():
    return time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    pending = path.with_name(path.name + '.pending')
    pending.write_text(json.dumps(value, indent=2, ensure_ascii=True) + '\n', encoding='utf8')
    for attempt in range(30):
        try:
            os.replace(pending, path)
            return
        except PermissionError:
            if attempt == 29:
                raise
            time.sleep(.05)


def safe_name(name):
    name = name.replace('\\', '/')
    path = PurePosixPath(name)
    if path.is_absolute() or name.startswith('//') or re.match(r'^[A-Za-z]:', name) or '..' in path.parts:
        raise ValueError('Archive path escapes its output: ' + name)
    parts = []
    for part in path.parts:
        if part in ('', '.'):
            continue
        clean = re.sub(r'[<>:"|?*\x00-\x1f]', '_', part).rstrip(' .')
        if clean.upper().split('.')[0] in {'CON', 'PRN', 'AUX', 'NUL'} | {f'{s}{n}' for s in ('COM', 'LPT') for n in range(1, 10)}:
            clean = '_' + clean
        if not clean or len(clean) > 100 or clean != part:
            clean = (clean or 'unnamed')[:80] + '--' + hashlib.sha256(part.encode()).hexdigest()[:12]
        parts.append(clean)
    if not parts:
        raise ValueError('Empty archive path')
    return Path(*parts)


class StorageFull(RuntimeError):
    pass


class Ledger:
    def __init__(self, root, reserve=20 * 1024**3):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.reserve = reserve
        config_path = self.root / 'config.json'
        config = json.loads(config_path.read_bytes()) if config_path.is_file() else {}
        if config.get('corpus_roots'):
            raise ValueError('Select a corpus working directory; this root coordinates separate ledgers')
        self.corpus = config.get('corpus')
        self.database_reserve = config.get('database_reserve_bytes', SSD_RESERVE_BYTES)
        self.database_path = Path(config.get('database_path', self.root / 'ledger.sqlite')).resolve()
        if self.corpus and not self.database_path.is_file():
            raise FileNotFoundError('Migrated corpus database is missing: ' + str(self.database_path))
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.stop_file = Path(config.get('stop_file', self.root / 'stop.request'))
        self.sibling_databases = config.get('sibling_databases', [])
        # Archive reads and Excel calls must never retain an implicit writer lock.
        # Related metadata changes use transaction() after file I/O has finished.
        self.db = sqlite3.connect(str(self.database_path), timeout=60,
                                  isolation_level=None, uri=True)
        self.db.row_factory = sqlite3.Row
        self.db.execute('PRAGMA journal_mode=WAL')
        self.db.execute('PRAGMA busy_timeout=60000')
        self.db.executescript('''
          CREATE TABLE IF NOT EXISTS objects(
            sha TEXT PRIMARY KEY,path TEXT NOT NULL,bytes INTEGER,kind TEXT DEFAULT 'unknown',
            state TEXT DEFAULT 'pending',candidate INTEGER DEFAULT 0,stream_sha TEXT,
            verified INTEGER DEFAULT 0,conversion TEXT DEFAULT 'pending',output TEXT,
            output_sha TEXT,maximum TEXT DEFAULT 'pending',first_corpus TEXT,depth INTEGER DEFAULT 0);
          CREATE INDEX IF NOT EXISTS object_queue ON objects(state,first_corpus);
          CREATE INDEX IF NOT EXISTS conversion_queue ON objects(candidate,conversion);
          CREATE INDEX IF NOT EXISTS conversion_corpus_queue ON objects(candidate,conversion,first_corpus,state);
          CREATE TABLE IF NOT EXISTS roots(corpus TEXT,logical TEXT,sha TEXT,metadata TEXT,
            PRIMARY KEY(corpus,logical,sha));
          CREATE TABLE IF NOT EXISTS edges(parent TEXT,child TEXT,member TEXT,method TEXT,
            PRIMARY KEY(parent,member,child));
          CREATE INDEX IF NOT EXISTS edge_parent ON edges(parent);
          CREATE TABLE IF NOT EXISTS memberships(sha TEXT,corpus TEXT,PRIMARY KEY(sha,corpus));
          CREATE INDEX IF NOT EXISTS corpus_members ON memberships(corpus,sha);
          CREATE TABLE IF NOT EXISTS issues(id INTEGER PRIMARY KEY,sha TEXT,stage TEXT,detail TEXT,
            created TEXT,resolved INTEGER DEFAULT 0,UNIQUE(sha,stage,detail));
          CREATE TABLE IF NOT EXISTS reuse(sha TEXT,bytes INTEGER,bundle TEXT,baseline_sha TEXT,
            PRIMARY KEY(sha,bundle));
          CREATE TABLE IF NOT EXISTS targets(sha TEXT PRIMARY KEY,member TEXT,stream_sha TEXT);
          CREATE TABLE IF NOT EXISTS matches(target TEXT,object TEXT,evidence TEXT,
            PRIMARY KEY(target,object,evidence));
          CREATE TABLE IF NOT EXISTS feeds(name TEXT PRIMARY KEY,position INTEGER DEFAULT 0,complete INTEGER DEFAULT 0);
          CREATE TABLE IF NOT EXISTS source_cache(path TEXT PRIMARY KEY,size INTEGER,mtime INTEGER,sha TEXT);
        ''')
        self.db.commit()
        columns = {row['name'] for row in self.db.execute('PRAGMA table_info(objects)')}
        if 'inspection_revision' not in columns:
            self.db.execute('ALTER TABLE objects ADD COLUMN inspection_revision INTEGER NOT NULL DEFAULT 0')
        if 'inspection_priority' not in columns:
            self.db.execute('ALTER TABLE objects ADD COLUMN inspection_priority INTEGER NOT NULL DEFAULT 20')
        self.db.executescript('''
          CREATE INDEX IF NOT EXISTS inspection_order ON objects(state,inspection_priority,sha);
          CREATE TABLE IF NOT EXISTS inspection_reuse(
            sha TEXT PRIMARY KEY,source_database TEXT NOT NULL,source_state TEXT NOT NULL,
            source_revision INTEGER NOT NULL,imported_utc TEXT NOT NULL);
          CREATE TABLE IF NOT EXISTS inspection_sources(sha TEXT PRIMARY KEY,source_database TEXT NOT NULL);
          CREATE TABLE IF NOT EXISTS mime_bodies(
            parent TEXT,part INTEGER,content_type TEXT,bytes INTEGER,sha TEXT,
            PRIMARY KEY(parent,part));
          CREATE TABLE IF NOT EXISTS archive_text(
            parent TEXT,ordinal INTEGER,member TEXT,bytes INTEGER,sha TEXT,
            inspection_revision INTEGER NOT NULL,
            PRIMARY KEY(parent,ordinal));
        ''')

    @contextmanager
    def transaction(self):
        """Commit related metadata atomically, releasing the writer on failure.

        Callers finish file reads, extraction and external tools before entering.
        BEGIN IMMEDIATE avoids upgrading a stale read snapshot into a writer.
        """
        self.db.execute('BEGIN IMMEDIATE')
        try:
            yield
            self.db.commit()
        except BaseException:
            self.db.rollback()
            raise

    def room(self, required=0):
        if shutil.disk_usage(self.root).free < self.reserve + required:
            raise StorageFull('Data drive reserve reached; extraction and conversion are paused')
        if self.corpus and shutil.disk_usage(self.database_path.parent).free < self.database_reserve:
            raise StorageFull('SQLite SSD reserve reached; corpus processing is paused')

    def issue(self, sha, stage, detail):
        self.db.execute('''INSERT INTO issues(sha,stage,detail,created) VALUES(?,?,?,?)
            ON CONFLICT(sha,stage,detail) DO UPDATE SET resolved=0,created=excluded.created''',
                        (sha, stage, str(detail), utc()))

    def membership(self, sha, corpus):
        if self.db.execute('SELECT 1 FROM memberships WHERE sha=? AND corpus=?', (sha, corpus)).fetchone():
            return
        self.db.execute('''WITH RECURSIVE reach(sha) AS
            (SELECT ? UNION SELECT e.child FROM edges e JOIN reach r ON e.parent=r.sha)
            INSERT OR IGNORE INTO memberships SELECT sha,? FROM reach''', (sha, corpus))

    def add(self, path, corpus=None, logical=None, expected=None, size=None, metadata=None,
            parent=None, member=None, method='archive', verified=False, depth=0):
        if self.corpus:
            self.room()
            if corpus is None and parent is None:
                corpus = self.corpus
        path = Path(path).resolve()
        if self.corpus and corpus and corpus != self.corpus:
            raise ValueError('Source belongs to a different corpus database: ' + corpus)
        if self.corpus:
            inherited_corpus = self.corpus
        elif not corpus and parent:
            row=self.db.execute('SELECT first_corpus FROM objects WHERE sha=?',(parent,)).fetchone()
            inherited_corpus=row['first_corpus'] if row else ''
        else:inherited_corpus=corpus or ''
        stat = path.stat()
        if size is not None and stat.st_size != size:
            raise ValueError('Source length changed: ' + str(path))
        if expected:
            sha = expected
        else:
            previous = self.db.execute('SELECT * FROM source_cache WHERE path=?', (str(path),)).fetchone()
            if previous and previous['size'] == stat.st_size and previous['mtime'] == stat.st_mtime_ns:
                sha = previous['sha']
            else:
                sha = digest(path)
            verified = True
        existing = self.db.execute('SELECT path FROM objects WHERE sha=?', (sha,)).fetchone()
        missing_path = (existing['path'] if existing and existing['path'] != str(path)
                        and not Path(existing['path']).is_file() else None)
        with self.transaction():
            if not expected:
                self.db.execute('INSERT OR REPLACE INTO source_cache VALUES(?,?,?,?)',
                                (str(path), stat.st_size, stat.st_mtime_ns, sha))
            priority = inspection_priority(path, member=member or logical)
            self.db.execute('''INSERT OR IGNORE INTO objects(sha,path,bytes,verified,first_corpus,depth,inspection_priority)
                VALUES(?,?,?,?,?,?,?)''', (sha, str(path), stat.st_size, int(verified), inherited_corpus, depth, priority))
            if missing_path is not None:
                self.db.execute('UPDATE objects SET path=?,bytes=?,verified=? WHERE sha=? AND path=?',
                                (str(path), stat.st_size, int(verified), sha, missing_path))
            if parent:
                self.db.execute('INSERT OR IGNORE INTO edges VALUES(?,?,?,?)', (parent, sha, member, method))
                for r in self.db.execute('SELECT corpus FROM memberships WHERE sha=?', (parent,)).fetchall():
                    self.membership(sha, r['corpus'])
            if corpus:
                self.db.execute('INSERT OR IGNORE INTO roots VALUES(?,?,?,?)',
                                (corpus, logical or str(path), sha, json.dumps(metadata or {}, ensure_ascii=True)))
                self.membership(sha, corpus)
            self.db.execute('INSERT OR IGNORE INTO matches SELECT sha,?,? FROM targets WHERE sha=?',
                            (sha, 'identical_file_bytes', sha))
        return sha

    def child_path(self, parent, ordinal, name):
        relative = safe_name(name)
        base = self.root / 'extracted' / parent[:24] / f'{ordinal:07d}'
        if len(str(base / relative)) > 220:
            relative = Path('_long_paths') / (hashlib.sha256(name.encode()).hexdigest()[:16] + '-' + relative.name[:80])
        return base / relative

    def store(self, parent, ordinal, name, stream, method, depth, expected_size=None):
        self.room()
        reused = self.reused_member(parent, ordinal, name, method, expected_size)
        if reused is not None:
            return reused
        self.room(min(expected_size or 0, 1024**3))
        target = self.child_path(parent, ordinal, name)
        target.parent.mkdir(parents=True, exist_ok=True)
        pending = target.with_name(target.name + '.partial')
        h = hashlib.sha256()
        size = 0
        try:
            with pending.open('wb') as dst:
                while True:
                    block = stream.read(1024 * 1024)
                    if not block:
                        break
                    dst.write(block); h.update(block); size += len(block)
                    if size % (32 * 1024**2) == 0:
                        self.room()
            if expected_size is not None and size != expected_size:
                raise ValueError(f'Truncated member: expected {expected_size}, read {size}')
            os.replace(pending, target)
        finally:
            pending.unlink(missing_ok=True)
        sha = self.add(target, parent=parent, member=name + f' [entry {ordinal}]', method=method,
                       expected=h.hexdigest(), verified=True, depth=depth)
        if ordinal % 50 == 0:
            self.db.commit()
        return sha

    def reused_member(self, parent, ordinal, name, method, expected_size):
        """Keep a verified extraction at its existing path after moving the ledger.

        Parent, entry ordinal and extraction method identify the old member.
        Verify the saved bytes before reusing it in the new working directory.
        """
        rows = self.db.execute('''SELECT o.sha,o.path,o.bytes FROM edges e
            JOIN objects o ON o.sha=e.child WHERE e.parent=? AND e.member=?
            AND e.method=? AND o.verified=1''', (parent, name + f' [entry {ordinal}]', method)).fetchall()
        if len(rows) != 1:
            return None
        row = rows[0]
        path = Path(row['path'])
        if not path.is_file() or path.stat().st_size != row['bytes']:
            return None
        if expected_size is not None and row['bytes'] != expected_size:
            return None
        return row['sha'] if digest(path) == row['sha'] else None

    def workbook(self, sha, stream_sha=None):
        with self.transaction():
            self.db.execute('UPDATE objects SET candidate=1,stream_sha=?,inspection_priority=0 WHERE sha=?', (stream_sha, sha))
            if stream_sha:
                self.db.execute('INSERT OR IGNORE INTO matches SELECT sha,?,? FROM targets WHERE stream_sha=?',
                                (sha, 'identical_workbook_stream_not_whole_file', stream_sha))

    def claim(self, role, preferred='fuse', exclude=(), lane=0, lanes=1):
        """Atomically claim one object; conversion pools may exclude active hashes."""
        with self.transaction():
            if self.corpus:
                if role == 'inspect':
                    if lanes == 1:
                        row = self.db.execute("SELECT * FROM objects WHERE state='pending' ORDER BY inspection_priority,sha LIMIT 1").fetchone()
                    else:
                        lower, upper = inspection_lane_bounds(lane, lanes)
                        row = None
                        # The queue index seeks directly into each priority and
                        # hash range, without scanning another lane's objects.
                        for priority in (0, 10, 20, 30, 40, 50):
                            row = self.db.execute("SELECT * FROM objects WHERE state='pending' AND inspection_priority=? AND sha>=? AND sha<? ORDER BY sha LIMIT 1",
                                                  (priority, lower, upper)).fetchone()
                            if row is not None:
                                break
                    if row:
                        self.db.execute("UPDATE objects SET state='working' WHERE sha=?", (row['sha'],))
                else:
                    sql = "SELECT * FROM objects WHERE candidate=1 AND conversion='pending' AND state IN ('complete','review')"
                    if exclude:
                        sql += ' AND sha NOT IN (' + ','.join('?' for _ in exclude) + ')'
                    row = self.db.execute(sql + ' LIMIT 1', tuple(exclude)).fetchone()
                    if row:
                        self.db.execute("UPDATE objects SET conversion='working' WHERE sha=?", (row['sha'],))
                return dict(row) if row else None
            row = None
            # Small indexed lookups keep FUSE and PST work ahead of large archives.
            for corpus in dict.fromkeys((preferred, 'enron-pst', 'fuse')):
                if role == 'inspect':
                    row = self.db.execute("SELECT * FROM objects WHERE state='pending' AND first_corpus=? LIMIT 1", (corpus,)).fetchone()
                else:
                    row = self.db.execute("SELECT * FROM objects WHERE candidate=1 AND conversion='pending' AND first_corpus=? AND state IN ('complete','review') LIMIT 1", (corpus,)).fetchone()
                if row is not None:
                    break
            if role == 'inspect':
                if row is None:
                    row = self.db.execute("SELECT * FROM objects WHERE state='pending' LIMIT 1").fetchone()
                if row:
                    self.db.execute("UPDATE objects SET state='working' WHERE sha=?", (row['sha'],))
            else:
                if row is None:
                    row = self.db.execute("SELECT * FROM objects WHERE candidate=1 AND conversion='pending' AND state IN ('complete','review') LIMIT 1").fetchone()
                if row:
                    self.db.execute("UPDATE objects SET conversion='working' WHERE sha=?", (row['sha'],))
        return dict(row) if row else None

    def reuse_candidates(self, sha):
        """Find exact-source conversions across independent corpus databases.

        Sibling connections are read-only and close before file copying starts.
        """
        candidates = [dict(row) for row in self.db.execute('SELECT * FROM reuse WHERE sha=?', (sha,))]
        for name in self.sibling_databases:
            path = Path(name)
            if path.resolve() == self.database_path:
                continue
            connection = sqlite3.connect(path.as_uri() + '?mode=ro', uri=True, timeout=10)
            try:
                row = connection.execute("SELECT bytes,output,output_sha FROM objects WHERE sha=? AND conversion IN ('accepted','reused') AND output IS NOT NULL", (sha,)).fetchone()
                if row:
                    candidates.append({'sha': sha, 'bytes': row[0], 'bundle': row[1], 'baseline_sha': row[2]})
            finally:
                connection.close()
        return candidates

    def corpus_counts(self):
        """Count one owned corpus using queue indices instead of a provenance join."""
        return {
            'unique_objects': self.db.execute('SELECT COUNT(*) FROM objects').fetchone()[0],
            'inspected': self.db.execute("SELECT COUNT(*) FROM objects WHERE state IN ('complete','review')").fetchone()[0],
            'workbooks': self.db.execute('SELECT COUNT(*) FROM objects WHERE candidate=1').fetchone()[0],
            'xlsx': self.db.execute("SELECT COUNT(*) FROM objects WHERE candidate=1 AND conversion IN ('accepted','reused')").fetchone()[0],
            'conversion_review': self.db.execute("SELECT COUNT(*) FROM objects WHERE candidate=1 AND conversion='review'").fetchone()[0],
        }

    def status(self):
        corpora = {}
        if self.corpus:
            corpora[self.corpus] = self.corpus_counts()
        else:
            for r in self.db.execute('''SELECT m.corpus,o.state,o.candidate,o.conversion,COUNT(*) n
                FROM memberships m JOIN objects o ON o.sha=m.sha GROUP BY 1,2,3,4'''):
                c = corpora.setdefault(r['corpus'], {'unique_objects':0,'inspected':0,'workbooks':0,'xlsx':0,'conversion_review':0})
                c['unique_objects'] += r['n']
                c['inspected'] += r['n'] if r['state'] in ('complete','review') else 0
                if r['candidate']:
                    c['workbooks'] += r['n']
                    c['xlsx'] += r['n'] if r['conversion'] in ('accepted','reused') else 0
                    c['conversion_review'] += r['n'] if r['conversion'] == 'review' else 0
        matched = self.db.execute('''SELECT COUNT(DISTINCT target) FROM matches x JOIN memberships m
            ON m.sha=x.object WHERE m.corpus IN ('enron-pst','enron-xml')''').fetchone()[0]
        return {'updated_utc':utc(),'corpora':corpora,'open_issues':self.db.execute('SELECT COUNT(*) FROM issues WHERE resolved=0').fetchone()[0],
                'old_enron_targets':self.db.execute('SELECT COUNT(*) FROM targets').fetchone()[0],
                'old_enron_targets_matched_in_new_sources':matched}

    def export(self, names=None):
        for name, sql in {
            'open-issues': 'SELECT * FROM issues WHERE resolved=0 ORDER BY id',
            'workbooks': 'SELECT sha,path,kind,bytes,conversion,output,output_sha FROM objects WHERE candidate=1',
            'enron-matches': "SELECT x.*,m.corpus,o.path FROM matches x JOIN memberships m ON m.sha=x.object JOIN objects o ON o.sha=x.object WHERE m.corpus IN ('enron-pst','enron-xml')",
            'container-members': 'SELECT * FROM edges',
            'sources': 'SELECT * FROM roots',
        }.items():
            if names is not None and name not in names:
                continue
            path = self.root / 'reports' / (name + '.csv'); path.parent.mkdir(exist_ok=True)
            cursor = self.db.execute(sql)
            with path.open('w',newline='',encoding='utf-8-sig') as f:
                writer=csv.writer(f);writer.writerow([c[0] for c in cursor.description]);writer.writerows(cursor)
