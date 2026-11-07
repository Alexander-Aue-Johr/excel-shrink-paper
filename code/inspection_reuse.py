"""Reuse checksum-verified inspection graphs across the independent ledgers.

Source ledgers are read-only. File verification finishes before target write
transactions. Graph edges, corpus membership and unresolved inspection issues
are imported before any inspected state is published.
"""

from collections import defaultdict, deque
from pathlib import Path
import os
import sqlite3

from containers import INSPECTION_REVISION, classify
from ledger import digest, inspection_priority, utc

CONVERSION_STAGES = {'conversion', 'conversion_worker', 'reuse_unavailable',
                     'reconstruction_review', 'maximum_cleaning'}


def refresh_sources(target, sources):
    """Index reusable source hashes locally so ordinary files need no sibling lookup."""
    for source in sources:
        path = Path(source).resolve()
        if path == target.database_path:
            continue
        connection = sqlite3.connect(path.as_uri() + '?mode=ro', uri=True, timeout=30)
        try:
            rows = connection.execute("SELECT sha FROM objects WHERE verified=1 AND state IN ('complete','review') AND (kind<>'other' OR inspection_revision>=?)", (INSPECTION_REVISION,))
            while True:
                batch = rows.fetchmany(2000)
                if not batch:
                    break
                with target.transaction():
                    target.db.executemany('INSERT OR IGNORE INTO inspection_sources VALUES(?,?)',
                                          ((row[0], str(path)) for row in batch))
        finally:
            connection.close()


def load_graph(source, roots=None):
    """Read one consistent source snapshot, releasing it before filesystem I/O."""
    connection = sqlite3.connect(Path(source).resolve().as_uri() + '?mode=ro', uri=True, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute('BEGIN')
    try:
        has_archive_text = connection.execute("SELECT 1 FROM sqlite_master WHERE name='archive_text' AND type='table'").fetchone() is not None
        if roots is None:
            objects = {row['sha']: dict(row) for row in connection.execute('SELECT * FROM objects')}
            edges = [dict(row) for row in connection.execute('SELECT * FROM edges')]
        else:
            objects, edges = {}, []
            pending = list(roots)
            while pending:
                sha = pending.pop()
                if sha in objects:
                    continue
                row = connection.execute('SELECT * FROM objects WHERE sha=?', (sha,)).fetchone()
                if row is None:
                    raise ValueError('Inspection graph references a missing object: ' + sha)
                objects[sha] = dict(row)
                children = [dict(row) for row in connection.execute('SELECT * FROM edges WHERE parent=?', (sha,))]
                edges.extend(children)
                pending.extend(row['child'] for row in children)
        if roots is None:
            issues = [dict(row) for row in connection.execute('SELECT * FROM issues')
                      if row['stage'] not in CONVERSION_STAGES]
            matches = [tuple(row) for row in connection.execute('SELECT * FROM matches')]
            bodies = [dict(row) for row in connection.execute('SELECT * FROM mime_bodies')]
            archive_text = [dict(row) for row in connection.execute('SELECT * FROM archive_text')] if has_archive_text else []
        else:
            issues, matches, bodies = [], [], []
            archive_text = []
            keys = list(objects)
            for start in range(0, len(keys), 500):
                batch = keys[start:start + 500]
                parameters = ','.join('?' for _ in batch)
                issues.extend(dict(row) for row in connection.execute('SELECT * FROM issues WHERE sha IN (' + parameters + ')', batch)
                              if row['stage'] not in CONVERSION_STAGES)
                matches.extend(tuple(row) for row in connection.execute('SELECT * FROM matches WHERE object IN (' + parameters + ')', batch))
                bodies.extend(dict(row) for row in connection.execute('SELECT * FROM mime_bodies WHERE parent IN (' + parameters + ')', batch))
                if has_archive_text:
                    archive_text.extend(dict(row) for row in connection.execute('SELECT * FROM archive_text WHERE parent IN (' + parameters + ')', batch))
        return objects, edges, issues, matches, bodies, archive_text
    finally:
        connection.close()


def import_graph(target, source, roots, *, claimed_root=None, status=None, verify_payloads=True):
    """Import reachable objects without replacing target conversions or provenance.

    Legacy `other` objects receive the current content sniff so newly supported
    text spreadsheets are not hidden by an earlier generic classification.
    The offline bulk pass may reuse ingestion's verified hashes for immutable
    owned payloads, checking existence and length instead of rereading all bytes.
    """
    roots = set(roots)
    if not roots:
        return {'reused_roots': 0, 'objects': 0}
    objects, edges, issues, matches, bodies, archive_text = load_graph(source, None if len(roots) > 100 else roots)
    if not verify_payloads:
        roots = {sha for sha in roots if objects[sha]['kind'] != 'other'
                 or objects[sha].get('inspection_revision', 0) >= INSPECTION_REVISION}
        if not roots:
            return {'reused_roots': 0, 'objects': 0}
    children, parents = defaultdict(list), defaultdict(list)
    for edge in edges:
        children[edge['parent']].append(edge['child'])
        parents[edge['child']].append(edge['parent'])
    reachable, todo = set(), list(roots)
    while todo:
        sha = todo.pop()
        if sha in reachable:
            continue
        if sha not in objects:
            raise ValueError('Missing inspection source object: ' + sha)
        reachable.add(sha)
        todo.extend(children[sha])
    invalid = set()
    classified = {}
    file_cache = {}
    directories = {}

    def stamp_for(path):
        if verify_payloads:
            stamp = path.stat()
            return stamp.st_size, stamp.st_mtime_ns
        # On Windows scandir supplies size/time in the directory record. One
        # directory pass avoids reopening thousands of content-addressed files.
        parent = str(path.parent)
        if parent not in directories:
            entries = {}
            with os.scandir(parent) as directory:
                for entry in directory:
                    if entry.is_file(follow_symlinks=False):
                        stamp = entry.stat(follow_symlinks=False)
                        entries[entry.name] = (stamp.st_size, stamp.st_mtime_ns)
            directories[parent] = entries
        return directories[parent].get(path.name)

    def verified(path, sha, size, recorded=False):
        path = Path(path)
        try:
            stamp = stamp_for(path)
        except OSError:
            return False
        if stamp is None:
            return False
        size_now, modified = stamp
        key = (str(path), size_now, modified, sha)
        if key not in file_cache:
            try:
                file_cache[key] = size_now == size and (
                    (recorded and not verify_payloads) or digest(path) == sha)
            except OSError:
                file_cache[key] = False
        return file_cache[key]

    for index, sha in enumerate(sorted(reachable, key=lambda sha: objects[sha]['path']), 1):
        obj = objects[sha]
        if not verified(obj['path'], sha, obj['bytes'], recorded=bool(obj['verified'])):
            invalid.add(sha)
        elif obj['state'] in ('complete', 'review'):
            kind = obj['kind']
            if obj.get('inspection_revision', 0) < INSPECTION_REVISION and kind == 'other' and verify_payloads:
                kind = classify(obj['path'])
            classified[sha] = kind
        if status and index % 1000 == 0:
            status({'stage': 'checking_registered_payloads' if not verify_payloads else 'verifying_payloads',
                    'checked': index, 'objects': len(reachable)})
    # A missing descendant requires extraction from a readable ancestor again.
    todo = list(invalid)
    while todo:
        sha = todo.pop()
        for parent in parents[sha]:
            if parent not in invalid:
                invalid.add(parent)
                todo.append(parent)
    usable_roots = set()
    for sha in roots - invalid:
        obj = objects[sha]
        if obj['state'] not in ('complete', 'review') or not obj['verified']:
            continue
        current = target.db.execute('SELECT path,bytes,state,depth,verified FROM objects WHERE sha=?', (sha,)).fetchone()
        if current is None or current['state'] not in ('pending', 'working'):
            continue
        if current['state'] == 'working' and sha != claimed_root:
            continue
        if classified.get(sha) != obj['kind'] and classified.get(sha) != 'legacy_excel':
            continue
        same_owned_path = current['path'] == obj['path'] or (
            Path(current['path']).resolve() == Path(obj['path']).resolve())
        if verified(current['path'], sha, current['bytes'], recorded=bool(current['verified'] and same_owned_path)):
            usable_roots.add(sha)
    selected, depths = set(), {}
    todo = deque()
    for sha in usable_roots:
        depth = target.db.execute('SELECT depth FROM objects WHERE sha=?', (sha,)).fetchone()[0]
        todo.append((sha, depth))
    while todo:
        sha, depth = todo.popleft()
        if sha in depths and depths[sha] <= depth:
            continue
        depths[sha] = depth
        selected.add(sha)
        todo.extend((child, depth + 1) for child in children[sha])
    states = []
    rows = []
    reuse = []
    for sha in selected:
        obj = objects[sha]
        kind = classified.get(sha, obj['kind'])
        compatible = (kind == obj['kind'] or kind == 'legacy_excel') and not (
            not verify_payloads and kind == 'other' and obj.get('inspection_revision', 0) < INSPECTION_REVISION)
        inspected = obj['state'] in ('complete', 'review') and obj['verified'] and compatible
        candidate = int(obj['candidate'] or kind == 'legacy_excel')
        priority = inspection_priority(obj['path'], candidate=bool(candidate), kind=kind)
        rows.append((sha, obj['path'], obj['bytes'], kind, candidate, obj['stream_sha'],
                     obj['verified'], obj['first_corpus'], depths[sha], priority))
        if inspected:
            states.append((kind, candidate, obj['stream_sha'], obj['state'], INSPECTION_REVISION, priority, sha))
        if obj['conversion'] in ('accepted', 'reused') and obj['output'] and obj['output_sha']:
            reuse.append((sha, obj['bytes'], obj['output'], obj['output_sha']))
    for start in range(0, len(rows), 1000):
        target.room()
        with target.transaction():
            target.db.executemany('''INSERT OR IGNORE INTO objects
                (sha,path,bytes,kind,candidate,stream_sha,verified,first_corpus,depth,inspection_priority)
                VALUES(?,?,?,?,?,?,?,?,?,?)''', rows[start:start + 1000])
    with target.transaction():
        target.db.executemany('INSERT OR IGNORE INTO edges VALUES(?,?,?,?)',
                              ((r['parent'], r['child'], r['member'], r['method']) for r in edges
                               if r['parent'] in selected and r['child'] in selected))
        target.db.executemany('INSERT OR IGNORE INTO memberships VALUES(?,?)', ((sha, target.corpus) for sha in selected))
        target.db.executemany('INSERT OR IGNORE INTO issues(sha,stage,detail,created,resolved) VALUES(?,?,?,?,?)',
                              ((r['sha'], r['stage'], r['detail'], r['created'], r['resolved']) for r in issues if r['sha'] in selected))
        target.db.executemany('INSERT OR IGNORE INTO matches VALUES(?,?,?)', (r for r in matches if r[1] in selected))
        target.db.executemany('INSERT OR IGNORE INTO mime_bodies VALUES(?,?,?,?,?)',
                              ((r['parent'], r['part'], r['content_type'], r['bytes'], r['sha']) for r in bodies if r['parent'] in selected))
        target.db.executemany('INSERT OR IGNORE INTO archive_text VALUES(?,?,?,?,?,?)',
                              ((r['parent'], r['ordinal'], r['member'], r['bytes'], r['sha'], r['inspection_revision'])
                               for r in archive_text if r['parent'] in selected))
        target.db.executemany('INSERT OR IGNORE INTO reuse VALUES(?,?,?,?)', reuse)
    for start in range(0, len(states), 1000):
        target.room()
        batch = states[start:start + 1000]
        with target.transaction():
            target.db.executemany('''UPDATE objects SET kind=?,candidate=MAX(candidate,?),stream_sha=COALESCE(?,stream_sha),
                state=?,inspection_revision=?,inspection_priority=?
                WHERE sha=? AND (state='pending' OR (state='working' AND sha=?))''',
                                  (row + (claimed_root or '',) for row in batch))
            target.db.executemany('INSERT OR REPLACE INTO inspection_reuse VALUES(?,?,?,?,?)',
                                  ((r[6], str(source), r[3], objects[r[6]].get('inspection_revision', 0), utc()) for r in batch))
    return {'reused_roots': len(usable_roots), 'objects': len(selected), 'inspection_states': len(states),
            'unavailable_roots': len(roots & invalid)}


def try_reuse(target, obj):
    """Reuse one claimed object only when a local hash index names a source."""
    row = target.db.execute('SELECT source_database FROM inspection_sources WHERE sha=?', (obj['sha'],)).fetchone()
    if row is None:
        return False
    result = import_graph(target, row['source_database'], [obj['sha']], claimed_root=obj['sha'])
    return result['reused_roots'] == 1
