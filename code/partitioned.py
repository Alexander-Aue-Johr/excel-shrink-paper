"""Coordinate independent corpus ledgers while keeping corpus files on F."""

import json
from pathlib import Path
import shutil
import subprocess
import sys
import time

from ledger import Ledger, SSD_RESERVE_BYTES, digest, utc, write_json


def corpus_roots(root):
    """Read the explicit corpus-to-working-directory mapping."""
    config = json.loads((root / 'config.json').read_bytes())
    return {name: Path(path) for name, path in config['corpus_roots'].items()}


def inspection_pending(root):
    """Include active work when deciding whether priority inspection has finished."""
    database = Ledger(root)
    try:
        return database.db.execute("SELECT 1 FROM objects WHERE state IN ('pending','working') LIMIT 1").fetchone() is not None
    finally:
        database.db.close()


def run_shared_worker(root, role):
    """Dispatch the bounded converter or inspect fair, prioritized batches."""
    if role == 'convert-all':
        from conversion_pool import run
        return run(root)
    if role != 'inspect-rest':
        raise ValueError('Unknown shared worker role: ' + role)
    from excel_shrink_research.enron.common import RunLock
    from workers import inspect_loop
    roots = corpus_roots(root)
    parallel = {name for name, child in roots.items()
                if json.loads((child / 'config.json').read_bytes()).get('inspection_workers', 1) > 1}
    with RunLock(root / 'locks' / role):
        while not (root / 'stop.request').exists():
            processed = 0
            if any(inspection_pending(roots[name])
                                            for name in ('fuse', 'enron-pst')
                                            if not (roots[name] / 'pause.request').exists()):
                time.sleep(10)
                continue
            for name, child in roots.items():
                if (root / 'stop.request').exists():
                    return
                if (child / 'pause.request').exists():
                    continue
                if name in ('fuse', 'enron-pst') or name in parallel:
                    continue
                processed += inspect_loop(child, limit=256)
            if processed == 0:
                time.sleep(3)


def run_seeding(root):
    """Resume each source feed from its migrated per-corpus journal position."""
    from excel_shrink_research.enron.common import RunLock
    from seeding import run
    from inspection_reuse import refresh_sources
    refreshed = {}
    with RunLock(root / 'locks' / 'seed-all'):
        while not (root / 'stop.request').exists():
            for child in corpus_roots(root).values():
                if (root / 'stop.request').exists():
                    return
                if (child / 'pause.request').exists():
                    continue
                sources = json.loads((child / 'config.json').read_bytes()).get('inspection_reuse_sources', [])
                if sources and time.monotonic() - refreshed.get(str(child), 0) >= 60:
                    database = Ledger(child)
                    try:
                        refresh_sources(database, sources)
                    finally:
                        database.db.close()
                    refreshed[str(child)] = time.monotonic()
                run(child, once=True)
            time.sleep(10)


def collect_statistics(databases, legacy_root=None):
    """Aggregate the same corpus statistics from the independent SSD ledgers."""
    corpora = {}
    open_issues = 0
    targets = set()
    matches = set()
    needle_names = set()
    for name, database in databases.items():
        summary = database.status()
        database.export(names=('open-issues', 'workbooks', 'enron-matches'))
        corpora.update(summary['corpora'])
        open_issues += summary['open_issues']
        targets.update(row[0] for row in database.db.execute('SELECT sha FROM targets'))
        if name in ('enron-pst', 'enron-xml'):
            matches.update(row[0] for row in database.db.execute('SELECT DISTINCT target FROM matches'))
        needle_names.update(path.name for path in (database.root / 'needles').glob('*.json'))
    if legacy_root:
        needle_names.update(path.name for path in (legacy_root / 'needles').glob('*.json'))
    return {'corpora': corpora, 'open_issues': open_issues,
            'old_enron_targets': len(targets), 'old_enron_targets_matched_in_new_sources': len(matches),
            'needle_candidates': len(needle_names), 'statistics_updated_utc': utc()}


def start_worker(root, name, command):
    """Start one owned worker with its logs and process receipt on F."""
    logs = root / 'logs'
    logs.mkdir(exist_ok=True)
    with (logs / (name + '.stdout.log')).open('ab') as out, (logs / (name + '.stderr.log')).open('ab') as err:
        process = subprocess.Popen(command, stdout=out, stderr=err, cwd=root,
                                   creationflags=0x08000000 if sys.platform == 'win32' else 0)
    write_json(root / (name + '-launch.json'), {'pid': process.pid, 'command': command, 'started_utc': utc()})
    return process


def run_supervisor(root, config):
    """Run two dedicated inspectors against different databases and file roots."""
    from excel_shrink_research.enron.common import RunLock
    from supervisor import queue_finished
    for record in json.loads((root / 'code-files.json').read_bytes()):
        if digest(root / 'code' / record['path']) != record['sha256']:
            raise ValueError('Frozen pipeline code changed: ' + record['path'])
    roots = corpus_roots(root)
    requirements = {name: set(json.loads((child / 'config.json').read_bytes())['required_feeds'])
                    for name, child in roots.items()}
    databases = {name: Ledger(child, config['reserve_bytes']) for name, child in roots.items()}
    script = root / 'code' / 'workers.py'
    commands = {
        'inspect-fuse': [sys.executable, str(script), 'inspect', '--root', str(roots['fuse'])],
        'inspect-pst': [sys.executable, str(script), 'inspect', '--root', str(roots['enron-pst'])],
        'inspect-rest': [sys.executable, str(script), 'inspect-rest', '--root', str(root)],
        'convert': [sys.executable, str(script), 'convert-all', '--root', str(root)],
        'maximum': [sys.executable, str(script), 'maximum', '--root', str(roots['fuse'])],
        'seed': [sys.executable, str(root / 'code' / 'seeding.py'), '--all', '--root', str(root)],
    }
    parallel = {}
    for name, child in roots.items():
        count = json.loads((child / 'config.json').read_bytes()).get('inspection_workers', 1)
        if count <= 1:
            continue
        parallel[name] = count
        dedicated = {'fuse': 'inspect-fuse', 'enron-pst': 'inspect-pst'}
        commands.pop(dedicated.get(name), None)
        for lane in range(count):
            commands['inspect-' + name + '-' + str(lane)] = [
                sys.executable, str(script), 'inspect', '--root', str(child),
                '--lane', str(lane), '--lanes', str(count)]
    children = {}
    statistics = collect_statistics(databases, root)
    last_statistics = time.monotonic()
    reader_file = root / 'pst-all' / 'reader-status.json'
    try:
        with RunLock(root / 'locks' / 'supervisor'):
            while not (root / 'stop.request').exists():
                reader_finished = reader_file.is_file() and json.loads(reader_file.read_bytes()).get('state') in ('complete', 'complete_with_issues')
                if not reader_finished:
                    commands['pst-reader'] = [config['node'], '--max-old-space-size=2048', str(root / 'code' / 'read_pst_attachments.cjs'), str(root / 'pst-job.json')]
                else:
                    commands.pop('pst-reader', None)
                downloads = root / 'download-status.json'
                if not downloads.is_file() or json.loads(downloads.read_bytes()).get('state') != 'complete':
                    commands['downloads'] = [sys.executable, str(root / 'code' / 'fetch_sources.py'), '--root', str(root)]
                else:
                    commands.pop('downloads', None)
                for name, command in commands.items():
                    previous = children.get(name)
                    if previous and (previous[0].poll() is None or time.monotonic() - previous[1] < 60):
                        continue
                    process = start_worker(root, name, command)
                    children[name] = (process, time.monotonic())
                for name, count in parallel.items():
                    lanes = []
                    for lane in range(count):
                        path = roots[name] / ('inspect-status-' + str(lane) + '.json')
                        if path.exists():
                            lanes.append(dict(json.loads(path.read_bytes()), lane=lane))
                    running = [value for value in lanes if value['state'] == 'running']
                    latest = max(running or lanes, key=lambda value: value['updated_utc']) if lanes else {'state': 'starting'}
                    write_json(roots[name] / 'inspect-status.json', dict(latest, parallel_workers=lanes))
                if time.monotonic() - last_statistics >= config.get('statistics_interval_seconds', 1800):
                    statistics = collect_statistics(databases, root)
                    last_statistics = time.monotonic()
                enough_space = (shutil.disk_usage(config['database_directory']).free >= config.get('database_reserve_bytes', SSD_RESERVE_BYTES)
                                and shutil.disk_usage(root).free >= config['reserve_bytes'] + 1024**3)
                status = dict(statistics, updated_utc=utc(), state='running' if enough_space else 'waiting_for_space',
                              pst_full_retention_complete=reader_finished, fuse_profile='maximum-cleaning',
                              temporary_directory=str(root / 'temp'),
                              paused_corpora=[name for name, child in roots.items()
                                              if (child / 'pause.request').exists()],
                              conversion_paused=(root / 'conversion.pause.request').exists(),
                              databases={name: str(db.database_path) for name, db in databases.items()},
                              workers={name: {'pid': process.pid, 'running': process.poll() is None,
                                              'exit_code': process.poll()} for name, (process, _) in children.items()})
                write_json(root / 'status.json', status)
                finished = True
                for name, database in databases.items():
                    feeds = {row['name']: row['complete'] for row in database.db.execute('SELECT * FROM feeds')}
                    if not queue_finished(database, reader_finished, requirements[name], feeds):
                        finished = False
                        break
                if finished:
                    status.update(collect_statistics(databases, root))
                    status['state'] = 'complete_with_open_cases' if status['open_issues'] else 'complete'
                    write_json(root / 'status.json', status)
                    (root / 'stop.request').write_text('Corpus queues finished. Review the recorded open cases.\n')
                    break
                time.sleep(10)
    finally:
        for database in databases.values():
            database.db.close()
