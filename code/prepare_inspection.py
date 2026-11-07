"""Backfill scheduling hints and import existing Enron inspection graphs."""

import argparse
import json
from pathlib import Path

from bootstrap import setup
from ledger import Ledger, utc, write_json
from inspection_reuse import import_graph, refresh_sources


def prepare(root, report):
    """Run while workers are stopped; every import is safe to repeat."""
    config = json.loads((root / 'config.json').read_bytes())
    results = {}
    for corpus, directory in config['corpus_roots'].items():
        database = Ledger(Path(directory))
        try:
            # This is a one-time index backfill. Unknown numbered/native files
            # remain ahead of mail. All decisions still use file content.
            database.db.execute('''UPDATE objects SET inspection_priority=CASE
                WHEN candidate=1 THEN 0
                WHEN kind IN ('ole','zip','tar','gzip','bzip2','xz','seven','rtf','xml_container','legacy_excel') THEN 10
                WHEN lower(path) GLOB '*.xls*' OR lower(path) GLOB '*.doc*' OR lower(path) GLOB '*.ppt*'
                  OR lower(path) GLOB '*.zip' OR lower(path) GLOB '*.7z' OR lower(path) GLOB '*.rar'
                  OR lower(path) GLOB '*.rtf' OR lower(path) GLOB '*.xml' THEN 10
                WHEN lower(path) GLOB '*.eml' OR lower(path) GLOB '*.mht*' THEN 30
                WHEN lower(path) GLOB '*.txt' OR lower(path) GLOB '*.text' THEN 40
                WHEN lower(path) GLOB '*.jpg' OR lower(path) GLOB '*.png' OR lower(path) GLOB '*.gif' THEN 50
                ELSE 20 END WHERE state='pending' ''')
        finally:
            database.db.close()
    target = Ledger(Path(config['corpus_roots']['enron-xml']))
    try:
        sources = []
        for name in ('enron-pst', 'enron-old'):
            child = Path(config['corpus_roots'][name])
            sources.append(Path(json.loads((child / 'config.json').read_bytes())['database_path']))
        child_config_path = target.root / 'config.json'
        child_config = json.loads(child_config_path.read_bytes())
        child_config['inspection_reuse_sources'] = [str(path) for path in sources]
        write_json(child_config_path, child_config)
        target.db.execute('DELETE FROM inspection_sources')
        refresh_sources(target, sources)
        for source in sources:
            target.db.execute('ATTACH DATABASE ? AS source', (source.as_uri() + '?mode=ro',))
            roots = [row[0] for row in target.db.execute('''SELECT t.sha FROM objects t JOIN source.objects s ON s.sha=t.sha
                WHERE t.state='pending' AND s.verified=1 AND s.state IN ('complete','review')
                AND (s.kind<>'other' OR s.inspection_revision>=1)''')]
            target.db.execute('DETACH DATABASE source')
            def status(value):
                write_json(report / 'status.json', dict(value, source=str(source), updated_utc=utc()))
            results[source.stem] = import_graph(target, source, roots, status=status, verify_payloads=False)
            write_json(report / 'reuse-results.json', results)
    finally:
        target.db.close()
    for name in ('enron-pst', 'enron-old'):
        database = Ledger(Path(config['corpus_roots'][name]))
        try:
            database.db.execute("UPDATE objects SET state='pending' WHERE kind='other' AND inspection_revision=0 AND state='complete'")
        finally:
            database.db.close()
    write_json(report / 'status.json', {'state': 'complete', 'results': results, 'updated_utc': utc()})


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    setup(args.root)
    prepare(args.root, args.report)
