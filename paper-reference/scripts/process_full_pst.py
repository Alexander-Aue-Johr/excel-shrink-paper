"""Reuse the verified two-PST probe, then index the remaining pinned mailboxes.

All payloads are hashed. Excel candidates, containers, text and exact targets are
retained separately; other bytes remain recoverable from the unchanged PSTs.
"""
import argparse
import csv
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def prepare(root):
    root = Path(root).resolve()
    code = root / 'code'
    code.mkdir(parents=True, exist_ok=True)
    files = []
    scripts = Path(__file__).resolve().parent
    for name in ('read_pst_attachments.cjs', 'compare_pst_attachments.py', 'process_full_pst.py'):
        target = code / name
        shutil.copy2(scripts / name, target)
        files.append({'path': str(target), 'sha256': digest(target)})
    files.append({'path': str(Path(__file__).resolve()), 'sha256': digest(__file__)})
    receipt = {'state': 'ready', 'files': files,
               'scope': 'Index attachments from the pinned release and compare all 146 unresolved originals. Reuse the two verified probe mailboxes. No Office or conversion.'}
    (root / 'pipeline-ready.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')


def main(args):
    root, previous = Path(args.output).resolve(), Path(args.reuse).resolve()
    if args.prepare_only:
        prepare(root)
        return
    ready = json.loads((root / 'pipeline-ready.json').read_text())
    for item in ready['files']:
        if digest(item['path']) != item['sha256']:
            raise ValueError('Reviewed attachment pipeline file changed: ' + item['path'])
    manifest = json.loads((root / 'sources.json').read_text())
    old_manifest = json.loads((previous / 'sources.json').read_text())
    if manifest['revision'] != old_manifest['revision'] or manifest['dataset'] != old_manifest['dataset']:
        raise ValueError('The previous probe belongs to a different publication')
    expected = {r['name']: r['sha256'] for r in manifest['files']}
    for row in old_manifest['files']:
        if expected.get(row['name']) != row['sha256']:
            raise ValueError('Previous PST source does not match the full publication')
    previous_summary = json.loads((previous / 'summary.json').read_text())
    if previous_summary['source_manifest_sha256'] != digest(previous / 'sources.json') or previous_summary['reader_issues']:
        raise ValueError('Previous reader results are not reusable without review')
    runtime = previous / 'runtime'
    node = runtime / 'node-v22.23.4-win-x64/node.exe'
    reader = runtime / 'reader/node_modules/pst-extractor'
    package = json.loads((reader / 'package.json').read_text())
    if package['version'] != '1.12.0':
        raise ValueError('PST reader version changed')
    runtime_receipt = {'node': str(node), 'node_sha256': digest(node), 'reader': str(reader),
                       'reader_version': package['version'], 'package_lock_sha256': digest(runtime / 'reader/package-lock.json'),
                       'python_dependencies': str(runtime / 'python')}
    (root / 'runtime.json').write_text(json.dumps(runtime_receipt, indent=2) + '\n', encoding='utf-8')
    checkpoint = root / 'reader-checkpoints.json'
    if not checkpoint.exists():
        (root / 'attachments').mkdir(exist_ok=True)
        copied = set()
        journal = root / 'attachments.jsonl'
        with (previous / 'attachments.jsonl').open(encoding='utf-8') as source, journal.open('w', encoding='utf-8') as target:
            for line in source:
                row = json.loads(line)
                if row.get('path') and row['path'] not in copied:
                    source_path = (previous / row['path']).resolve()
                    target_path = (root / row['path']).resolve()
                    source_path.relative_to(previous)
                    target_path.relative_to(root)
                    if digest(source_path) != row['sha256']:
                        raise ValueError('Saved probe attachment checksum mismatch')
                    if not target_path.exists():
                        try:
                            os.link(source_path, target_path)
                        except OSError:
                            shutil.copy2(source_path, target_path)
                    copied.add(row['path'])
                target.write(json.dumps(row) + '\n')
        shutil.copy2(previous / 'reader-issues.jsonl', root / 'reader-issues.jsonl')
        counts = json.loads((previous / 'reader-status.json').read_text())['counts']
        receipt = {'completed_files': [r['name'] for r in old_manifest['files']], 'counts': counts,
                   'journal_bytes': journal.stat().st_size,
                   'issue_journal_bytes': (root / 'reader-issues.jsonl').stat().st_size,
                   'reused_probe_manifest_sha256': digest(previous / 'sources.json')}
        checkpoint.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
    with Path('F:/corpora/enron-v2/comparison/old-not-found.csv').open(newline='', encoding='utf-8-sig') as stream:
        targets = list(csv.DictReader(stream))
    if len(targets) != 146:
        raise ValueError('The 146-file search list changed')
    job = {'output': str(root), 'reader_module': str(reader), 'append_journal': True,
           'retain_payloads': 'candidates_and_targets',
           'known_target_sha256': [r['sha256'] for r in targets],
           'known_target_sizes': sorted({int(r['size']) for r in targets}),
           'pst_files': [str(root / 'raw' / row['name']) for row in manifest['files']]}
    (root / 'reader-job.json').write_text(json.dumps(job, indent=2) + '\n', encoding='utf-8')
    code = root / 'code'
    with (root / 'reader.stdout.log').open('w') as stdout, (root / 'reader.stderr.log').open('w') as stderr:
        subprocess.run([str(node), '--max-old-space-size=2048', str(code / 'read_pst_attachments.cjs'), str(root / 'reader-job.json')],
                       stdout=stdout, stderr=stderr, check=True)
    with (root / 'comparison.stdout.log').open('w') as stdout, (root / 'comparison.stderr.log').open('w') as stderr:
        subprocess.run([sys.executable, str(code / 'compare_pst_attachments.py'), '--output', str(root)],
                       stdout=stdout, stderr=stderr, check=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    parser.add_argument('--reuse', required=True)
    parser.add_argument('--prepare-only', action='store_true', help='Freeze the reviewed scripts and save the ready receipt without reading PSTs')
    main(parser.parse_args())
