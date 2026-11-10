"""Download the entire pinned public PST release with resumable, verified files.

Complete download means complete for this modified publication, not all Enron
sources. Existing XML archives and earlier probes remain separate.
"""
import argparse
import concurrent.futures
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import threading
import time
import urllib.request

from .pst_probe import DATASET, REVISION, file_hash, now, save


def resumable_download(url, path, expected_size, expected_sha256, progress=None):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        if path.stat().st_size != expected_size or file_hash(path) != expected_sha256:
            raise ValueError('Existing final file failed verification; preserved for review: ' + path.name)
        if progress:
            progress(expected_size, 'verified')
        return
    pending = path.with_suffix(path.suffix + '.partial')
    for attempt in range(8):
        try:
            offset = pending.stat().st_size if pending.exists() else 0
            if offset > expected_size:
                raise ValueError('Partial file exceeds declared source size')
            if offset < expected_size:
                request = urllib.request.Request(url, headers={'Range': 'bytes=%d-' % offset} if offset else {})
                with urllib.request.urlopen(request, timeout=60) as response:
                    status = getattr(response, 'status', 200)
                    if offset and status == 206:
                        content_range = response.headers.get('Content-Range', '')
                        if not content_range.startswith('bytes %d-' % offset) or not content_range.endswith('/%d' % expected_size):
                            raise ValueError('Server returned an unexpected range')
                        mode = 'ab'
                    elif status == 200:
                        offset, mode = 0, 'wb'
                    else:
                        raise ValueError('Unexpected download response: %s' % status)
                    written, last = offset, 0
                    with pending.open(mode) as destination:
                        while True:
                            block = response.read(1024 * 1024)
                            if not block:
                                break
                            if written + len(block) > expected_size:
                                raise ValueError('Response exceeds declared source size')
                            destination.write(block)
                            written += len(block)
                            if progress and time.monotonic() - last >= 2:
                                progress(written, 'downloading')
                                last = time.monotonic()
            if pending.stat().st_size != expected_size:
                raise ValueError('Incomplete PST download')
            if file_hash(pending) != expected_sha256:
                # Keep untrusted bytes, but never resume or publish them as valid.
                pending.rename(pending.with_name(pending.name + '.bad-sha256-' + str(time.time_ns())))
                raise ValueError('PST SHA-256 mismatch; untrusted bytes were quarantined')
            pending.replace(path)
            if progress:
                progress(expected_size, 'verified')
            return
        except Exception:
            if attempt == 7:
                raise
            time.sleep(min(60, 2 ** (attempt + 1)))


def get_manifest():
    url = 'https://huggingface.co/api/datasets/' + DATASET + '/tree/' + REVISION + '/data?limit=1000'
    with urllib.request.urlopen(url, timeout=30) as response:
        entries = json.load(response)
    files = []
    for row in entries:
        if row.get('type') != 'file' or not row.get('path', '').endswith('.pst'):
            continue
        name = Path(row['path']).name
        if row['path'] != 'data/' + name or name in ('.', '..'):
            raise ValueError('Unexpected PST source path')
        expected = row.get('lfs', {}).get('oid', '')
        if len(expected) != 64 or any(c not in '0123456789abcdef' for c in expected):
            raise ValueError('PST source has no valid SHA-256')
        files.append({'name': name, 'bytes': row['size'], 'sha256': expected,
                      'url': 'https://huggingface.co/datasets/' + DATASET + '/resolve/' + REVISION + '/' + row['path']})
    if len(files) != 171 or len({r['name'] for r in files}) != 171 or sum(r['bytes'] for r in files) != 83083881472:
        raise ValueError('Pinned publication does not match the reviewed 171-file inventory')
    return {'dataset': DATASET, 'revision': REVISION, 'files': files, 'total_bytes': 83083881472,
            'target_scope': 'all_146_missing', 'source_warning':
            'This publication removed 1,015 messages and 39,147 attachments. Download completion is not a guarantee of all historical Enron or Nuix content.'}


def run(args):
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=True)
    status = {'state': 'running', 'stage': 'inventory', 'started_utc': now(), 'pid': os.getpid(), 'files': {}}
    save(root / 'status.json', status)
    lock = threading.Lock()
    try:
        manifest = get_manifest()
        old_manifest = root / 'sources.json'
        if old_manifest.exists() and json.loads(old_manifest.read_text()) != manifest:
            raise ValueError('Saved publication manifest differs; no source files were changed')
        save(old_manifest, manifest)
        free = shutil.disk_usage(root).free
        required = sum(row['bytes'] for row in manifest['files'] if not (root / 'raw' / row['name']).exists())
        if free < required + 20 * 1024 ** 3:
            raise ValueError('Insufficient free space for remaining PSTs and a 20 GiB reserve')
        status.update(stage='downloading', total_files=171, total_bytes=manifest['total_bytes'])
        source_root = Path(args.reuse) / 'raw'
        (root / 'raw').mkdir(exist_ok=True)
        for row in manifest['files']:
            previous, target = source_root / row['name'], root / 'raw' / row['name']
            if previous.exists() and not target.exists():
                if previous.stat().st_size != row['bytes'] or file_hash(previous) != row['sha256']:
                    raise ValueError('Prior probe PST failed verification')
                try:
                    os.link(previous, target)
                except OSError:
                    shutil.copy2(previous, target)

        def report(name, written, state):
            with lock:
                status['files'][name] = {'bytes': written, 'state': state}
                status.update(updated_utc=now(), verified_files=sum(r['state'] == 'verified' for r in status['files'].values()),
                              verified_bytes=sum(r['bytes'] for r in status['files'].values() if r['state'] == 'verified'),
                              transferred_or_reused_bytes=sum(r['bytes'] for r in status['files'].values()))
                save(root / 'status.json', status)

        failures = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
            jobs = {pool.submit(resumable_download, row['url'], root / 'raw' / row['name'], row['bytes'], row['sha256'],
                                lambda written, state, name=row['name']: report(name, written, state)): row for row in manifest['files']}
            for future in concurrent.futures.as_completed(jobs):
                try:
                    future.result()
                except Exception as exc:
                    failures.append({'file': jobs[future]['name'], 'error': str(exc)})
        if failures:
            save(root / 'download-errors.json', failures)
            raise RuntimeError('%d PST downloads failed; verified files and resumable partials were preserved' % len(failures))
        save(root / 'download-complete.json', {'state': 'complete', 'files': 171,
             'bytes': manifest['total_bytes'], 'revision': REVISION, 'manifest_sha256': file_hash(root / 'sources.json'), 'finished_utc': now()})
        status.update(stage='waiting_for_attachment_pipeline', updated_utc=now())
        save(root / 'status.json', status)
        while not Path(args.ready).exists():
            time.sleep(15)
        status.update(stage='extracting_and_comparing', updated_utc=now())
        save(root / 'status.json', status)
        with (root / 'attachments.stdout.log').open('w') as stdout, (root / 'attachments.stderr.log').open('w') as stderr:
            subprocess.run([sys.executable, str(Path(args.after).resolve()), '--output', str(root), '--reuse', str(Path(args.reuse).resolve())],
                           stdout=stdout, stderr=stderr, check=True)
        result = json.loads((root / 'summary.json').read_text())
        status.update(state='complete_with_issues' if result.get('reader_issues') or result.get('attachment_ole_issues') else 'complete',
                      stage='finished', finished_utc=now(), result=result)
        save(root / 'status.json', status)
    except Exception as exc:
        status.update(state='failed', error=str(exc), updated_utc=now())
        save(root / 'status.json', status)
        raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    parser.add_argument('--reuse', required=True)
    parser.add_argument('--after', required=True)
    parser.add_argument('--ready', required=True)
    run(parser.parse_args())
