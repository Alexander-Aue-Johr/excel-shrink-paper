"""Download the published FUSE archive and start its frozen maximum-only workflow."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import time
import urllib.request

from .enron.pst_probe import save, now

URL = 'https://zenodo.org/api/records/581678/files/fuse.zip/content'
SIZE = 9392830118
MD5 = '13e955c44f0b77d1c36088c0bbb3366d'


def checksums(path):
    md5, sha = hashlib.md5(), hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            md5.update(block)
            sha.update(block)
    return md5.hexdigest(), sha.hexdigest()


def run(args):
    root = Path(args.output).resolve()
    (root / 'raw').mkdir(parents=True, exist_ok=True)
    target = root / 'raw/fuse.zip'
    pending = target.with_suffix('.zip.partial')
    status = {'state': 'running', 'stage': 'downloading', 'started_utc': now(), 'expected_bytes': SIZE,
              'downloaded_bytes': 0, 'profile': 'maximum-cleaning', 'needle_remaining_xml_pct': 0.001}
    save(root / 'status.json', status)
    try:
        with urllib.request.urlopen('https://zenodo.org/api/records/581678', timeout=60) as response:
            metadata = json.load(response)
        row = next(row for row in metadata['files'] if row['key'] == 'fuse.zip')
        if row['size'] != SIZE or row['checksum'] != 'md5:' + MD5:
            raise ValueError('Published FUSE source differs from the reviewed archive')
        save(root / 'source-metadata.json', metadata)
        save(root / 'source.json', {'record': 581678, 'url': URL, 'bytes': SIZE, 'md5': MD5,
                                   'expected_unique_spreadsheets': 249376, 'profile': 'maximum-cleaning'})
        if target.exists():
            md5, sha = checksums(target)
            if target.stat().st_size != SIZE or md5 != MD5:
                raise ValueError('Existing FUSE archive failed verification; preserved for review')
        else:
            for attempt in range(8):
                try:
                    offset = pending.stat().st_size if pending.exists() else 0
                    if offset > SIZE:
                        raise ValueError('Partial archive exceeds source size')
                    if offset < SIZE:
                        request = urllib.request.Request(URL, headers={'Range': 'bytes=%d-' % offset} if offset else {})
                        with urllib.request.urlopen(request, timeout=90) as response:
                            if offset and response.status == 206:
                                if not response.headers.get('Content-Range', '').startswith('bytes %d-' % offset):
                                    raise ValueError('Unexpected resumed download range')
                                mode = 'ab'
                            elif response.status == 200:
                                offset, mode = 0, 'wb'
                            else:
                                raise ValueError('Unexpected source response')
                            last, written = 0, offset
                            with pending.open(mode) as stream:
                                while True:
                                    block = response.read(1024 * 1024)
                                    if not block:
                                        break
                                    stream.write(block)
                                    written += len(block)
                                    if written > SIZE:
                                        raise ValueError('Archive response exceeds published size')
                                    if time.monotonic() - last > 3:
                                        status.update(downloaded_bytes=written, updated_utc=now())
                                        save(root / 'status.json', status)
                                        last = time.monotonic()
                    md5, sha = checksums(pending)
                    if pending.stat().st_size != SIZE or md5 != MD5:
                        raise ValueError('FUSE archive size or MD5 mismatch')
                    pending.replace(target)
                    break
                except Exception:
                    if attempt == 7:
                        raise
                    time.sleep(min(60, 2 ** (attempt + 1)))
        save(root / 'download-complete.json', {'state': 'complete', 'bytes': SIZE, 'md5': md5, 'sha256': sha,
                                              'verified_utc': now()})
        status.update(stage='waiting_for_workflow', downloaded_bytes=SIZE, updated_utc=now())
        save(root / 'status.json', status)
        while not Path(args.ready).exists():
            time.sleep(10)
        status.update(stage='processing', updated_utc=now())
        save(root / 'status.json', status)
        with (root / 'workflow.stdout.log').open('w') as stdout, (root / 'workflow.stderr.log').open('w') as stderr:
            subprocess.run([sys.executable, str(Path(args.after).resolve()), '--output', str(root)],
                           stdout=stdout, stderr=stderr, check=True)
        status.update(state='complete', stage='finished', finished_utc=now())
        save(root / 'status.json', status)
    except Exception as exc:
        status.update(state='failed', error=str(exc), updated_utc=now())
        save(root / 'status.json', status)
        raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    parser.add_argument('--after', required=True)
    parser.add_argument('--ready', required=True)
    run(parser.parse_args())
