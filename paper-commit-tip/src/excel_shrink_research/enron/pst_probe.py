"""Download two pinned PST sources and run an isolated attachment comparison.

The selected Hugging Face release is modified. A negative result cannot establish
absence from the historical Nuix corpus. Office is never started.
"""
import argparse
import concurrent.futures
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import time
import urllib.request
import zipfile

REVISION = '1c6490b25d3a16b71a070cbbef87d8fe394c3df3'
DATASET = 'intellekthq/enron-ferc-pst'
SOURCES = [('zl_saibi-e_000.pst', 99566592, 'e8f06b8bca0456fdf44d66b13a9a60ac71f0df710647fda5d3072a92f673eadf'),
           ('zl_sanders-r_000.pst', 540427264, '129d75c9cd84bba0eb06d41f1a46dc47ef5ec484656f4b3595aefe20d1987301')]
NODE = 'node-v22.23.4-win-x64'
NODE_SHA256 = 'cbc74b5e0982c9dbb41525fd8fb5dcd32fbee237863f7829072f4639a1ce5f20'


def now():
    return datetime.now(timezone.utc).isoformat()


def save(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix('.tmp')
    for attempt in range(8):
        try:
            temporary.write_text(json.dumps(value, indent=2) + '\n', encoding='utf-8')
            temporary.replace(path)
            return
        except PermissionError:
            if attempt == 7:
                raise
            time.sleep(0.1 * (attempt + 1))


def file_hash(path):
    result = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            result.update(block)
    return result.hexdigest()


def download(url, path, expected, size=None, progress=None):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and (size is None or path.stat().st_size == size) and file_hash(path) == expected:
        if progress:
            progress(path.stat().st_size, 'verified')
        return
    pending = path.with_suffix(path.suffix + '.partial')
    for attempt in range(3):
        try:
            h, written, last = hashlib.sha256(), 0, 0
            with urllib.request.urlopen(url, timeout=60) as response, pending.open('wb') as destination:
                while True:
                    block = response.read(1024 * 1024)
                    if not block:
                        break
                    h.update(block)
                    destination.write(block)
                    written += len(block)
                    if progress and time.monotonic() - last > 2:
                        progress(written, 'downloading')
                        last = time.monotonic()
            if h.hexdigest() != expected or (size is not None and written != size):
                raise ValueError('Source size or SHA-256 mismatch: ' + path.name)
            pending.replace(path)
            if progress:
                progress(written, 'verified')
            return
        except Exception:
            if attempt == 2:
                raise
            time.sleep(2 * (attempt + 1))


def run(args):
    root = Path(args.output).resolve()
    root.mkdir(parents=True, exist_ok=True)
    runtime = root / 'runtime'
    status = {'state': 'running', 'stage': 'downloading', 'started_utc': now(), 'downloads': {}}
    lock = threading.Lock()

    def report(name, size, state):
        with lock:
            status['downloads'][name] = {'bytes': size, 'state': state}
            status['updated_utc'] = now()
            save(root / 'status.json', status)

    try:
        save(root / 'status.json', status)
        manifest = {'dataset': DATASET, 'revision': REVISION, 'reader': 'pst-extractor@1.12.0',
                    'node': NODE, 'node_sha256': NODE_SHA256, 'selected_custodians': ['Eric Saibi', 'Richard Sanders'],
                    'source_warning': 'Modified release; messages and attachments were removed. Negative results do not establish absence from the Nuix PST release.',
                    'files': [{'name': name, 'bytes': size, 'sha256': digest,
                               'url': 'https://huggingface.co/datasets/' + DATASET + '/resolve/' + REVISION + '/data/' + name}
                              for name, size, digest in SOURCES]}
        save(root / 'sources.json', manifest)
        card_url = 'https://huggingface.co/datasets/' + DATASET + '/raw/' + REVISION + '/README.md'
        with urllib.request.urlopen(card_url, timeout=30) as response:
            (root / 'upstream-dataset-card.md').write_bytes(response.read())
        jobs = [(row['url'], root / 'raw' / row['name'], row['sha256'], row['bytes']) for row in manifest['files']]
        jobs.append(('https://nodejs.org/dist/v22.23.4/' + NODE + '.zip', runtime / (NODE + '.zip'), NODE_SHA256, None))
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
            futures = [pool.submit(download, url, path, digest, size,
                                   lambda written, state, name=path.name: report(name, written, state))
                       for url, path, digest, size in jobs]
            for future in futures:
                future.result()
        status['stage'] = 'installing_reader'
        save(root / 'status.json', status)
        node_root = runtime / NODE
        if not (node_root / 'node.exe').exists():
            with zipfile.ZipFile(runtime / (NODE + '.zip')) as archive:
                for info in archive.infolist():
                    (runtime / info.filename).resolve().relative_to(runtime.resolve())
                archive.extractall(runtime)
        env = dict(os.environ)
        env['PATH'] = str(node_root) + os.pathsep + env.get('PATH', '')
        env['npm_config_cache'] = str(runtime / 'npm-cache')
        npm = node_root / 'node_modules/npm/bin/npm-cli.js'
        reader = runtime / 'reader'
        reader.mkdir(parents=True, exist_ok=True)
        with (runtime / 'install.stdout.log').open('w') as stdout, (runtime / 'install.stderr.log').open('w') as stderr:
            subprocess.run([str(node_root / 'node.exe'), str(npm), 'install', '--prefix', str(reader),
                            '--ignore-scripts', '--no-audit', '--no-fund', '--registry=https://registry.npmjs.org',
                            'pst-extractor@1.12.0'], env=env, stdout=stdout, stderr=stderr, check=True, timeout=600)
            subprocess.run([sys.executable, '-m', 'pip', 'install', '--disable-pip-version-check', '--target',
                            str(runtime / 'python'), 'olefile==0.47'], stdout=stdout, stderr=stderr, check=True, timeout=300)
        status['stage'] = 'waiting_for_reader_script'
        save(root / 'status.json', status)
        deadline = time.monotonic() + 1800
        while not Path(args.ready).exists():
            if time.monotonic() > deadline:
                raise RuntimeError('Reader script was not marked ready; downloaded sources were preserved')
            time.sleep(5)
        status['stage'] = 'reading_attachments'
        save(root / 'status.json', status)
        job = {'reader_module': str(reader / 'node_modules/pst-extractor'),
               'output': str(root), 'pst_files': [str(root / 'raw' / name) for name, _, _ in SOURCES]}
        save(root / 'reader-job.json', job)
        with (root / 'reader.stdout.log').open('w') as stdout, (root / 'reader.stderr.log').open('w') as stderr:
            subprocess.run([str(node_root / 'node.exe'), str(Path(args.worker).resolve()), str(root / 'reader-job.json')],
                           env=env, stdout=stdout, stderr=stderr, check=True, timeout=21600)
        status['stage'] = 'comparing_attachments'
        save(root / 'status.json', status)
        env['PYTHONPATH'] = str(runtime / 'python') + os.pathsep + str(Path(__file__).resolve().parents[2])
        with (root / 'comparison.stdout.log').open('w') as stdout, (root / 'comparison.stderr.log').open('w') as stderr:
            subprocess.run([sys.executable, str(Path(args.compare).resolve()), '--output', str(root)],
                           env=env, stdout=stdout, stderr=stderr, check=True, timeout=3600)
        status.update(state='complete', stage='finished', finished_utc=now(),
                      result=json.loads((root / 'summary.json').read_text()))
        save(root / 'status.json', status)
    except Exception as exc:
        status.update(state='failed', error=str(exc), updated_utc=now())
        save(root / 'status.json', status)
        raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    parser.add_argument('--worker', required=True)
    parser.add_argument('--compare', required=True)
    parser.add_argument('--ready', required=True)
    run(parser.parse_args())
