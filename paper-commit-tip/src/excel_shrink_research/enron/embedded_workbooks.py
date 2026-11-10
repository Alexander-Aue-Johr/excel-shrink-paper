"""Extract Word-embedded Excel storages and compare unchanged streams with Enron.

Requires Windows, pywin32 311 and olefile 0.47. Office is never started. Output files are
reconstructed OLE containers; their SHA-256 is not a content-equivalence test.
"""
import argparse
import csv
import hashlib
import io
import json
from pathlib import Path
import tempfile
import zipfile

import pythoncom
import olefile
import pywintypes

READ = 0x20  # STGM_READ | STGM_SHARE_DENY_WRITE
CREATE = 0x1012  # STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE


class StorageReader:
    """Read streams without asking Windows to convert directory FILETIME values.

    The small IStorage-compatible surface keeps hashing shared with native test
    fixtures. Only names, types, lengths, classes and stream bytes are consumed.
    """

    def __init__(self, ole, node, path=()):
        self.ole = ole
        self.node = node
        self.path = path

    def EnumElements(self):
        return [(child.name, child.entry_type, child.size) for child in self.node.kids]

    def Stat(self):
        clsid = self.node.clsid or '00000000-0000-0000-0000-000000000000'
        return (None,) * 8 + (pywintypes.IID('{' + clsid.strip('{}') + '}'),)

    def OpenStorage(self, name, reserved, mode):
        child = next(child for child in self.node.kids if child.name == name)
        return StorageReader(self.ole, child, self.path + (name,))

    def OpenStream(self, name, reserved, mode, flags):
        stream = self.ole.openstream(list(self.path + (name,)))
        stream.Read = stream.read
        return stream

    def CopyTo(self, exclusions, names, target):
        target.SetClass(self.Stat()[8])
        for name, kind, size in self.EnumElements():
            if kind == 1:
                destination = target.CreateStorage(name, CREATE, 0, 0)
                self.OpenStorage(name, None, 0).CopyTo((), None, destination)
                destination.Commit(0)
                destination = None
            elif kind == 2:
                source = self.OpenStream(name, None, 0, 0)
                destination = target.CreateStream(name, CREATE, 0, 0)
                while True:
                    block = source.Read(1024 * 1024)
                    if not block:
                        break
                    destination.Write(block)
                destination = None
            else:
                raise ValueError('Unsupported structured-storage entry type: %s' % kind)

    def close(self):
        self.ole.close()


def open_storage(path):
    ole = olefile.OleFileIO(str(path))
    return StorageReader(ole, ole.root)


def digest(data):
    return hashlib.sha256(data).hexdigest()


def entries(storage):
    return list(storage.EnumElements())


def fingerprint(storage):
    """Hash every named stream and storage class, independent of CFB allocation."""
    records = []

    def visit(node, path):
        records.append([path, 'storage', str(node.Stat()[8])])
        for item in sorted(entries(node), key=lambda x: x[0]):
            name, kind = item[:2]
            child_path = path + [name]
            if kind == 1:
                visit(node.OpenStorage(name, None, 0x10), child_path)
            elif kind == 2:
                stream = node.OpenStream(name, None, 0x10, 0)
                h = hashlib.sha256()
                size = 0
                while True:
                    block = stream.Read(1024 * 1024)
                    if not block:
                        break
                    h.update(block)
                    size += len(block)
                if size != item[2]:
                    raise ValueError('Truncated stream: %r' % child_path)
                records.append([child_path, 'stream', item[2], h.hexdigest()])
            else:
                raise ValueError('Unsupported structured-storage entry type: %s' % kind)

    visit(storage, [])
    return digest(json.dumps(records, ensure_ascii=True, separators=(',', ':')).encode())


def workbook_hash(storage):
    names = {item[0] for item in entries(storage) if item[1] == 2}
    name = next((n for n in ('Workbook', 'Book') if n in names), None)
    if name is None:
        return None
    stream = storage.OpenStream(name, None, 0x10, 0)
    h = hashlib.sha256()
    while True:
        block = stream.Read(1024 * 1024)
        if not block:
            break
        h.update(block)
    return h.hexdigest()


def embedded(storage, path=()):
    for item in entries(storage):
        if item[1] == 1:
            child = storage.OpenStorage(item[0], None, 0x10)
            child_path = path + (item[0],)
            if workbook_hash(child):
                yield child_path, child
            yield from embedded(child, child_path)


def extract(storage, destination):
    before = fingerprint(storage)
    target = pythoncom.StgCreateDocfile(str(destination.resolve()), CREATE, 0)
    storage.CopyTo((), None, target)
    target.SetClass(storage.Stat()[8])
    target.Commit(0)
    target = None
    check = open_storage(destination.resolve())
    try:
        if fingerprint(check) != before:
            raise ValueError('Extracted storage differs from the source streams')
    finally:
        check.close()
    return before


def write_json(path, value):
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(value, indent=2, ensure_ascii=True), encoding='utf-8')
    temporary.replace(path)


def run(args):
    output = Path(args.output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    workbooks = output / 'workbooks'
    workbooks.mkdir(exist_ok=True)
    issues = []
    storage = None
    extracted = []
    counts = {}
    status = {'state': 'running', 'stage': 'extracting', 'counts': counts}
    write_json(output / 'status.json', status)
    with open(args.candidates, newline='', encoding='utf-8-sig') as source:
        rows = list(csv.DictReader(source))
    documents = {}
    for row in rows:
        if json.loads(row['member_chain'])[-1].lower().endswith('.doc'):
            documents.setdefault(row['sha256'], []).append(row)
    counts['word_documents'] = len(documents)
    # Save all source aliases, not only the occurrence chosen for extraction.
    write_json(output / 'source-documents.json', documents)
    with tempfile.TemporaryDirectory(prefix='embedded-', dir=output) as scratch:
        staged = Path(scratch) / 'source.ole'
        for number, (sha, aliases) in enumerate(documents.items(), 1):
            row = aliases[0]
            try:
                archive = zipfile.ZipFile(Path(args.raw) / row['archive'])
                opened = [archive]
                try:
                    chain = json.loads(row['index_chain'])
                    names = json.loads(row['member_chain'])
                    for depth, index in enumerate(chain):
                        info = archive.infolist()[index]
                        if info.filename != names[depth]:
                            raise ValueError('Archive member mapping changed')
                        data = archive.read(info)
                        if depth < len(chain) - 1:
                            archive = zipfile.ZipFile(io.BytesIO(data))
                            opened.append(archive)
                    if digest(data) != sha:
                        raise ValueError('Source document checksum mismatch')
                    staged.write_bytes(data)
                finally:
                    for archive in reversed(opened):
                        archive.close()
                storage = open_storage(staged)
                if 'WordDocument' not in {item[0] for item in entries(storage)}:
                    raise ValueError('Not a Word structured-storage document')
                for object_path, child in embedded(storage):
                    relative = Path('workbooks') / (sha + '-' + digest('/'.join(object_path).encode())[:16] + '.xls')
                    tree_sha = extract(child, output / relative)
                    extracted.append({'document_sha256': sha, 'object_path': list(object_path),
                                      'path': relative.as_posix(), 'workbook_sha256': workbook_hash(child),
                                      'storage_sha256': tree_sha,
                                      'file_sha256': digest((output / relative).read_bytes())})
                child = None
                storage.close()
                storage = None
            except Exception as exc:
                issues.append({'stage': 'extracting', 'document_sha256': sha, 'error': str(exc)})
            finally:
                child = None
                if storage is not None:
                    storage.close()
                storage = None
            counts['documents_scanned'] = number
            counts['embedded_objects'] = len(extracted)
            write_json(output / 'status.json', status)
        write_json(output / 'extracted.json', extracted)
        status['stage'] = 'indexing_old_corpus'
        write_json(output / 'status.json', status)
        old = []
        with open(args.old_index, encoding='utf-8-sig') as source:
            records = [json.loads(line) for line in source if line.strip()]
        targets = {item['workbook_sha256'] for item in extracted}
        with zipfile.ZipFile(args.old_zip) as archive, (output / 'old-stream-index.jsonl').open('w', encoding='utf-8') as journal:
            infos = archive.infolist()
            for number, row in enumerate(records, 1):
                if not row.get('spreadsheet_scope'):
                    continue
                try:
                    info = infos[row['index']]
                    if info.filename != row['member']:
                        raise ValueError('Old archive mapping changed')
                    data = archive.read(info)
                    if digest(data) != row['sha256']:
                        raise ValueError('Old archive checksum mismatch')
                    result = {'member': row['member'], 'file_sha256': row['sha256']}
                    if data.startswith(bytes.fromhex('d0cf11e0a1b11ae1')):
                        staged.write_bytes(data)
                        storage = open_storage(staged)
                        result['workbook_sha256'] = workbook_hash(storage)
                        if result['workbook_sha256'] in targets:
                            result['storage_sha256'] = fingerprint(storage)
                        storage.close()
                        storage = None
                    elif row.get('kind') == 'biff':
                        result['workbook_sha256'] = digest(data)
                    old.append(result)
                    journal.write(json.dumps(result) + '\n')
                except Exception as exc:
                    issues.append({'stage': 'indexing_old_corpus', 'member': row['member'], 'error': str(exc)})
                finally:
                    if storage is not None:
                        storage.close()
                    storage = None
                counts['old_files_scanned'] = len(old)
                if number % 100 == 0:
                    write_json(output / 'status.json', status)
    matches = {}
    for row in old:
        matches.setdefault(row.get('workbook_sha256'), []).append(row)
    for row in extracted:
        found = matches.get(row['workbook_sha256'], [])
        row['old_workbook_stream_matches'] = [item['member'] for item in found]
        row['old_storage_matches'] = [item['member'] for item in found if item.get('storage_sha256') == row['storage_sha256']]
        row['old_file_matches'] = [item['member'] for item in old if item['file_sha256'] == row['file_sha256']]
        row['comparison'] = ('same_storage' if row['old_storage_matches'] else
                             'same_workbook_stream' if found else 'no_workbook_stream_match')
    write_json(output / 'comparison.json', extracted)
    with (output / 'comparison.csv').open('w', newline='', encoding='utf-8') as destination:
        columns = ['document_sha256', 'object_path', 'path', 'workbook_sha256', 'storage_sha256', 'file_sha256',
                   'comparison', 'old_workbook_stream_matches', 'old_storage_matches', 'old_file_matches']
        writer = csv.DictWriter(destination, fieldnames=columns)
        writer.writeheader()
        for row in extracted:
            writer.writerow({key: json.dumps(value) if isinstance(value, list) else value for key, value in row.items()})
    for classification in ('same_storage', 'same_workbook_stream', 'no_workbook_stream_match'):
        subset = [row for row in extracted if row['comparison'] == classification]
        counts[classification + '_objects'] = len(subset)
        counts[classification + '_unique_workbook_streams'] = len({row['workbook_sha256'] for row in subset})
    counts['unique_workbook_streams'] = len({row['workbook_sha256'] for row in extracted})
    counts['issues'] = len(issues)
    write_json(output / 'issues.json', issues)
    status['state'] = 'complete' if not issues else 'complete_with_issues'
    status['stage'] = 'finished'
    status['scope'] = 'Word .doc documents in new-only.csv; embedded Book/Workbook storages'
    status['interpretation'] = 'Same storage means identical named stream bytes and storage classes. Same Workbook stream alone does not establish full workbook equivalence. No match is provisional if scan issues exist. No Office application was started.'
    write_json(output / 'status.json', status)
    print(json.dumps(status), flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--candidates', required=True)
    parser.add_argument('--raw', required=True)
    parser.add_argument('--old-zip', required=True)
    parser.add_argument('--old-index', required=True)
    parser.add_argument('--output', required=True)
    run(parser.parse_args())

