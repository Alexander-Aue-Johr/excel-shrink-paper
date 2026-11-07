"""Inspect remaining Enron candidates by content without starting Office.

PowerPoint extraction includes all stored revisions, not only objects on current
slides. Source bytes and existing preparation runs are left unchanged.
"""
import argparse
import collections
import csv
from email import policy
from email.parser import BytesParser
import io
import json
from pathlib import Path
import struct
import time
import zipfile
import zlib

import olefile
from excel_shrink_research.enron import embedded_workbooks as e

OLE = bytes.fromhex('d0cf11e0a1b11ae1')
MAX_OBJECT = 256 * 1024 * 1024


def write_json(path, value):
    # Windows scanners can briefly hold the destination between status updates.
    for attempt in range(8):
        try:
            return e.write_json(path, value)
        except PermissionError:
            if attempt == 7:
                raise
            time.sleep(0.05 * (attempt + 1))


def ppt_objects(data, start=0, end=None, depth=0):
    """Yield ExOleObjStg payloads with stream offsets (MS-PPT 2.10.34-36)."""
    if depth > 32:
        raise ValueError('PowerPoint record nesting limit exceeded')
    end = len(data) if end is None else end
    position = start
    while position < end:
        if end - position < 8:
            if any(data[position:end]):
                raise ValueError('Truncated PowerPoint record header')
            break
        flags, kind, size = struct.unpack_from('<HHI', data, position)
        body, next_record = position + 8, position + 8 + size
        if next_record > end:
            raise ValueError('Truncated PowerPoint record at %d' % position)
        if kind == 0x1011:
            instance = flags >> 4
            if instance == 0:
                payload = data[body:next_record]
            elif instance == 1:
                if size < 4:
                    raise ValueError('Missing decompressed object size')
                expected = struct.unpack_from('<I', data, body)[0]
                if expected > MAX_OBJECT:
                    raise ValueError('PowerPoint object exceeds extraction limit')
                decoder = zlib.decompressobj()
                payload = decoder.decompress(data[body + 4:next_record], expected + 1)
                if len(payload) != expected or not decoder.eof or decoder.unconsumed_tail or decoder.unused_data:
                    raise ValueError('PowerPoint object decompression size or framing mismatch')
            else:
                raise ValueError('Unsupported PowerPoint storage instance')
            if len(payload) > MAX_OBJECT:
                raise ValueError('PowerPoint object exceeds extraction limit')
            yield position, payload
        elif flags & 15 == 15:
            yield from ppt_objects(data, body, next_record, depth + 1)
        position = next_record


def source_bytes(raw, row):
    opened = []
    try:
        archive = zipfile.ZipFile(raw / row['archive'])
        opened.append(archive)
        names = json.loads(row['member_chain'])
        chain = json.loads(row['index_chain'])
        for depth, index in enumerate(chain):
            info = archive.infolist()[index]
            if info.filename != names[depth]:
                raise ValueError('Source archive mapping changed')
            data = archive.read(info)
            if depth < len(chain) - 1:
                archive = zipfile.ZipFile(io.BytesIO(data))
                opened.append(archive)
        if e.digest(data) != row['sha256']:
            raise ValueError('Source checksum mismatch')
        return data
    finally:
        for archive in reversed(opened):
            archive.close()


def native_payload(data):
    """Read the length-delimited payload of an Ole10Native package.

    Paths in the wrapper are metadata only; they never become output paths.
    """
    if len(data) < 6:
        raise ValueError('Truncated OLE native package')
    position = 6  # native-data length and format word
    for _ in range(2):  # original filename and source path
        position = data.index(b'\0', position, min(position + 4096, len(data))) + 1
    position += 8
    position = data.index(b'\0', position, min(position + 4096, len(data))) + 1
    if position + 4 > len(data):
        raise ValueError('OLE native package has no stored payload (possibly a link)')
    size = struct.unpack_from('<I', data, position)[0]
    position += 4
    if size > MAX_OBJECT or position + size > len(data):
        raise ValueError('Invalid or truncated OLE native payload')
    return data[position:position + size]


def run(args):
    output = Path(args.output).resolve()
    (output / 'workbooks').mkdir(parents=True, exist_ok=True)
    previous = Path(args.previous)
    old = collections.defaultdict(list)
    for line in (previous / 'old-stream-index.jsonl').read_text().splitlines():
        row = json.loads(line)
        if row.get('workbook_sha256'):
            old[row['workbook_sha256']].append(row['member'])
    word = collections.defaultdict(list)
    for row in json.loads((previous / 'comparison.json').read_text()):
        word[row['workbook_sha256']].append(row['path'])
    sources = collections.defaultdict(list)
    with open(args.candidates, newline='', encoding='utf-8-sig') as stream:
        for row in csv.DictReader(stream):
            sources[row['sha256']].append(row)
    # One document may occur under both .doc and a numbered filename.
    for sha in json.loads((previous / 'source-documents.json').read_text()):
        sources.pop(sha, None)
    if getattr(args, 'only_sha', None):
        sources = {sha: rows for sha, rows in sources.items() if sha in args.only_sha}
    write_json(output / 'source-documents.json', sources)
    inventory, extracted, issues = [], [], []
    status = {'state': 'running', 'sources': len(sources), 'scanned': 0, 'embedded_workbooks': 0}

    def save_book(storage, source_sha, path):
        identity = source_sha + '-' + e.digest(json.dumps(path).encode())[:20]
        relative = 'workbooks/' + identity + '.xls'
        tree = e.extract(storage, output / relative)
        h = e.workbook_hash(storage)
        extracted.append({'source_sha256': source_sha, 'object_path': path, 'path': relative,
                          'file_sha256': e.digest((output / relative).read_bytes()),
                          'storage_sha256': tree, 'workbook_sha256': h,
                          'old_workbook_stream_matches': old.get(h, []),
                          'previous_word_stream_matches': word.get(h, [])})

    def inspect(data, source_sha, path, depth=0):
        if depth > 8:
            raise ValueError('Embedded object nesting limit exceeded')
        if data.startswith(OLE):
            with olefile.OleFileIO(data) as ole:
                storage = e.StorageReader(ole, ole.root)
                names = {item[0] for item in e.entries(storage)}
                category = ('word' if 'WordDocument' in names else 'powerpoint' if 'PowerPoint Document' in names
                            else 'excel' if e.workbook_hash(storage) else 'other_ole')
                if path and e.workbook_hash(storage):
                    save_book(storage, source_sha, path)
                for child_path, child in e.embedded(storage):
                    save_book(child, source_sha, path + list(child_path))
                child = None
                for stream_path in ole.listdir():
                    name = stream_path[-1]
                    if name == 'PowerPoint Document':
                        for offset, payload in ppt_objects(ole.openstream(stream_path).read()):
                            try:
                                inspect(payload, source_sha, path + stream_path + ['record@%d' % offset], depth + 1)
                            except Exception as exc:
                                issues.append({'source_sha256': source_sha, 'object_path': path + stream_path,
                                               'record_offset': offset, 'error': str(exc)})
                    elif name == 'Package':
                        inspect(ole.openstream(stream_path).read(), source_sha, path + stream_path, depth + 1)
                    elif name == '\x01Ole10Native':
                        inspect(native_payload(ole.openstream(stream_path).read()), source_sha,
                                path + stream_path, depth + 1)
                return category
        if data.startswith(b'PK'):
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                names = archive.namelist()
                if 'xl/workbook.xml' in names:
                    if path:
                        relative = 'workbooks/' + source_sha + '-' + e.digest(json.dumps(path).encode())[:20] + '.xlsx'
                        (output / relative).write_bytes(data)
                        extracted.append({'source_sha256': source_sha, 'object_path': path, 'path': relative,
                                          'file_sha256': e.digest(data), 'workbook_sha256': None,
                                          'comparison': 'OOXML payload; legacy Workbook-stream comparison does not apply'})
                    return 'excel_ooxml'
            return 'other_zip'
        prefix = data[:65536].lstrip().lower()
        if prefix.startswith((b'<!doctype html', b'<html')) or b'<html' in prefix[:2000]:
            return 'html'
        message = BytesParser(policy=policy.default).parsebytes(data)
        if message.is_multipart():
            for index, part in enumerate(message.walk()):
                if part.is_multipart():
                    continue
                payload = part.get_payload(decode=True)
                if payload:
                    inspect(payload, source_sha, path + ['mime-part-%d' % index], depth + 1)
            return 'mime_message'
        return 'other_non_ole'

    for number, (sha, aliases) in enumerate(sources.items(), 1):
        before = len(extracted)
        row = aliases[0]
        try:
            category = inspect(source_bytes(Path(args.raw), row), sha, [])
            inventory.append({'source_sha256': sha, 'source_name': json.loads(row['member_chain'])[-1],
                              'extension': Path(json.loads(row['member_chain'])[-1]).suffix.lower(),
                              'content_type': category, 'embedded_workbooks': len(extracted) - before})
        except Exception as exc:
            issues.append({'source_sha256': sha, 'error': str(exc)})
        status.update(scanned=number, embedded_workbooks=len(extracted), issues=len(issues))
        write_json(output / 'status.json', status)
    write_json(output / 'inventory.json', inventory)
    write_json(output / 'comparison.json', extracted)
    write_json(output / 'issues.json', issues)
    for name, rows in (('inventory', inventory), ('comparison', extracted)):
        columns = sorted({key for row in rows for key in row})
        with (output / (name + '.csv')).open('w', newline='', encoding='utf-8') as stream:
            writer = csv.DictWriter(stream, fieldnames=columns)
            writer.writeheader()
            writer.writerows({key: json.dumps(value) if isinstance(value, list) else value
                             for key, value in row.items()} for row in rows)
    hashes = {row['workbook_sha256'] for row in extracted if row.get('workbook_sha256')}
    status.update(state='complete_with_issues' if issues else 'complete',
                  content_types=dict(collections.Counter(row['content_type'] for row in inventory)),
                  unique_workbook_streams=len(hashes),
                  unique_streams_matching_old=sum(h in old for h in hashes),
                  unique_streams_matching_word=sum(h in word for h in hashes),
                  unique_streams_matching_neither=sum(h not in old and h not in word for h in hashes),
                  scope='All remaining unique new-only candidates, by content. PowerPoint objects include stored historical revisions. Comparison reuses the previous old-corpus index with its documented two damaged-file exceptions. Parse failures and unstored native links are reported in issues.json.')
    write_json(output / 'status.json', status)
    print(json.dumps(status), flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--candidates', required=True)
    parser.add_argument('--raw', required=True)
    parser.add_argument('--previous', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--only-sha', action='append', help='Inspect only this source SHA-256; repeat as needed')
    run(parser.parse_args())
