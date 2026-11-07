"""Inspect two damaged Enron containers; save only a fully recovered Workbook stream.

No source is modified. A workbook-only reconstruction omits all other streams,
including VBA, and must not replace the source or enter a validated corpus.
"""
import hashlib
import json
from pathlib import Path
import struct
import zipfile


def sha(data):
    return hashlib.sha256(data).hexdigest()


def recover(data):
    if data[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
        raise ValueError('Not a compound file')
    if struct.unpack_from('<HH', data, 26) != (3, 65534):
        raise ValueError('Only little-endian version 3 compound files are supported')
    sector_size = 1 << struct.unpack_from('<H', data, 30)[0]
    if sector_size != 512:
        raise ValueError('Unexpected sector size')
    fat_ids = struct.unpack_from('<109I', data, 76)

    def next_sector(number):
        index, slot = divmod(number, 128)
        if index >= len(fat_ids) or fat_ids[index] >= 0xfffffffa:
            raise ValueError('Required FAT sector is not in the header DIFAT')
        offset = (fat_ids[index] + 1) * 512 + slot * 4
        if offset + 4 > len(data):
            raise ValueError('Missing FAT entry for sector %d' % number)
        return struct.unpack_from('<I', data, offset)[0]

    directories = [struct.unpack_from('<I', data, 48)[0]]

    def entry(index):
        block, slot = divmod(index, 4)
        while len(directories) <= block:
            number = next_sector(directories[-1])
            if number >= 0xfffffffa or number in directories:
                raise ValueError('Invalid directory chain')
            directories.append(number)
        offset = (directories[block] + 1) * 512 + slot * 128
        raw = data[offset:offset + 128]
        if len(raw) != 128:
            raise ValueError('Missing directory entry %d' % index)
        length = struct.unpack_from('<H', raw, 64)[0]
        if not 2 <= length <= 64 or length % 2:
            raise ValueError('Invalid directory name')
        return raw, raw[:length - 2].decode('utf-16-le')

    root, root_name = entry(0)
    if root[66] != 5:
        raise ValueError('Missing root storage')
    pending = [struct.unpack_from('<I', root, 76)[0]]
    visited = set()
    while pending:
        index = pending.pop()
        if index >= 0xfffffffa:
            continue
        if index in visited:
            raise ValueError('Repeated directory entry')
        visited.add(index)
        raw, name = entry(index)
        if raw[66] == 2 and name in ('Workbook', 'Book'):
            break
        left, right = struct.unpack_from('<II', raw, 68)
        pending.extend([right, left])
    else:
        raise ValueError('No root Workbook or Book stream')
    size = struct.unpack_from('<Q', raw, 120)[0]
    if size < 4096:
        raise ValueError('Mini-stream recovery is not supported')
    number = struct.unpack_from('<I', raw, 116)[0]
    chunks, visited = [], set()
    report = {'stream': name, 'declared_bytes': size, 'complete': False}
    try:
        while len(chunks) * 512 < size:
            if number >= 0xfffffffa or number in visited:
                raise ValueError('Invalid or repeated Workbook sector')
            visited.add(number)
            offset = (number + 1) * 512
            chunk = data[offset:offset + 512]
            if len(chunk) != 512:
                raise ValueError('Missing or incomplete Workbook sector %d at byte %d' % (number, offset))
            chunks.append(chunk)
            number = next_sector(number)
        if number != 0xfffffffe:
            raise ValueError('Workbook chain does not end at its declared size')
        report['complete'] = True
    except ValueError as exc:
        report['error'] = str(exc)
    stream = b''.join(chunks)[:size]
    report['read_bytes'] = len(stream)
    if report['complete']:
        report['workbook_sha256'] = sha(stream)
        position, records, endings = 0, 0, 0
        while position + 4 <= len(stream):
            kind, length = struct.unpack_from('<HH', stream, position)
            if position + 4 + length > len(stream):
                break
            position += 4 + length
            records += 1
            endings += kind == 10
        report['biff_record_bytes'] = position
        report['biff_record_count'] = records
        report['biff_eof_records'] = endings
    return report, stream


def main():
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    base = Path('F:/corpora/enron-v2/embedded-workbooks-v2')
    issues = json.loads((base / 'issues.json').read_text())
    comparisons = json.loads((base / 'comparison.json').read_text())
    import olefile
    import pythoncom
    sizes = {}
    for item in comparisons:
        with olefile.OleFileIO(str(base / item['path'])) as ole:
            name = 'Workbook' if ole.exists('Workbook') else 'Book'
            sizes.setdefault(ole.get_size(name), []).append(item['path'])
    results = []
    with zipfile.ZipFile('F:/corpora/enron.zip') as archive:
        for index, issue in enumerate(issues, 1):
            data = archive.read(issue['member'])
            original = output / ('original-%d.xls' % index)
            original.write_bytes(data)
            result, stream = recover(data)
            result.update(member=issue['member'], source_sha256=sha(data), source_bytes=len(data))
            result['embedded_objects_with_same_declared_stream_size'] = len(sizes.get(result['declared_bytes'], []))
            if result['complete']:
                stream_path = output / ('recovered-%d-Workbook.bin' % index)
                stream_path.write_bytes(stream)
                recovered_path = output / ('recovered-%d-WORKBOOK-ONLY.xls' % index)
                storage = pythoncom.StgCreateDocfile(str(recovered_path), 0x1012, 0)
                destination = storage.CreateStream(result['stream'], 0x1012, 0, 0)
                destination.Write(stream)
                destination = None
                storage.Commit(0)
                storage = None
                with olefile.OleFileIO(str(recovered_path)) as ole:
                    if sha(ole.openstream(result['stream']).read()) != result['workbook_sha256']:
                        raise ValueError('Recovered stream verification failed')
                result['reconstructed_file'] = recovered_path.name
                result['warning'] = 'Workbook stream only. Other streams, including VBA and metadata, are omitted. Excel openability and workbook behavior have not been validated.'
            results.append(result)
    (output / 'report.json').write_text(json.dumps(results, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(results, indent=2))


if __name__ == '__main__':
    main()
