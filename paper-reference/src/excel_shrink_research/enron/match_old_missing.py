"""Match the 146 unresolved old Enron originals against saved Excel embeddings.

File equality, Workbook/Book stream equality and complete storage fingerprints
are separate evidence levels. A negative result does not close archive gaps.
"""
import argparse
import collections
import csv
import io
import json
from pathlib import Path
import zipfile

import olefile
from excel_shrink_research.enron import embedded_workbooks as e


def match(args):
    base, output = Path(args.root), Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    inputs = [base / 'comparison/old-not-found.csv']
    by_file, by_stream = collections.defaultdict(list), collections.defaultdict(list)
    objects = []
    for folder in ('embedded-workbooks-v2', 'container-embeddings'):
        root = base / folder
        inputs.append(root / 'comparison.json')
        aliases = json.loads((root / 'source-documents.json').read_text())
        for row in json.loads((root / 'comparison.json').read_text()):
            sha = row.get('document_sha256', row.get('source_sha256'))
            candidate = {'extraction': folder, 'path': str(root / row['path']),
                         'object_path': row['object_path'], 'source_sha256': sha,
                         'source_occurrences': aliases.get(sha, []),
                         'file_sha256': row['file_sha256'], 'workbook_sha256': row.get('workbook_sha256'),
                         'storage_sha256': row.get('storage_sha256')}
            objects.append(candidate)
            by_file[candidate['file_sha256']].append(candidate)
            if candidate['workbook_sha256']:
                by_stream[candidate['workbook_sha256']].append(candidate)
    inputs += [base / 'embedded-workbooks-v2/recovery/report.json', Path(args.old_zip)]
    recovery = {r['member']: r for r in json.loads(inputs[-2].read_text())}
    with inputs[0].open(newline='', encoding='utf-8-sig') as stream:
        missing = list(csv.DictReader(stream))
    rows, evidence = [], []
    checked_candidates = set()

    def verify(candidate):
        path = candidate['path']
        if path not in checked_candidates:
            if e.digest(Path(path).read_bytes()) != candidate['file_sha256']:
                raise ValueError('Extracted candidate checksum changed: ' + path)
            checked_candidates.add(path)

    with zipfile.ZipFile(args.old_zip) as archive:
        infos = archive.infolist()
        for source in missing:
            names, indices = json.loads(source['member_chain']), json.loads(source['index_chain'])
            if len(indices) != 1:
                raise ValueError('Unexpected old-corpus member chain')
            info = infos[indices[0]]
            if info.filename != names[0]:
                raise ValueError('Old source mapping changed')
            data = archive.read(info)
            if e.digest(data) != source['sha256']:
                raise ValueError('Old source checksum changed')
            row = {'old_member': info.filename, 'old_sha256': source['sha256'], 'bytes': len(data),
                   'previous_kind': source['kind'], 'excel_streams_checked': 0,
                   'match_count': 0, 'comparison': '', 'limitation': ''}
            stream_hashes = []
            try:
                for candidate in by_file.get(source['sha256'], []):
                    verify(candidate)
                    evidence.append({'old_member': info.filename, 'old_object_path': [],
                                     'evidence': 'identical_file_bytes', **candidate})
                if data.startswith(bytes.fromhex('d0cf11e0a1b11ae1')):
                    with olefile.OleFileIO(data) as ole:
                        root = e.StorageReader(ole, ole.root)

                        def compare_storage(storage, path):
                            h = e.workbook_hash(storage)
                            if not h:
                                return
                            stream_hashes.append(h)
                            candidates = by_stream.get(h, [])
                            tree = e.fingerprint(storage) if candidates else None
                            for candidate in candidates:
                                verify(candidate)
                                level = ('identical_named_streams_and_storage_classes'
                                         if candidate['storage_sha256'] == tree else 'identical_workbook_stream')
                                evidence.append({'old_member': info.filename, 'old_object_path': list(path),
                                                 'evidence': level, **candidate})

                        compare_storage(root, [])
                        for path, child in e.embedded(root):
                            compare_storage(child, path)
                        child = None
                elif info.filename in recovery and recovery[info.filename].get('complete'):
                    stream_hashes.append(recovery[info.filename]['workbook_sha256'])
            except Exception as exc:
                saved = recovery.get(info.filename, {})
                if saved.get('complete'):
                    h = saved['workbook_sha256']
                    stream_hashes.append(h)
                    for candidate in by_stream.get(h, []):
                        verify(candidate)
                        evidence.append({'old_member': info.filename, 'old_object_path': [],
                                         'evidence': 'identical_recovered_workbook_stream', **candidate})
                    row['limitation'] = 'Only the previously verified recovered Workbook stream was comparable.'
                else:
                    row['limitation'] = str(exc)
            found = [item for item in evidence if item['old_member'] == info.filename]
            row.update(excel_streams_checked=len(stream_hashes), workbook_sha256=stream_hashes,
                       match_count=len(found), matches=[{'path': item['path'], 'evidence': item['evidence'],
                                                       'object_path': item['object_path']} for item in found])
            row['comparison'] = ('match_found' if found else 'incomplete_read' if row['limitation'] and not stream_hashes
                                 else 'no_workbook_stream_match' if stream_hashes else 'no_file_match_no_workbook_stream')
            rows.append(row)
    e.write_json(output / 'results.json', rows)
    e.write_json(output / 'matches.json', evidence)
    with (output / 'results.csv').open('w', newline='', encoding='utf-8') as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows({k: json.dumps(v) if isinstance(v, list) else v for k, v in row.items()} for row in rows)
    recovered_checks = []
    for index, (name, saved) in enumerate(recovery.items(), 1):
        check = {'old_member': name, 'in_146_search_list': any(row['old_member'] == name for row in rows),
                 'complete_workbook_stream_recovered': saved.get('complete', False)}
        if saved.get('complete'):
            payload = base / 'embedded-workbooks-v2/recovery' / ('recovered-%d-Workbook.bin' % index)
            if e.digest(payload.read_bytes()) != saved['workbook_sha256']:
                raise ValueError('Recovered Workbook stream checksum changed')
            matches = by_stream.get(saved['workbook_sha256'], [])
            for candidate in matches:
                verify(candidate)
            check.update(workbook_sha256=saved['workbook_sha256'],
                         matched_extracted_paths=[candidate['path'] for candidate in matches])
        else:
            check['limitation'] = 'Incomplete Workbook stream; no complete-stream equality claim is possible.'
        recovered_checks.append(check)
    e.write_json(output / 'recovery-checks.json', recovered_checks)
    summary = {'old_originals': len(rows), 'extracted_objects': len(objects),
               'unique_extracted_workbook_streams': len(by_stream),
               'outcomes': dict(collections.Counter(row['comparison'] for row in rows)),
               'old_files_with_excel_streams': sum(row['excel_streams_checked'] > 0 for row in rows),
               'old_excel_streams_checked': sum(row['excel_streams_checked'] for row in rows),
               'recovery_exceptions_in_search_list': [name for name in recovery if any(row['old_member'] == name for row in rows)],
               'matched_old_originals': sum(row['match_count'] > 0 for row in rows),
               'recovered_stream_checks': recovered_checks,
               'inputs': [{'path': str(p), 'sha256': e.digest(p.read_bytes())} for p in inputs[:-1]],
               'scope': 'Exact file bytes and Book/Workbook stream bytes against saved embeddings only. No semantic comparison or additional archive scanning. Non-OLE inputs without a Workbook stream received file-byte comparison only. Prior scan gaps remain open.'}
    e.write_json(output / 'summary.json', summary)
    print(json.dumps(summary, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', default='F:/corpora/enron-v2')
    parser.add_argument('--old-zip', default='F:/corpora/enron.zip')
    parser.add_argument('--output', required=True)
    match(parser.parse_args())
