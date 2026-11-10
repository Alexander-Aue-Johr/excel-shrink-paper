"""Match PST attachments against the scoped list of unresolved Nuix originals."""
import argparse
import collections
import csv
import hashlib
import json
from pathlib import Path
import sys
import zipfile


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def workbook_hash(path):
    import olefile
    with Path(path).open('rb') as stream:
        if stream.read(8) != bytes.fromhex('d0cf11e0a1b11ae1'):
            return None
    with olefile.OleFileIO(str(path)) as ole:
        for name in ('Workbook', 'Book'):
            if ole.exists(name) and ole.get_type(name) == 2:
                stream = ole.openstream(name)
                h, size = hashlib.sha256(), 0
                for block in iter(lambda: stream.read(1024 * 1024), b''):
                    size += len(block)
                    h.update(block)
                if size != ole.get_size(name):
                    raise ValueError('Truncated Excel stream')
                return h.hexdigest()
    return None


def compare(root):
    root = Path(root).resolve()
    sys.path.insert(0, str(root / 'runtime/python'))
    if (root / 'runtime.json').exists():
        sys.path.insert(0, json.loads((root / 'runtime.json').read_text())['python_dependencies'])
    import olefile
    if olefile.__version__ != '0.47':
        raise ValueError('Expected the pinned olefile 0.47 reader')
    base = Path('F:/corpora/enron-v2')
    full_release = json.loads((root / 'sources.json').read_text()).get('target_scope') == 'all_146_missing'
    with (base / 'comparison/old-not-found.csv').open(newline='', encoding='utf-8-sig') as stream:
        targets = [row for row in csv.DictReader(stream) if full_release or
                   Path(json.loads(row['member_chain'])[-1]).name.startswith(('eric_saibi_', 'richard_sanders_'))]
    if len(targets) != (146 if full_release else 131):
        raise ValueError('The reviewed search list changed; inspect before continuing')
    old_index = {row['member']: row for row in
                 (json.loads(line) for line in (base / 'embedded-workbooks-v2/old-stream-index.jsonl').read_text().splitlines())}
    by_file, by_stream = collections.defaultdict(list), collections.defaultdict(list)
    outcomes = {}
    with zipfile.ZipFile('F:/corpora/enron.zip') as archive:
        infos = archive.infolist()
        for row in targets:
            member = json.loads(row['member_chain'])[0]
            info = infos[json.loads(row['index_chain'])[0]]
            if info.filename != member or hashlib.sha256(archive.read(info)).hexdigest() != row['sha256']:
                raise ValueError('Old original checksum or mapping changed')
            stream_hash = old_index.get(member, {}).get('workbook_sha256')
            result = {'old_member': member, 'old_sha256': row['sha256'], 'old_kind': row['kind'],
                      'workbook_sha256': stream_hash, 'exact_file_occurrences': 0, 'same_excel_stream_occurrences': 0}
            outcomes[row['sha256']] = result
            by_file[row['sha256']].append(result)
            if stream_hash:
                by_stream[stream_hash].append(result)
    matches, issues, seen = [], [], {}
    attachments = 0
    with (root / 'attachments.jsonl').open(encoding='utf-8') as journal:
        for line in journal:
            row = json.loads(line)
            if row['type'] != 'payload':
                continue
            attachments += 1
            sha = row['sha256']
            if sha not in seen:
                if not row.get('path'):
                    if sha in by_file:
                        raise ValueError('An exact target attachment was not retained')
                    seen[sha] = None
                    continue
                path = (root / row['path']).resolve()
                path.relative_to(root)
                if path.stat().st_size != row['bytes'] or digest(path) != sha:
                    raise ValueError('Extracted attachment checksum changed')
                try:
                    seen[sha] = workbook_hash(path)
                except Exception as exc:
                    seen[sha] = None
                    issues.append({'path': row['path'], 'sha256': sha, 'error': str(exc)})
            candidates = {r['old_sha256']: (r, 'identical_workbook_stream') for r in by_stream.get(seen[sha], [])}
            candidates.update({r['old_sha256']: (r, 'identical_file_bytes') for r in by_file.get(sha, [])})
            for result, evidence in candidates.values():
                result['exact_file_occurrences' if evidence == 'identical_file_bytes' else 'same_excel_stream_occurrences'] += 1
                matches.append({'old_member': result['old_member'], 'old_sha256': result['old_sha256'],
                                'evidence': evidence, 'attachment': row})
    for row in outcomes.values():
        row['result'] = ('identical_file_found' if row['exact_file_occurrences'] else
                         'identical_excel_stream_found' if row['same_excel_stream_occurrences'] else 'no_match_in_selected_modified_psts')
    for name, data in [('matches', matches), ('results', list(outcomes.values())), ('attachment-read-issues', issues)]:
        (root / (name + '.json')).write_text(json.dumps(data, indent=2) + '\n', encoding='utf-8')
    with (root / 'results.csv').open('w', newline='', encoding='utf-8') as stream:
        writer = csv.DictWriter(stream, fieldnames=list(next(iter(outcomes.values()))))
        writer.writeheader()
        writer.writerows(outcomes.values())
    with (root / 'matches.csv').open('w', newline='', encoding='utf-8') as stream:
        columns = ['old_member', 'evidence', 'pst', 'folder_path', 'message_node_id', 'internet_message_id',
                   'subject', 'delivered', 'original_filename', 'attachment_index', 'sha256', 'path']
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        for match in matches:
            attachment = match['attachment']
            record = {key: attachment.get(key) for key in columns}
            record.update(old_member=match['old_member'], evidence=match['evidence'])
            record['folder_path'] = json.dumps(record['folder_path'])
            writer.writerow(record)
    reader = json.loads((root / 'reader-status.json').read_text())
    summary = {'old_originals': len(targets), 'pst_files_read': reader['counts']['pst_files'],
               'attachments_read': attachments, 'unique_attachment_payloads': len(seen),
               'matched_old_originals': sum(r['result'] != 'no_match_in_selected_modified_psts' for r in outcomes.values()),
               'exact_file_matches': sum(r['exact_file_occurrences'] > 0 for r in outcomes.values()),
               'stream_only_matches': sum(r['same_excel_stream_occurrences'] > 0 and not r['exact_file_occurrences'] for r in outcomes.values()),
               'reader_issues': reader['counts']['issues'], 'attachment_ole_issues': len(issues),
               'outcomes': dict(collections.Counter(r['result'] for r in outcomes.values())),
               'source_manifest_sha256': digest(root / 'sources.json'),
               'old_search_list_sha256': digest(base / 'comparison/old-not-found.csv'),
               'old_stream_index_sha256': digest(base / 'embedded-workbooks-v2/old-stream-index.jsonl'),
               'source_download_complete': (root / 'download-complete.json').exists(),
               'limitation': 'The PST source is a modified release, not a verified copy of the historical Nuix PSTs. Workbook-stream equality alone does not establish complete-workbook equivalence. Negative results do not close earlier archive gaps or rule out other source versions.'}
    (root / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(summary))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    compare(parser.parse_args().output)
