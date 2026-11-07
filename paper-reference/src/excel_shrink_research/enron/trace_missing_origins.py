"""Trace unresolved Enron originals to their published extraction and metadata.

PST export suffixes are not treated as message IDs. Same-mailbox, same-size
metadata records are leads only; only verified file bytes establish a match.
"""
import argparse
import collections
import csv
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

from .common import read_json, sha256, utc_now, write_json


def mailbox_identity(filename):
    found = re.fullmatch(r'(.+?)_((?:\d+_)*\d+)\.pst\.(\d+)\.xls', filename, re.I)
    if not found:
        raise ValueError('Unrecognized Nuix export filename: ' + filename)
    person, parts, token = found.groups()
    words = person.split('_')
    surname = words[-2] if words[-1].lower() in ('jr', 'sr') else words[-1]
    return {'custodian_from_filename': person, 'source_pst': person + '_' + parts + '.pst',
            'export_suffix': token, 'edrm_custodian_code': surname.lower() + '-' + words[0][0].lower()}


def document_record(element):
    tags = {child.get('TagName'): child.get('TagValue') for child in element.iter()
            if child.tag.rsplit('}', 1)[-1] == 'Tag'}
    files = []
    for child in element.iter():
        if child.tag.rsplit('}', 1)[-1] == 'File':
            for external in child:
                if external.tag.rsplit('}', 1)[-1] == 'ExternalFile':
                    files.append({'representation': child.get('FileType'), **external.attrib})
    return {'doc_id': element.get('DocID'), 'doc_type': element.get('DocType'), 'mime_type': element.get('MimeType'),
            'original_filename': tags.get('#FileName'), 'tags': tags, 'files': files}


def run(args):
    root, output = Path(args.root), Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    with (root / 'comparison/old-not-found.csv').open(newline='', encoding='utf-8-sig') as stream:
        missing = list(csv.DictReader(stream))
    targets, by_md5, by_size = {}, collections.defaultdict(list), collections.defaultdict(list)
    with zipfile.ZipFile(args.old_zip) as archive:
        infos = archive.infolist()
        readme = next(info for info in infos if info.filename.count('/') == 1 and info.filename.endswith('/README.md'))
        description = archive.read(readme)
        (output / 'upstream-extraction-README.md').write_bytes(description)
        for old in missing:
            info = infos[json.loads(old['index_chain'])[0]]
            if info.filename != json.loads(old['member_chain'])[0]:
                raise ValueError('Old source member mapping changed')
            data = archive.read(info)
            if hashlib.sha256(data).hexdigest() != old['sha256']:
                raise ValueError('Old file checksum mismatch')
            filename = Path(info.filename).name
            identity = mailbox_identity(filename)
            if '/nuix/' not in info.filename:
                raise ValueError('Search list includes a source other than nuix')
            row = {'old_member': info.filename, 'old_sha256': old['sha256'], 'old_md5': hashlib.md5(data).hexdigest(),
                   'bytes': len(data), 'previous_kind': old['kind'], 'source_branch': 'SheetJS Nuix PST extraction',
                   **identity, 'source_description': 'upstream-extraction-README.md#nuix-dataset',
                   'original_attachment_name': None, 'verified_metadata_matches': 0, 'size_leads': 0,
                   'export_suffix_is_message_id': False}
            targets[old['sha256']] = row
            by_md5[row['old_md5']].append(row)
            by_size[(identity['edrm_custodian_code'], len(data))].append(row)
        listing = next(info for info in infos if info.filename.endswith('/artifacts/md5.lst'))
        historical = collections.defaultdict(list)
        for line in archive.read(listing).decode('utf-8', errors='replace').splitlines():
            match = re.fullmatch(r'MD5 \((.*)\) = ([0-9a-fA-F]{32})', line)
            if match and match[2].lower() in by_md5:
                historical[match[2].lower()].append(match[1])
        for row in targets.values():
            row['upstream_md5_catalog_references'] = historical.get(row['old_md5'], [])
    hits, leads, issues, receipts = [], [], [], []
    archives = sorted((root / 'raw').glob('*.zip'))
    status = {'state': 'running', 'old_files': len(targets), 'archives_total': len(archives), 'archives_scanned': 0,
              'documents_scanned': 0, 'verified_matches': 0}
    write_json(output / 'status.json', status)
    for number, path in enumerate(archives, 1):
        code = path.name.removeprefix('edrm-enron-v2_').removesuffix('_xml.zip')
        per_document, last_message = collections.defaultdict(list), None
        try:
            with zipfile.ZipFile(path) as archive:
                metadata = [i for i in archive.infolist() if '/' not in i.filename and i.filename.lower().endswith('.xml')]
                if not metadata:
                    raise ValueError('No top-level XML metadata found')
                for info in metadata:
                    with archive.open(info) as stream:
                        for _, element in ET.iterparse(stream, events=('end',)):
                            tag = element.tag.rsplit('}', 1)[-1]
                            if tag == 'Document':
                                document = document_record(element)
                                status['documents_scanned'] += 1
                                if document['doc_type'] == 'Message':
                                    last_message = document
                                for external in document['files']:
                                    md5 = external.get('Hash', '').lower()
                                    exact = by_md5.get(md5, [])
                                    try:
                                        size = int(external.get('FileSize', '-1'))
                                    except ValueError:
                                        size = -1
                                    weak = by_size.get((code, size), []) if external['representation'] == 'Native' and document['doc_type'] == 'File' else []
                                    for target in exact + [r for r in weak if r not in exact and r['size_leads'] < 10]:
                                        candidate = {'old_sha256': target['old_sha256'], 'archive': path.name,
                                                     'metadata_member': info.filename, 'doc_id': document['doc_id'],
                                                     'original_filename': document['original_filename'], 'mime_type': document['mime_type'],
                                                     'external_file': external, 'parent_doc_id': None, 'parent_relationship_verified': False,
                                                     'evidence': 'same_custodian_and_size_only'}
                                        parent_prefix = document['doc_id'].rsplit('.', 1)[0]
                                        if last_message and parent_prefix == last_message['doc_id']:
                                            candidate['_parent_hint'] = {'doc_id': last_message['doc_id'],
                                                'subject': last_message['tags'].get('#Subject'), 'sent': last_message['tags'].get('#DateSent'),
                                                'from': last_message['tags'].get('#From'),
                                                'native_message': next((f for f in last_message['files'] if f['representation'] == 'Native'), None)}
                                        if target in exact:
                                            member = external.get('FilePath', '').replace('\\', '/').strip('/') + '/' + external['FileName']
                                            payload = archive.read(member.lstrip('/'))
                                            if hashlib.sha256(payload).hexdigest() != target['old_sha256']:
                                                raise ValueError('Metadata MD5 lead did not verify by SHA-256')
                                            candidate['evidence'] = 'identical_bytes_verified_sha256'
                                            hits.append(candidate)
                                            target['verified_metadata_matches'] += 1
                                            if document['original_filename']:
                                                target['original_attachment_name'] = document['original_filename']
                                        else:
                                            leads.append(candidate)
                                            target['size_leads'] += 1
                                        per_document[document['doc_id']].append(candidate)
                                element.clear()
                            elif tag == 'Relationship' and element.get('Type') == 'Attachment':
                                for candidate in per_document.get(element.get('ChildDocId'), []):
                                    candidate['parent_doc_id'] = element.get('ParentDocId')
                                    candidate['parent_relationship_verified'] = True
                                    hint = candidate.get('_parent_hint', {})
                                    if hint.get('doc_id') == candidate['parent_doc_id']:
                                        candidate['email'] = hint
                                element.clear()
                    receipts.append({'archive': path.name, 'metadata_member': info.filename, 'xml_bytes': info.file_size,
                                     'xml_crc32': '%08x' % info.CRC})
        except Exception as exc:
            issues.append({'archive': path.name, 'error': str(exc)})
        status.update(archives_scanned=number, verified_matches=len(hits), candidate_leads=len(leads), issues=len(issues), last_archive=path.name,
                      updated_utc=utc_now())
        write_json(output / 'status.json', status)
    for candidate in hits + leads:
        candidate.pop('_parent_hint', None)
    rows = list(targets.values())
    for row in rows:
        row['result'] = 'verified_metadata_match' if row['verified_metadata_matches'] else 'source_branch_identified_message_unresolved'
    write_json(output / 'origins.json', rows)
    with (output / 'origins.csv').open('w', newline='', encoding='utf-8') as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows({k: json.dumps(v) if isinstance(v, list) else v for k, v in row.items()} for row in rows)
    write_json(output / 'verified-metadata-matches.json', hits)
    write_json(output / 'metadata-leads.json', leads)
    write_json(output / 'issues.json', issues)
    write_json(output / 'metadata-inputs.json', receipts)
    status.update(state='complete_with_issues' if issues else 'complete',
                  mailboxes=len({r['source_pst'] for r in rows}), custodians=len({r['custodian_from_filename'] for r in rows}),
                  files_with_verified_matches=sum(r['verified_metadata_matches'] > 0 for r in rows),
                  files_with_size_leads=sum(r['size_leads'] > 0 for r in rows),
                  custodian_counts=dict(collections.Counter(r['custodian_from_filename'] for r in rows)),
                  old_search_list_sha256=sha256(root / 'comparison/old-not-found.csv'),
                  upstream_readme_sha256=hashlib.sha256(description).hexdigest(),
                  limitation='Mailbox identity comes from Nuix export names and the upstream README. An export suffix is not a message ID. Size leads are capped at ten per old file and do not identify it, even with verified EDRM parent relationships. No semantic comparison or new PST download was performed.')
    write_json(output / 'status.json', status)
    print(json.dumps(status), flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', default='F:/corpora/enron-v2')
    parser.add_argument('--old-zip', default='F:/corpora/enron.zip')
    parser.add_argument('--output', required=True)
    run(parser.parse_args())
