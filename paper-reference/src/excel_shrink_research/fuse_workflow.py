"""Process FUSE with maximum cleaning only and retain the first nonempty needle.

Raw sources stay unchanged. Native XLSX needs no Excel. Legacy conversion waits
for EUSES to release Excel and uses the existing owned-process timeout helper.
"""
import csv
import hashlib
import json
import os
import posixpath
from pathlib import Path
import re
import shutil
import subprocess
import time
import xml.etree.ElementTree as ET
import zipfile

from excel_shrink_research.enron.common import RunLock, read_json, sha256, utc_now, write_json
from excel_shrink_research.enron.excel import convert_file, ExcelInfrastructureError, _is_xlsx

MAXIMUM = ['--remove-textless-default-font-representatives', '--remove-textless-style-representatives',
           '--remove-textless-column-style-representatives', '--remove-unreferenced-whitespace-cells',
           '--inactive-protection-cell-policy', 'remove']


def safe_member(name):
    normalized = name.replace('\\', '/')
    if normalized.startswith('/') or re.match(r'^[A-Za-z]:', normalized) or '..' in normalized.split('/'):
        raise ValueError('Archive member escapes extraction directory: ' + name)


def classify(path):
    with path.open('rb') as stream:
        prefix = stream.read(16)
    if prefix.startswith(b'PK') and _is_xlsx(path):
        return 'native_xlsx'
    if (re.fullmatch(r'[a-fA-F0-9]{40,64}', path.name) or prefix.startswith(bytes.fromhex('d0cf11e0a1b11ae1'))
            or path.suffix.lower() in ('.xls', '.xlsb', '.xlsm', '.xlt', '.xltx', '.xltm', '.xlw', '.csv', '.tsv', '.slk', '.ods')):
        return 'legacy'
    return 'metadata_or_unclassified'


def sheet_reductions(original, output, threshold):
    rows = []
    with zipfile.ZipFile(original) as before, zipfile.ZipFile(output) as after:
        current = {i.filename: i for i in after.infolist()}
        names = {}
        if 'xl/workbook.xml' in before.namelist() and 'xl/_rels/workbook.xml.rels' in before.namelist():
            links = {}
            for link in ET.fromstring(before.read('xl/_rels/workbook.xml.rels')):
                target = link.get('Target', '')
                links[link.get('Id')] = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
            for sheet in ET.fromstring(before.read('xl/workbook.xml')).iter():
                if sheet.tag.rsplit('}', 1)[-1] == 'sheet':
                    relation = next((v for k, v in sheet.attrib.items() if k.endswith('}id')), None)
                    names[links.get(relation)] = sheet.get('name')
        for info in before.infolist():
            if not info.filename.startswith('xl/worksheets/') or not info.filename.endswith('.xml') or '/_rels/' in info.filename:
                continue
            target = current.get(info.filename)
            if not target or not info.file_size or not target.file_size:
                continue
            remaining = 100 * target.file_size / info.file_size
            row = {'part': info.filename, 'sheet_name': names.get(info.filename), 'original_xml_bytes': info.file_size, 'shrunk_xml_bytes': target.file_size,
                   'remaining_xml_pct': remaining, 'reduction_factor': info.file_size / target.file_size,
                   'original_zip_bytes': info.compress_size, 'shrunk_zip_bytes': target.compress_size}
            if remaining <= threshold:
                if target.file_size > 8 * 1024 * 1024:
                    row['meaningful_cells'] = None
                else:
                    count = 0
                    xml = ET.fromstring(after.read(target))
                    for cell in xml.iter():
                        if cell.tag.rsplit('}', 1)[-1] == 'c' and any(
                            item.tag.rsplit('}', 1)[-1] in ('v', 't', 'f') and item.text not in (None, '') for item in cell.iter()):
                            count += 1
                    row['meaningful_cells'] = count
            rows.append(row)
    return sorted(rows, key=lambda row: row['remaining_xml_pct'])


def nested_archives(outer):
    # Numbered volumes form one archive; pass only its first volume to 7-Zip.
    return sorted(path for path in outer.rglob('*') if path.is_file()
                  and (path.name.lower().endswith('.7z') or path.name.lower().endswith('.7z.001')))


def extract(root, config, report):
    marker = root / 'extraction-complete.json'
    if marker.exists():
        return
    outer = root / 'unpacked/outer'
    outer.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(root / 'raw/fuse.zip') as archive:
        for info in archive.infolist():
            safe_member(info.filename)
            if (info.external_attr >> 16) & 0o170000 == 0o120000:
                raise ValueError('Archive contains a symbolic link')
        for info in archive.infolist():
            target = outer / info.filename.replace('\\', '/')
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as source, target.open('wb') as destination:
                    shutil.copyfileobj(source, destination, length=1024 * 1024)
    nested = nested_archives(outer)
    for index, source in enumerate(nested):
        target = root / 'unpacked' / ('payload-%03d' % index)
        target.mkdir(exist_ok=True)
        listing = root / ('archive-%03d-list.txt' % index)
        with listing.open('w', encoding='utf-8') as stream:
            subprocess.run([config['extractor'], 'l', '-slt', '-ba', str(source)], stdout=stream, stderr=subprocess.STDOUT, check=True)
        for line in listing.read_text(encoding='utf-8-sig', errors='replace').splitlines():
            if line.startswith('Path = '):
                safe_member(line[7:])
            if line.startswith(('Symbolic Link = ', 'Hard Link = ')):
                raise ValueError('Nested archive contains a link')
        report(stage='extracting', archive=source.name)
        with (root / ('archive-%03d-extract.log' % index)).open('w') as stream:
            subprocess.run([config['extractor'], 'x', str(source), '-o' + str(target), '-y', '-bb0', '-bsp0'],
                           stdout=stream, stderr=subprocess.STDOUT, check=True)
    write_json(marker, {'nested_7z_archives': len(nested), 'finished_utc': utc_now()})


def inventory(root, report):
    target = root / 'inventory.jsonl'
    if (root / 'inventory-complete.json').exists():
        return
    counts = {'native_xlsx': 0, 'legacy': 0, 'metadata_or_unclassified': 0, 'read_errors': 0}
    with target.open('w', encoding='utf-8') as stream:
        for path in (root / 'unpacked').rglob('*'):
            if not path.is_file() or re.search(r'\.7z(?:\.\d+)?$', path.name, re.IGNORECASE):
                continue
            relative = path.relative_to(root).as_posix()
            row = {'id': hashlib.sha256(relative.encode()).hexdigest()[:24], 'source': relative, 'bytes': path.stat().st_size}
            try:
                row['kind'] = classify(path)
            except Exception as exc:
                row.update(kind='read_errors', error=str(exc))
            counts[row['kind']] += 1
            stream.write(json.dumps(row) + '\n')
            if sum(counts.values()) % 1000 == 0:
                report(stage='inventory', inventory=counts)
    write_json(root / 'inventory-complete.json', {'counts': counts, 'finished_utc': utc_now(),
               'expected_published_spreadsheets': 249376,
               'candidate_count_matches_publication': counts['native_xlsx'] + counts['legacy'] == 249376})


def wait_for_excel(root, report):
    checkpoint = Path('F:/corpora/preparation-queue/checkpoints.json')
    while True:
        if (root / 'stop.request').exists():
            raise RuntimeError('Stop requested; completed FUSE work was preserved')
        done = read_json(checkpoint) if checkpoint.exists() else {}
        if 'D' in done and done['D'].get('counts', {}).get('preparation_remaining') == 0:
            return
        report(stage='waiting_for_euses', message='Native XLSX pass finished. Legacy conversion waits for EUSES.')
        time.sleep(30)


def run(root):
    root = Path(root).resolve()
    config = read_json(root / 'workflow.json')
    for file in config['reviewed_files']:
        if sha256(file['path']) != file['sha256']:
            raise RuntimeError('Pinned workflow file changed: ' + file['path'])
    if config['profile'] != 'maximum-cleaning' or config['flags'] != MAXIMUM:
        raise ValueError('The FUSE workflow must use maximum cleaning only')
    counts = {'processed': 0, 'failed': 0, 'converted': 0, 'needle_candidates': 0}

    def report(**extra):
        write_json(root / 'progress.json', {'state': 'running', 'profile': 'maximum-cleaning',
                   'updated_utc': utc_now(), 'counts': counts, **extra})

    with RunLock(root):
        report(stage='extracting')
        extract(root, config, report)
        inventory(root, report)
        # Read snapshots in one operation; avoid tiny reads on the live data drive.
        rows = [json.loads(line) for line in (root / 'inventory.jsonl').read_text().splitlines()]
        if not any(row['kind'] in ('native_xlsx', 'legacy') for row in rows):
            raise RuntimeError('No spreadsheet candidates inventoried; inspect archive extraction before processing')
        finished = {}
        journal = root / 'results.jsonl'
        if journal.exists():
            for line in journal.read_text().splitlines():
                row = json.loads(line)
                finished[row['id']] = row
                counts['processed' if row['status'] == 'processed' else 'failed'] += 1
        for folder in ('inputs', 'maximum', 'attempts', 'batches', 'reports', 'needle'):
            (root / folder).mkdir(exist_ok=True)
        env = dict(os.environ, RAYON_NUM_THREADS='2')

        def append(row):
            with journal.open('a', encoding='utf-8') as stream:
                stream.write(json.dumps(row) + '\n')
            finished[row['id']] = row
            counts['processed' if row['status'] == 'processed' else 'failed'] += 1

        def batch(pending):
            if not pending:
                return False
            batch_id = hashlib.sha256(''.join(row['id'] for row in pending).encode()).hexdigest()[:20]
            directory = root / 'batches' / batch_id
            directory.mkdir(exist_ok=True)
            needed = []
            for row in pending:
                target = root / 'maximum' / (row['id'] + '.xlsx')
                if not target.exists():
                    stage = directory / (row['id'] + '.xlsx')
                    if not stage.exists():
                        try:
                            os.link(root / 'inputs' / stage.name, stage)
                        except OSError:
                            shutil.copy2(root / 'inputs' / stage.name, stage)
                    needed.append(row)
            if needed:
                command = [config['executable'], str(directory), str(root / 'maximum'), '--batch-report',
                           str(root / 'reports' / (batch_id + '.json')), '--max-active-workbooks', '2', '--writer-workers', '1', *MAXIMUM]
                report(stage='shrinking_maximum', batch_size=len(needed))
                with (root / 'reports' / (batch_id + '.log')).open('w') as log:
                    result = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT, env=env, timeout=1800, cwd=root)
                if result.returncode not in (0, 1):
                    raise RuntimeError('Maximum-cleaning batch stopped unexpectedly; see reports/' + batch_id + '.log')
            found = False
            for row in pending:
                before = root / 'inputs' / (row['id'] + '.xlsx')
                after = root / 'maximum' / before.name
                try:
                    if not after.exists():
                        raise ValueError('No shrunk workbook; inspect batch report')
                    sizes = sheet_reductions(before, after, config['needle_remaining_xml_pct'])
                    result = {**row, 'status': 'processed', 'profile': 'maximum-cleaning', 'input': str(before), 'output': str(after),
                              'baseline_sha256': sha256(before), 'shrunk_sha256': sha256(after),
                              'baseline_bytes': before.stat().st_size, 'shrunk_bytes': after.stat().st_size,
                              'best_sheet': sizes[0] if sizes else None}
                    if sizes:
                        best_path = root / 'best-sheet.json'
                        old = read_json(best_path) if best_path.exists() else None
                        if old is None or sizes[0]['remaining_xml_pct'] < old['best_sheet']['remaining_xml_pct']:
                            write_json(best_path, result)
                    for sheet in sizes:
                        if sheet['remaining_xml_pct'] <= config['needle_remaining_xml_pct']:
                            counts['needle_candidates'] += 1
                            candidate = {**result, 'sheet': sheet, 'visual_validation': 'not_performed'}
                            write_json(root / 'needle' / (row['id'] + '-' + Path(sheet['part']).name + '.json'), candidate)
                            if sheet.get('meaningful_cells', 0):
                                shutil.copy2(before, root / 'needle' / (row['id'] + '-original.xlsx'))
                                shutil.copy2(after, root / 'needle' / (row['id'] + '-maximum.xlsx'))
                                write_json(root / 'needle-found.json', candidate)
                                found = True
                    append(result)
                except Exception as exc:
                    append({**row, 'status': 'failed', 'stage': 'shrinking_or_sheet_comparison', 'error': str(exc)})
            for stage in directory.glob('*.xlsx'):
                stage.unlink()
            report(stage='batch_finished')
            return found

        for kind in ('native_xlsx', 'legacy'):
            if kind == 'legacy':
                wait_for_excel(root, report)
            pending = []
            selected = sorted((r for r in rows if r['kind'] == kind and r['id'] not in finished), key=lambda r: r['bytes'], reverse=True)
            for row in selected:
                if (root / 'stop.request').exists():
                    raise RuntimeError('Stop requested; completed FUSE work was preserved')
                source, baseline = root / row['source'], root / 'inputs' / (row['id'] + '.xlsx')
                row = {**row, 'source_sha256': sha256(source)}
                if not baseline.exists():
                    if kind == 'native_xlsx':
                        shutil.copy2(source, baseline)
                    else:
                        report(stage='converting_legacy', current_source=row['source'])
                        try:
                            result = convert_file(source, baseline, root / 'attempts' / row['id'], stage_timeout=120, total_timeout=360)
                        except ExcelInfrastructureError:
                            raise
                        except Exception as exc:
                            append({**row, 'status': 'failed', 'stage': 'conversion', 'error': str(exc)})
                            continue
                        if result['status'] != 'accepted':
                            append({**row, 'status': 'failed', 'stage': 'conversion', 'conversion': result})
                            continue
                        counts['converted'] += 1
                pending.append(row)
                if len(pending) == (128 if kind == 'native_xlsx' else 16):
                    found = batch(pending)
                    pending = []
                    if found and config['stop_on_needle']:
                        write_json(root / 'progress.json', {'state': 'needle_found', 'profile': 'maximum-cleaning', 'counts': counts,
                                   'message': 'Stopped after the first nonempty worksheet meeting the requested size ratio. Corpus processing is not exhaustive.'})
                        return
            found = batch(pending)
            if found and config['stop_on_needle']:
                write_json(root / 'progress.json', {'state': 'needle_found', 'profile': 'maximum-cleaning', 'counts': counts})
                return
        write_json(root / 'progress.json', {'state': 'complete', 'profile': 'maximum-cleaning', 'counts': counts,
                   'needle_found': (root / 'needle-found.json').exists(), 'inventory': read_json(root / 'inventory-complete.json')})
