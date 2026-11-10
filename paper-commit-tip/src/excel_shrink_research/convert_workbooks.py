"""Convert a workbook directory or ZIP to verified XLSX copies with source records."""
import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import shutil
import uuid
import zipfile

from .enron.common import RunLock, write_json, utc_now
from .enron.excel import convert_file, preflight

SUFFIXES = {'.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.xltm', '.xlw', '.ods', '.csv', '.tsv', '.slk'}


def safe_name(name):
    path = PurePosixPath(name.replace('\\', '/'))
    if path.is_absolute() or '..' in path.parts or ':' in name:
        raise ValueError('Unsafe input member: ' + name)
    return path


def digest(path):
    value = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            value.update(block)
    return value.hexdigest()


def run(args):
    source, output = args.input.resolve(), args.output.resolve()
    if source.is_dir() and (output == source or source in output.parents):
        raise ValueError('Keep the output outside the source directory.')
    if args.stage_timeout <= 0 or args.file_timeout <= 0:
        raise ValueError('Timeouts must be positive.')
    preflight()
    selected = set(json.loads(args.members_file.read_text(encoding='utf-8-sig'))) if args.members_file else None
    package = Path(__file__).parent
    code = {str(p.relative_to(package)): digest(p) for p in [Path(__file__), package / 'enron/common.py', package / 'enron/excel.py',
            package / 'enron/excel_worker.ps1', package / 'enron/OwnedExcel.cs']}
    code_hash = hashlib.sha256(json.dumps(code, sort_keys=True).encode()).hexdigest()
    with RunLock(output):
        write_json(output / 'code.json', {'sha256': code_hash, 'files': code})
        archive = zipfile.ZipFile(source) if source.is_file() else None
        try:
            names = ([item.filename for item in archive.infolist() if not item.is_dir()] if archive else
                     [p.relative_to(source).as_posix() for p in source.rglob('*') if p.is_file()])
            names = [n for n in names if safe_name(n).suffix.lower() in SUFFIXES and (selected is None or n in selected)]
            if len(names) != len(set(names)):
                raise ValueError('Duplicate archive member names require explicit source disambiguation.')
            if selected is not None and set(names) != selected:
                raise ValueError('Some selected workbook members are missing or unsupported.')
            if not names:
                raise ValueError('No workbook inputs selected.')
            completed = []
            for name in names:
                if (output / 'stop.request').exists():
                    write_json(output / 'status.json', {'state': 'stopped', 'checked': len(completed), 'total': len(names)})
                    return
                identity = hashlib.sha256(name.encode()).hexdigest()[:24]
                state = output / 'records' / (identity + '.json')
                directory = output / 'attempts' / identity / uuid.uuid4().hex[:8]
                directory.mkdir(parents=True, exist_ok=True)
                local = directory / ('source' + safe_name(name).suffix)
                if archive:
                    with archive.open(name) as src, local.open('wb') as dst:
                        shutil.copyfileobj(src, dst, length=1024 * 1024)
                else:
                    shutil.copyfile(source.joinpath(*safe_name(name).parts), local)
                source_hash = digest(local)
                previous = json.loads(state.read_text()) if state.exists() else None
                if (previous and previous.get('status') == 'accepted' and previous.get('code_sha256') == code_hash
                        and previous.get('source_sha256') == source_hash and Path(previous['output']).is_file()
                        and digest(previous['output']) == previous['output_sha256']):
                    result = previous
                else:
                    target = output / 'xlsx' / identity / safe_name(name).with_suffix('.xlsx').name
                    if target.exists():
                        archived = output / 'superseded' / identity / (digest(target) + '.xlsx')
                        archived.parent.mkdir(parents=True, exist_ok=True)
                        target.replace(archived)
                    result = convert_file(local, target, directory / 'excel', args.stage_timeout, args.file_timeout)
                    result.update(member=name, source=str(source), code_sha256=code_hash, updated_utc=utc_now())
                    if result['status'] == 'accepted':
                        result['output'] = str(target)
                    write_json(state, result)
                local.unlink()
                completed.append(result)
                write_json(output / 'status.json', {'state': 'running', 'checked': len(completed), 'total': len(names),
                           'accepted': sum(r['status'] == 'accepted' for r in completed), 'current_member': name})
            # One compact manifest is sufficient for reuse and a future corpus release.
            with (output / 'conversions.jsonl').open('w', encoding='utf-8') as stream:
                for result in completed:
                    stream.write(json.dumps(result, ensure_ascii=False) + '\n')
            status = {'state': 'complete', 'checked': len(completed), 'total': len(names),
                      'accepted': sum(r['status'] == 'accepted' for r in completed), 'updated_utc': utc_now()}
            write_json(output / 'status.json', status)
            print(json.dumps(status), flush=True)
        finally:
            if archive:
                archive.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--members-file', type=Path, help='JSON array of exact ZIP members or relative directory paths.')
    parser.add_argument('--stage-timeout', type=float, default=120)
    parser.add_argument('--file-timeout', type=float, default=360)
    args = parser.parse_args()
    try:
        run(args)
    except Exception as exc:
        write_json(args.output / 'status.json', {'state': 'failed', 'error': str(exc), 'updated_utc': utc_now()})
        raise


if __name__ == '__main__':
    main()
