#!/usr/bin/env python3
"""Inventory every tracked file at the reviewed production and experiment revisions.

Source files receive a family, role, line count and lexical symbol index. This
is an editorial coverage ledger, not a compiler analysis or correctness audit.
Binary artifacts remain listed by Git object identity without decoding them.
"""
import argparse
import csv
import gzip
import hashlib
import io
import json
import re
import subprocess
from collections import Counter
from pathlib import Path

SOURCE_EXTENSIONS = {
    '.rs', '.py', '.js', '.mjs', '.cjs', '.ts', '.tsx', '.cs', '.bas', '.ps1',
    '.psm1', '.psd1', '.sh', '.html', '.css', '.xaml', '.vb', '.cpp', '.c',
    '.h', '.bat', '.cmd',
}
SYMBOLS = re.compile(
    r'^[ \t]*(?:(?:pub(?:\([^)]*\))?|async|export|default|static|public|private|'
    r'protected|internal|override|virtual|sealed)\s+)*'
    r'(?:fn|struct|enum|trait|class|interface|function\*?|def|Sub|Function)\s+(\w+)',
    re.MULTILINE,
)


def git(repository: Path, *arguments: str) -> bytes:
    return subprocess.check_output(['git', '-C', str(repository), *arguments])


def family(path: str) -> str:
    """Assign all code to a bounded editorial responsibility, including support."""
    if '/vendor/' in path or '/node_modules/' in path:
        return 'third-party'
    if path.startswith('docs/visual-validation/evidence/'):
        return 'portable-evidence-publication'
    if path.startswith('excel-shrink-core/'):
        if '/sheet/' in path:
            return 'cleanup-rules'
        if 'reference_scanner' in path:
            return 'cleanup-dependencies'
        if '/xml' in path:
            return 'byte-parser'
        if 'style' in path or 'shared_string' in path:
            return 'cleanup-context'
        return 'cleanup-model-and-instrumentation'
    if 'experiments/' in path or 'excel-lab/' in path or 'excel-height-check/' in path:
        return 'development-experiments'
    if 'targeted_comparison/conditional_formatting' in path:
        return 'conditional-dependencies'
    if 'targeted_comparison/' in path or path.endswith('targeted_comparison.rs'):
        return 'targeted-selection-and-evidence'
    if '/office_plan/' in path or 'streaming_xml' in path or 'worksheet_stream' in path:
        return 'validation-package-analysis'
    if '/formula_trace/' in path or 'value_stream' in path or 'value_views' in path:
        return 'value-and-formula-attribution'
    if '/regression/office_capture' in path or '/capture_host/' in path:
        return 'native-capture-protocol'
    if '/regression/archive' in path or 'validation_preparation' in path:
        return 'corpus-preparation-and-scheduling'
    if 'workbook_worker' in path or 'excel-automation' in path:
        return 'native-process-lifecycle'
    if 'artifact_cache' in path or 'deltasieve-engine' in path or 'metadata_store' in path:
        return 'evidence-identity-and-reuse'
    if 'validation_report' in path or 'visual_report' in path or 'diagnosis_report' in path:
        return 'portable-evidence-publication'
    if 'office-capture/web/' in path:
        return 'native-capture-protocol'
    if 'control-room/' in path or 'sheet-atlas/' in path or '/preparation-web/' in path:
        return 'visualization-and-user-interface'
    if path.endswith('/pipeline.rs') or path.endswith('/observed_cleanup.rs'):
        return 'cleanup-scheduling-and-output'
    if '/regression/' in path:
        return 'validation-orchestration-and-controls'
    if path.startswith('excel-shrink-cli/'):
        return 'application-and-study-control'
    if path.startswith(('apps/', 'crates/')):
        return 'shared-application-services'
    if path.startswith(('scripts/', 'tools/', '.vscode/')):
        return 'research-build-and-acquisition-tools'
    return 'repository-configuration'


def source_role(path: str, group: str) -> str:
    if group == 'third-party':
        return 'third-party'
    if ('.Designer.' in path or '/dist/' in path
            or path.startswith('docs/visual-validation/evidence/')):
        return 'generated'
    if re.search(r'(^|[/_.-])(tests?|fixtures?)([/_.-]|$)', path):
        return 'test-or-fixture'
    if group == 'development-experiments':
        return 'experiment'
    return 'implementation'


def read_sources(repository: Path, objects: list[str], cache: dict[str, bytes]) -> dict[str, bytes]:
    """Read each distinct source blob once through Git's batch protocol."""
    missing = [identity for identity in objects if identity not in cache]
    if not missing:
        return cache
    result = subprocess.run(
        ['git', '-C', str(repository), 'cat-file', '--batch'],
        input=('\n'.join(missing) + '\n').encode(), stdout=subprocess.PIPE, check=True,
    ).stdout
    stream = io.BytesIO(result)
    for expected in missing:
        identity, kind, size = stream.readline().decode().split()
        if identity != expected or kind != 'blob':
            raise ValueError(f'Unexpected Git batch object {identity}')
        cache[identity] = stream.read(int(size))
        if stream.read(1) != b'\n':
            raise ValueError('Invalid Git batch framing')
    return cache


def inspect_tree(repository: Path, revision: str, label: str, cache: dict, scripts_only=False):
    """Return one row for every tree entry, without silently omitting source files."""
    entries = []
    for record in git(repository, 'ls-tree', '-rz', '-l', revision).split(b'\0'):
        if not record:
            continue
        metadata, raw_path = record.split(b'\t', 1)
        mode, kind, identity, size = metadata.decode().split()
        path = raw_path.decode('utf-8')
        if scripts_only and not path.startswith('scripts/'):
            continue
        source = Path(path).suffix.lower() in SOURCE_EXTENSIONS and kind == 'blob'
        group = family(path) if source else 'non-source'
        entries.append(dict(tree=label, revision=revision, path=path, git_blob=identity,
                            bytes=int(size) if size != '-' else 0, source=source,
                            family=group, role=source_role(path, group) if source else kind))
    objects = sorted({item['git_blob'] for item in entries if item['source']})
    blobs = read_sources(repository, objects, cache)
    symbols = []
    for item in entries:
        if not item['source']:
            continue
        raw = blobs[item['git_blob']]
        source = raw.decode('utf-8-sig', errors='replace')
        item['sha256'] = hashlib.sha256(raw).hexdigest()
        item['lines'] = len(source.splitlines())
        found = []
        if item['role'] not in ['third-party', 'generated']:
            found = [dict(name=match.group(1), line=source.count('\n', 0, match.start()) + 1)
                     for match in SYMBOLS.finditer(source)]
        item['symbols'] = len(found)
        item['inline_test_attributes'] = len(re.findall(r'#\[test\]', source))
        symbols.append(dict(tree=label, path=item['path'], symbols=found))
    return entries, symbols


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repository', type=Path, default=Path('F:/repos/excel-shrink-rust'))
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    paper = Path(__file__).resolve().parents[2]
    production = git(args.repository, 'rev-parse', 'HEAD').decode().strip()
    historical = git(args.repository, 'rev-parse', 'origin/experiments/visual-validation').decode().strip()
    paper_revision = git(paper, 'rev-parse', 'HEAD').decode().strip()
    trees = [
        (args.repository, production, 'production', False),
        (args.repository, historical, 'capture-experiments', False),
        (paper, paper_revision, 'paper-scripts', True),
    ]
    seen = {production, historical}
    branches = []
    for reference in git(args.repository, 'for-each-ref', '--format=%(refname)',
                         'refs/remotes/origin').decode().splitlines():
        if reference.endswith('/HEAD'):
            continue
        revision = git(args.repository, 'rev-parse', reference).decode().strip()
        ancestor = subprocess.run(
            ['git', '-C', str(args.repository), 'merge-base', '--is-ancestor',
             revision, production], check=False,
        ).returncode
        if ancestor not in [0, 1]:
            raise RuntimeError(f'Cannot inspect branch ancestry for {reference}')
        branches.append(dict(reference=reference, revision=revision,
                             ancestor_of_production=ancestor == 0))
        if revision not in seen and ancestor == 1:
            trees.append((args.repository, revision, reference.removeprefix('refs/remotes/'), False))
            seen.add(revision)
    rows, symbols, cache = [], [], {}
    for repository, revision, label, scripts_only in trees:
        entries, index = inspect_tree(repository, revision, label, cache, scripts_only)
        rows.extend(entries)
        symbols.extend(index)
    args.output.mkdir(parents=True, exist_ok=True)
    fields = ['tree', 'revision', 'path', 'git_blob', 'bytes', 'source', 'family',
              'role', 'sha256', 'lines', 'symbols', 'inline_test_attributes']
    with (args.output / 'files.csv').open('w', newline='', encoding='utf-8') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    with gzip.GzipFile(args.output / 'symbols.json.gz', 'wb', mtime=0) as stream:
        stream.write(json.dumps(symbols, separators=(',', ':')).encode())
    summary = {
        'scope': 'Complete production and separate capture-experiment trees, every other '
                 'non-ancestor remote head, and tracked paper scripts. File and lexical '
                 'symbol inventory with editorial family review, not line-by-line '
                 'correctness verification. Temporary untracked local files are outside scope.',
        'revisions': dict(production=production, capture_experiments=historical,
                          paper_scripts=paper_revision),
        'remote_branches': branches,
        'source_extensions': sorted(SOURCE_EXTENSIONS),
        'unique_source_blobs': len({row['git_blob'] for row in rows if row['source']}),
        'trees': {},
    }
    for _, _, tree, _ in trees:
        selected = [item for item in rows if item['tree'] == tree]
        source = [item for item in selected if item['source']]
        summary['trees'][tree] = dict(files=len(selected), source_files=len(source),
                                    source_lines=sum(item['lines'] for item in source),
                                    roles=dict(sorted(Counter(item['role'] for item in source).items())),
                                    families=dict(sorted(Counter(item['family'] for item in source).items())))
    (args.output / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(summary, indent=2))


if __name__ == '__main__':
    main()
