#!/usr/bin/env python3
"""Verify selectable XML, embedded fonts and vector-only code panels."""
import json
import re
from pathlib import Path

from pypdf import PdfReader

PAPER = Path(__file__).resolve().parents[2] / 'papers/sheetshrink'
EVIDENCE = PAPER / 'evidence/xml-methods'


def compact(text: str) -> str:
    """Ignore PDF reader whitespace reconstruction when checking XML tokens."""
    return re.sub(r'\s+', '', text)


def inspect_code_figure(case: dict) -> list[str]:
    path = PAPER / 'figures/xml-methods' / (case['id'] + '.pdf')
    page = PdfReader(path).pages[0]
    extracted = compact(page.extract_text())
    problems = []
    for line in case['before'] + case['after']:
        if compact(line) not in extracted:
            problems.append(f"{case['id']} missing code text {line}")
    for resource in page['/Resources']['/Font'].get_object().values():
        font = resource.get_object()
        descriptor = font.get('/FontDescriptor')
        embedded = descriptor and any(
            key in descriptor.get_object()
            for key in ['/FontFile', '/FontFile2', '/FontFile3']
        )
        if not embedded:
            problems.append(f"{case['id']} unembedded font {font.get('/BaseFont')}")
    if page['/Resources'].get('/XObject'):
        problems.append(f"{case['id']} unexpected raster or external object")
    return problems


def main() -> None:
    inventory = json.loads((EVIDENCE / 'cases.json').read_text(encoding='utf-8'))
    problems = [
        problem
        for case in inventory['cases']
        for problem in inspect_code_figure(case)
    ]
    print(f"Code figures checked: {len(inventory['cases'])}. Problems: {len(problems)}.")
    if problems:
        raise SystemExit('\n'.join(problems))
    receipt_path = EVIDENCE / 'verification.json'
    receipt = json.loads(receipt_path.read_text(encoding='utf-8'))
    receipt['pdf_checks'] = {
        'code_figures': len(inventory['cases']),
        'code_text_preserved': True,
        'all_code_figure_fonts_embedded': True,
        'code_figures_have_no_raster_objects': True,
    }
    receipt_path.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
