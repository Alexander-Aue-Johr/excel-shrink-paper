#!/usr/bin/env python3
"""Verify selectable XML, embedded fonts and vector-only code panels."""
import json
import re
from pathlib import Path

from reportlab.lib.colors import HexColor
from pypdf import PdfReader
from pypdf.generic import ContentStream

from render import (ADDED_FILL, REMOVED_FILL, MARKER_WIDTH, Drawing, change_band)

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


def verify_bands() -> None:
    """Keep the marker, band and minimum-width stroke on the same paint contract."""
    for added, fill in [(False, REMOVED_FILL), (True, ADDED_FILL)]:
        drawing = Drawing(252, 200)
        change_band(drawing, 19, 233, 1 if added else 251,
                    29.4 if added else 20, 30.6 if added else 31.5,
                    120 if added else 119.4, 143 if added else 120.6, added)
        band, center, marker = drawing.contents
        assert band.fillColor == center.strokeColor == marker.fillColor == HexColor(fill)
        assert marker.height == MARKER_WIDTH
        assert center.strokeWidth >= marker.height
        assert center.strokeLineCap == 1


def verify_vector_outputs() -> None:
    """Reject raster fallbacks in the native figures and complete review atlas."""
    for stem in ['font-witness-native', 'fill-override-native', 'xml-cleanup-atlas']:
        reader = PdfReader(PAPER / 'figures/xml-methods' / (stem + '.pdf'))
        for page in reader.pages:
            assert not any(operator == b'INLINE IMAGE'
                           for _, operator in ContentStream(page.get_contents(), reader).operations)
            assert not any(resource.get_object().get('/Subtype') == '/Image'
                           for resource in page['/Resources'].get('/XObject', {}).values())


def main() -> None:
    verify_bands()
    verify_vector_outputs()
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
        'native_figures_and_atlas_have_no_raster_images': True,
        'band_stroke_at_least_marker_width': True,
        'band_and_marker_colors_match': True,
    }
    receipt_path.write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
