#!/usr/bin/env python3
"""Render the abbreviated DeltaSieve method panels and verify their PDF text.

The JSON inventory records the implementation anchors. These panels explain
contracts and deliberately omit transport, error-context and ownership details.
They are illustrations, not executable replacements for the production code.
"""
import json
import sys
from pathlib import Path

from pygments import lex
from pygments.lexers import get_lexer_by_name
from pypdf import PdfReader
from reportlab.graphics import renderPDF, renderSVG
from reportlab.graphics.shapes import Drawing
from reportlab.pdfbase import pdfmetrics

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'scripts/xml-methods'))
from render import embed_svg_fonts, rect, register_fonts, text, token_color

EVIDENCE = ROOT / 'papers/deltasieve/evidence/capture-development'
OUTPUT = ROOT / 'papers/deltasieve/figures/capture-methods'


def draw_panel(case: dict) -> Drawing:
    """Keep code at column scale and fail on lines exceeding the code box."""
    height = 14 + 11.5 * len(case['code'])
    drawing = Drawing(252, height)
    rect(drawing, 0, 0, 252, height, '#f8fafc', '#d9e1e8')
    lexer = get_lexer_by_name(case['lexer'], stripnl=False, ensurenl=False)
    for index, line in enumerate(case['code']):
        if pdfmetrics.stringWidth(line, 'Code', 8.2) > 236:
            raise ValueError(f"Line exceeds the column in {case['id']}: {line}")
        x = 8
        for token, value in lex(line, lexer):
            text(drawing, x, 6 + index * 11.5, value, 8.2, token_color(token), 'Code')
            x += pdfmetrics.stringWidth(value, 'Code', 8.2)
    return drawing


def main() -> None:
    register_fonts()
    OUTPUT.mkdir(parents=True, exist_ok=True)
    inventory = json.loads((EVIDENCE / 'snippets.json').read_text(encoding='utf-8'))
    for case in inventory['cases']:
        drawing = draw_panel(case)
        pdf = OUTPUT / (case['id'] + '.pdf')
        renderPDF.drawToFile(drawing, str(pdf))
        svg = OUTPUT / (case['id'] + '.svg')
        renderSVG.drawToFile(drawing, str(svg))
        embed_svg_fonts(svg, case['id'])
        page = PdfReader(pdf).pages[0]
        extracted = ''.join(page.extract_text().split())
        assert ''.join(case['title'].split()) not in extracted
        assert ''.join(case['explanation'].split()) not in extracted
        assert ''.join(case['boundary'].split()) not in extracted
        for line in case['code']:
            assert ''.join(line.split()) in extracted, (case['id'], line)
        assert not page['/Resources'].get('/XObject')
        for resource in page['/Resources']['/Font'].get_object().values():
            descriptor = resource.get_object()['/FontDescriptor'].get_object()
            assert any(key in descriptor for key in ['/FontFile', '/FontFile2', '/FontFile3'])
    print(f"Verified {len(inventory['cases'])} vector code panels.")


if __name__ == '__main__':
    main()
