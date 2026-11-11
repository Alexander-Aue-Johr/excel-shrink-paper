#!/usr/bin/env python3
"""Export compact manuscript snippets with prose supplied by LaTeX captions."""
import json

from reportlab.graphics.shapes import Drawing, Group

from render import (
    COLUMN, EVIDENCE, code_diff, code_line, rect, register_fonts, save_figure,
)


def diff_snippet(case):
    drawing = Drawing(COLUMN, 400)
    height = code_diff(drawing, case, 0) + 1
    contents = Group(*drawing.contents)
    contents.translate(0, height - drawing.height)
    result = Drawing(COLUMN, height)
    result.add(contents)
    return result


def fill_snippet(line, removed=False):
    drawing = Drawing(168, 21)
    rect(drawing, 0, 0, 168, 21, '#fbe9eb' if removed else '#f0f6fa')
    code_line(drawing, 6, 5, line, 156, 8)
    return drawing


def render_manuscript(cases):
    indexed = {case['id']: case for case in cases}
    save_figure(diff_snippet(indexed['02-font-witness']), 'font-witness-code')
    save_figure(diff_snippet(indexed['08-row-promotion']), '08-row-promotion')
    for stem in ['fill-original-code', 'fill-sheetshrink-code']:
        save_figure(fill_snippet('<c r="B31" s="98"/>'), stem)
    save_figure(fill_snippet('<!-- B31 removed -->', removed=True),
                'fill-check-performance-code')


if __name__ == '__main__':
    register_fonts()
    inventory = json.loads((EVIDENCE / 'cases.json').read_text(encoding='utf-8'))
    render_manuscript(inventory['cases'])
