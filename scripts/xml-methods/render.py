#!/usr/bin/env python3
"""Build editable XML figures and a review atlas from the pinned rule inventory.

Run from any directory. Pygments and ReportLab are local figure-build dependencies.
The manuscript consumes the resulting PDFs without executing this script.
"""
from __future__ import annotations

import base64
import difflib
import hashlib
import json
import re
import zipfile
from pathlib import Path

from PIL import Image, ImageChops
from pygments import lex
from pygments.lexers import XmlLexer
from pygments.token import Comment, Name, String as StringToken
from reportlab.graphics import renderPDF, renderSVG
from reportlab.graphics.shapes import (
    STATE_DEFAULTS,
    Drawing,
    Group,
    Image as FigureImage,
    Path as Curve,
    Rect,
    String,
)
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from native_insets import merge_native_insets, native_asset

ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / 'papers/sheetshrink/evidence/xml-methods'
OUTPUT = ROOT / 'papers/sheetshrink/figures/xml-methods'
COLUMN = 252.0
INK = '#25384b'
MUTED = '#536778'
RED = '#a33744'
GREEN = '#23745a'
BLUE = '#1b628b'
FONT_SIZE = 8.2
LEADING = 11.5
GUTTER = 19.0
MARKER_WIDTH = 1.2
REMOVED_FILL = '#efc6cf'
ADDED_FILL = '#c7e1d2'
FONT_FILES = {
    'Code': 'DejaVuSansMono.ttf',
    'CodeBold': 'DejaVuSansMono-Bold.ttf',
    'PaperSerif': 'Tinos-Regular.ttf',
    'PaperSerif-Bold': 'Tinos-Bold.ttf',
}


def register_fonts() -> None:
    """Embed the bundled, licensed monospaced fonts in each generated PDF."""
    fonts = Path(__file__).with_name('fonts')
    for family, filename in FONT_FILES.items():
        pdfmetrics.registerFont(TTFont(family, str(fonts / filename)))
    # The graphics renderer also emits an initial font state before drawing text.
    STATE_DEFAULTS['fontName'] = 'PaperSerif'


def text(drawing: Drawing, x: float, top: float, value: str, size=9, color=INK, font='PaperSerif') -> None:
    drawing.add(String(x, drawing.height - top - size, value, fontName=font,
                       fontSize=size, fillColor=HexColor(color)))


def rect(drawing: Drawing, x: float, top: float, width: float, height: float,
         fill: str, stroke: str | None = None) -> None:
    drawing.add(Rect(x, drawing.height - top - height, width, height,
                     fillColor=HexColor(fill), strokeColor=HexColor(stroke) if stroke else None,
                     strokeWidth=0.45))


def paragraph(drawing: Drawing, x: float, top: float, value: str,
              width: float, size=8.5, color=INK, font='PaperSerif') -> float:
    """Wrap prose without shrinking text to conceal an overflow."""
    line = ''
    y = top
    for word in value.split():
        candidate = f'{line} {word}'.strip()
        if line and pdfmetrics.stringWidth(candidate, font, size) > width:
            text(drawing, x, y, line, size, color, font)
            y += size * 1.35
            line = word
        else:
            line = candidate
    if line:
        text(drawing, x, y, line, size, color, font)
        y += size * 1.35
    return y


def token_color(token) -> str:
    if token in Comment:
        return '#71808d'
    if token in Name.Tag:
        return '#16577f'
    if token in Name.Attribute:
        return '#94354d'
    if token in StringToken:
        return '#27674f'
    return INK


def code_line(drawing: Drawing, x: float, y: float, line: str, available: float,
              size=FONT_SIZE) -> None:
    width = pdfmetrics.stringWidth(line, 'Code', size)
    if width > available + 0.1:
        raise ValueError(f'Code exceeds its panel by {width-available:.1f} pt: {line}')
    for token, value in lex(line, XmlLexer(stripnl=False, ensurenl=False)):
        if value:
            text(drawing, x, y, value, size, token_color(token), 'Code')
            x += pdfmetrics.stringWidth(value, 'Code', size)


def is_editorial(line: str) -> bool:
    """Explanatory comments and ellipses are not inserted OOXML records."""
    return line.lstrip().startswith('<!--') or line.strip() == '...'


def changed_lines(before: list[str], after: list[str], removed_placeholders: tuple[str, ...] = ()) -> tuple[set[int], set[int], list[tuple]]:
    opcodes = difflib.SequenceMatcher(a=before, b=after, autojunk=False).get_opcodes()
    removed, added = set(), set()
    for tag, a0, a1, b0, b1 in opcodes:
        if tag in ('replace', 'delete'):
            removed.update(index for index in range(a0, a1) if not is_editorial(before[index]) or before[index] in removed_placeholders)
        if tag in ('replace', 'insert'):
            added.update(index for index in range(b0, b1) if not is_editorial(after[index]))
    return removed, added, opcodes


def change_band(drawing: Drawing, panel_left: float, panel_right: float, gutter_edge: float,
                source_top: float, source_bottom: float,
                target_top: float, target_bottom: float, added: bool) -> None:
    """Join a highlighted block to its exact destination gap, or the reverse.

    Coordinates follow line-box boundaries. Removals use the right gutter and additions use the left gutter.
    A missing block has a 1.2-point marker across its panel at the insertion gap.
    """
    pale = ADDED_FILL if added else REMOVED_FILL
    band = Curve(fillColor=HexColor(pale), strokeColor=None)
    edge = panel_left + 0.15 if added else panel_right - 0.15
    band.moveTo(edge, drawing.height - source_top)
    band.curveTo(gutter_edge, drawing.height - source_top,
                 gutter_edge, drawing.height - target_top,
                 edge, drawing.height - target_top)
    band.lineTo(edge, drawing.height - target_bottom)
    band.curveTo(gutter_edge, drawing.height - target_bottom,
                 gutter_edge, drawing.height - source_bottom,
                 edge, drawing.height - source_bottom)
    band.closePath()
    drawing.add(band)
    # The vertical distance between two cubic boundaries is not their normal
    # distance. A same-color center stroke supplies the minimum normal width,
    # including the near-vertical bend where the original filled band pinched.
    center = Curve(fillColor=None, strokeColor=HexColor(pale),
                   strokeWidth=MARKER_WIDTH, strokeLineCap=1)
    source_center = (source_top + source_bottom) / 2
    target_center = (target_top + target_bottom) / 2
    center.moveTo(edge, drawing.height - source_center)
    center.curveTo(gutter_edge, drawing.height - source_center,
                   gutter_edge, drawing.height - target_center,
                   edge, drawing.height - target_center)
    drawing.add(center)
    marker_top = source_top if added else target_top
    rect(drawing, panel_left, marker_top, panel_right - panel_left, MARKER_WIDTH, pale)


def line_block(top: float, start: int, end: int) -> tuple[float, float]:
    """Return a changed block or a narrow, centered marker at an empty gap."""
    first = top + 20 + start * LEADING
    return (first - MARKER_WIDTH / 2, first + MARKER_WIDTH / 2) if start == end else (first, top + 20 + end * LEADING)


def panel(drawing: Drawing, x: float, top: float, width: float, label: str,
          lines: list[str], changed: set[int], add: bool) -> float:
    text(drawing, x, top, label, 8, MUTED, 'PaperSerif-Bold')
    box_top = top + 15
    height = 12 + len(lines) * LEADING
    rect(drawing, x, box_top, width, height, '#f8fafc', '#d9e1e8')
    color, pale, marker = (GREEN, ADDED_FILL, '+') if add else (RED, REMOVED_FILL, '-')
    for index, line in enumerate(lines):
        y = box_top + 6 + index * LEADING
        if index in changed:
            rect(drawing, x, y - 1, width, LEADING, pale)
            text(drawing, x + 3, y, marker, FONT_SIZE, color, 'CodeBold')
        code_line(drawing, x + 12, y, line, width - 16)
    return box_top + height


def code_diff(drawing: Drawing, case: dict, top: float, width=COLUMN) -> float:
    """Keep a complete before/after diff inside one column with folded bands."""
    before, after = case['before'], case['after']
    removed, added, opcodes = changed_lines(before, after, tuple(case.get('removed_placeholders', [])))
    decision = case.get('comparison') == 'decision'
    if decision:
        removed, added = set(), set()
    panel_left = GUTTER
    panel_width = width - 2 * GUTTER
    panel_right = panel_left + panel_width
    end_before = panel(drawing, panel_left, top, panel_width, 'Source', before, removed, False)
    after_top = end_before + 13
    label = 'Decision' if decision else ('Guarded result' if before != after else 'Retained by the guard')
    end_after = panel(drawing, panel_left, after_top, panel_width, label, after, added, True)
    if decision:
        return end_after
    for tag, a0, a1, b0, b1 in opcodes:
        # Replacements show the removed block and the added block separately.
        # Editorial comments describe omitted scope and never create green edits.
        if tag in ('replace', 'delete') and removed.intersection(range(a0, a1)):
            source = line_block(top, a0, a1)
            target = line_block(after_top, b0, b0)
            change_band(drawing, panel_left, panel_right, width - 1, *source, *target, added=False)
        if tag in ('replace', 'insert') and added.intersection(range(b0, b1)):
            source = line_block(top, a1, a1)
            target = line_block(after_top, b0, b1)
            change_band(drawing, panel_left, panel_right, 1, *source, *target, added=True)
    if not removed and not added:
        text(drawing, panel_right + 4, top + 20, '=', 12, BLUE, 'PaperSerif-Bold')
    return end_after


def rule_card(case: dict) -> Drawing:
    drawing = Drawing(COLUMN, 400)
    text(drawing, 0, 0, case['id'].split('-')[0], 10, BLUE, 'PaperSerif-Bold')
    text(drawing, 23, 0, case['title'], 10, INK, 'PaperSerif-Bold')
    top = 23
    if case.get('policy'):
        top = paragraph(drawing, 0, top, case['policy'], COLUMN, size=7.8, color=BLUE) + 5
    end = code_diff(drawing, case, top)
    end = paragraph(drawing, 0, end + 10, case['guard'], COLUMN, size=8.4)
    end = paragraph(drawing, 0, end + 5, 'Unsafe shortcut. ' + case['risk'], COLUMN,
                    size=8.4, color=RED)
    text(drawing, 0, end + 6, 'Abbreviated XML. Style IDs are local.', 7.3, MUTED)
    height = end + 23
    # Keep the same top coordinates while trimming only the unused bottom margin.
    group = Group(*drawing.contents)
    group.translate(0, height - drawing.height)
    result = Drawing(COLUMN, height)
    result.add(group)
    return result


def image(drawing: Drawing, path: Path, x: float, top: float, width: float) -> float:
    vector, preview = native_asset(path)
    with Image.open(preview) as bitmap:
        height = width * bitmap.height / bitmap.width
    drawing.add(FigureImage(x, drawing.height - top - height, width, height, str(preview)))
    if not hasattr(drawing, '_native_insets'):
        drawing._native_insets = []
    drawing._native_insets.append(dict(path=str(vector), x=x, top=top, width=width, height=height))
    return height


def draw_at(drawing: Drawing, pdf: canvas.Canvas, x: float, y: float) -> None:
    """Draw a panel and translate its declared native ranges into page coordinates."""
    renderPDF.draw(drawing, pdf, x, y)
    if not hasattr(pdf, '_native_insets'):
        pdf._native_insets = []
    for inset in getattr(drawing, '_native_insets', []):
        pdf._native_insets.append(dict(path=inset['path'], page=pdf.getPageNumber() - 1,
                                      x=x + inset['x'], y=y + drawing.height - inset['top'] - inset['height'],
                                      width=inset['width']))


def native_font_figure(case: dict) -> Drawing:
    """Pair the abbreviated transformation with three real Excel vector exports."""
    drawing = Drawing(COLUMN, 445)
    text(drawing, 0, 0, 'One font witness preserves row layout', 10, INK, 'PaperSerif-Bold')
    top = paragraph(drawing, 0, 18, 'A2 has Arial 10. Empty B2 and C2 have Arial 18.', COLUMN, 8.5)
    top = code_diff(drawing, case, top + 5)
    top += 12
    labels = [('source', 'Original control', '23.25 pt'),
              ('one-witness', 'One witness retained', '23.25 pt'),
              ('no-witness', 'Both witnesses removed', '12.75 pt')]
    for variant, label, height in labels:
        color = RED if variant == 'no-witness' else BLUE
        text(drawing, 0, top, label, 8.4, color, 'PaperSerif-Bold')
        text(drawing, 205, top, height, 8.4, color, 'PaperSerif-Bold')
        top += 15
        top += image(drawing, EVIDENCE / 'native-font' / f'{variant}.png', 0, top, COLUMN)
        top += 9
    top = paragraph(drawing, 0, top, 'Native Excel vector export after full rebuild. Removing the repeated witness is safe here. Removing both moves the next visible row.', COLUMN, 8.2)
    height = top + 3
    group = Group(*drawing.contents)
    group.translate(0, height - drawing.height)
    result = Drawing(COLUMN, height)
    result.add(group)
    result._native_insets = drawing._native_insets
    return result


def native_fill_figure() -> Drawing:
    """Connect a verified cell deletion to retained, unmodified Excel captures."""
    drawing = Drawing(516, 184)
    text(drawing, 0, 0, 'An empty gray cell masks an inherited pink fill', 11, INK, 'PaperSerif-Bold')
    paragraph(drawing, 0, 21, 'Column B uses XF 51 (pink). Cell B31 uses XF 98 (gray).', 516, 9)
    variants = [(0, 'Original', 'Original'), (174, 'Default', 'SheetShrink Default'),
                (348, 'Check-Performance', 'Excel Check Performance')]
    for x, variant, label in variants:
        color = RED if variant == 'Check-Performance' else BLUE
        text(drawing, x, 44, label, 8.8, color, 'PaperSerif-Bold')
        code = '<!-- B31 removed -->' if variant == 'Check-Performance' else '<c r="B31" s="98"/>'
        rect(drawing, x, 60, 168, 20, '#fbe9eb' if variant == 'Check-Performance' else '#f0f6fa')
        code_line(drawing, x + 6, 65, code, 156, 8)
        image(drawing, EVIDENCE / 'enron-fill' / f'{variant}.png', x, 89, 168)
    paragraph(drawing, 0, 137, 'PL-Consolidated!A25:H37, vector export after full rebuild. Original and Default match pixel-for-pixel here. Removing the empty gray cells exposes the inherited pink background.', 516, 9)
    return drawing


def schematic_effects() -> Drawing:
    """Illustrate four guard failures without presenting a model as native evidence."""
    drawing = Drawing(516, 318)
    titles = ['11  Contiguous centering span', '04  Local no-fill mask',
              '25  Empty column boundaries', '13  Visible exterior tail']
    for index, title in enumerate(titles):
        x = 0 if index % 2 == 0 else 270
        top = 0 if index < 2 else 166
        text(drawing, x, top, title, 10, INK, 'PaperSerif-Bold')
        text(drawing, x, top + 21, 'Keep', 8, BLUE, 'PaperSerif-Bold')
        text(drawing, x, top + 76, 'Unsafe shortcut', 8, RED, 'PaperSerif-Bold')
        if index == 0:
            for y, center in [(top + 37, x + 122), (top + 92, x + 80)]:
                for col in range(4):
                    rect(drawing, x + 34 + col * 44, y, 44, 24, '#f8fafc', '#b8c6d2')
                title_width = pdfmetrics.stringWidth('Heading', 'PaperSerif', 9)
                text(drawing, center - title_width / 2, y + 6, 'Heading', 9)
            paragraph(drawing, x, top + 124, 'One surviving blank cannot replace the complete span.', 246, 8.4)
        elif index == 1:
            for y, fills in [(top + 37, ['#73a0dd', '#ffffff', '#ffffff', '#73a0dd']),
                             (top + 92, ['#73a0dd', '#73a0dd', '#ffffff', '#73a0dd'])]:
                for col, fill in enumerate(fills):
                    rect(drawing, x + 34 + col * 44, y, 44, 24, fill, '#7d96ad')
                    text(drawing, x + 51 + col * 44, y + 6, ['J','K','L','M'][col], 8)
            paragraph(drawing, x, top + 124, 'Removing the K-cell mask exposes its blue parent fill.', 246, 8.4)
        elif index == 2:
            for y, widths in [(top + 37, [24, 108, 44]), (top + 92, [58, 60, 58])]:
                left = x + 34
                for width in widths:
                    rect(drawing, left, y, width, 24, '#ffffff', '#596f80')
                    left += width
            paragraph(drawing, x, top + 124, 'Changing widths moves borders even without values.', 246, 8.4)
        else:
            for y, fill in [(top + 37, '#2875ca'), (top + 92, '#ffffff')]:
                rect(drawing, x + 34, y, 176, 7, '#e5eaf0', '#9aa9b7')
                rect(drawing, x + 34, y + 7, 176, 20, fill, '#9aa9b7')
            paragraph(drawing, x, top + 124, 'A blue tail over white columns remains visible.', 246, 8.4)
    return drawing


def save_figure(drawing: Drawing, stem: str) -> None:
    pdf_path = OUTPUT / f'{stem}.pdf'
    renderPDF.drawToFile(drawing, str(pdf_path), initialFontName='PaperSerif')
    insets = [dict(path=inset['path'], page=0, x=inset['x'],
                   y=drawing.height - inset['top'] - inset['height'], width=inset['width'])
              for inset in getattr(drawing, '_native_insets', [])]
    merge_native_insets(pdf_path, insets)
    svg_path = OUTPUT / f'{stem}.svg'
    renderSVG.drawToFile(drawing, str(svg_path))
    embed_svg_fonts(svg_path, stem)


def embed_svg_fonts(svg_path: Path, title: str) -> None:
    """Keep editable SVG text portable with the same licensed fonts as the PDF."""
    svg = svg_path.read_text(encoding='utf-8')
    styles = []
    for family in sorted(set(re.findall(r'font-family: ([^;]+);', svg))):
        font_data = (Path(__file__).with_name('fonts') / FONT_FILES[family]).read_bytes()
        encoded = base64.b64encode(font_data).decode('ascii')
        styles.append(f'@font-face{{font-family:"{family}";src:url(data:font/ttf;base64,{encoded}) format("truetype");}}')
    svg = svg.replace('<title>...</title>', '<title>' + title + '</title><defs><style>' + ''.join(styles) + '</style></defs>')
    svg_path.write_text(svg, encoding='utf-8')


def header(pdf: canvas.Canvas, title: str, page: int, subtitle: str = '') -> None:
    pdf.setFillColor(HexColor(INK))
    pdf.setFont('PaperSerif-Bold', 15)
    pdf.drawString(38, 803, title)
    pdf.setFont('PaperSerif', 8.5)
    pdf.setFillColor(HexColor(MUTED))
    pdf.drawString(38, 785, subtitle)
    pdf.setFont('PaperSerif', 8)
    pdf.drawString(38, 23, 'SheetShrink | XML cleanup examples | 10 November 2026')
    pdf.drawRightString(557, 23, str(page))


def coverage_pages(pdf: canvas.Canvas, page: int) -> int:
    """Show the reviewed rule families and account for every declared cleanup cause."""
    coverage = json.loads((EVIDENCE / 'coverage.json').read_text(encoding='utf-8'))
    for title, entries, family_page in [
        ('Rule-family coverage', coverage['families'], True),
        ('Declared cleanup causes and their status', coverage['cleanup_causes'], False),
    ]:
        page += 1
        header(pdf, title, page, 'Coverage inventory. Disabled and no-op paths are explicitly distinguished.')
        top = 760
        for entry in entries:
            drawing = Drawing(519, 90)
            text(drawing, 0, 0, entry['name'], 10, INK, 'PaperSerif-Bold')
            identifiers = ', '.join(f'{number:02d}' for number in entry['cases']) or 'No active Default transformation shown'
            y = paragraph(drawing, 0, 17, 'Examples ' + identifiers, 519, 9, BLUE)
            status = entry['specification'] if family_page else entry['status']
            y = paragraph(drawing, 0, y + 3, status + '. ' + entry['note'], 519, 9)
            draw_at(drawing, pdf, 38, top - 90)
            top -= y + 19
        if family_page:
            drawing = Drawing(519, 170)
            paragraph(drawing, 0, 0, coverage['scope'], 519, 9)
            draw_at(drawing, pdf, 38, 45)
        pdf.showPage()
    return page


def atlas(cases: list[dict], drawings: list[Drawing], font: Drawing, fill: Drawing,
          revision: str) -> None:
    """Lay out full-size column candidates, native evidence and navigable provenance."""
    pdf = canvas.Canvas(str(OUTPUT / 'xml-cleanup-atlas.pdf'), pagesize=(595.28, 841.89), initialFontName='PaperSerif')
    pdf.setTitle('SheetShrink XML cleanup examples')
    pdf.setAuthor('SheetShrink project')
    page = 1
    header(pdf, 'XML cleanup and its preservation guards', page,
           'Abbreviated source examples, negative cases and native rendering evidence')
    intro = Drawing(519, 690)
    y = paragraph(intro, 0, 0,
        f'This atlas groups the reviewed behavior specification, implementation, commit history and focused regressions into {len(cases)} visual examples. Production Default rules and explicitly labeled diagnostic fixtures have separate scopes. Experimental Maximum and Too Aggressive settings have different preservation policies.', 519, 11)
    y = paragraph(intro, 0, y + 15,
        'Read each source/result pair at its native one-column width. Red backgrounds and minus signs identify removed lines on the right. Green backgrounds and plus signs identify added lines on the left. An equals sign identifies a protected structure. Folded color bands connect edited ranges to thin markers at their exact gaps. These cues supplement syntax color.', 519, 10)
    y = paragraph(intro, 0, y + 15,
        'The XML is illustrative. Namespace wrappers, irrelevant attributes, cached payloads and repeated cells are omitted. Style indices have meaning only within their own example. Comments describe omitted scope or an exclusion. A red omission label stands for deleted records. Explanatory comments are not literal output. The complete native font-control packages are retained separately.', 519, 10)
    y = paragraph(intro, 0, y + 15,
        'The source index pins each case to a file and a regression-test or implementation anchor. A matching unit test checks its own fixture. It does not prove that an abbreviated picture is a complete runnable workbook or that every possible Excel presentation is preserved.', 519, 10)
    y = paragraph(intro, 0, y + 15,
        'Native evidence. The reduced font control has three native Excel vector exports after CalculateFullRebuild. The Enron fill comparison is re-exported from the retained rebuild/Save controls. All other outcome descriptions are traced to their named tests or investigation records. No schematic is presented as an Excel screenshot.', 519, 10)
    y = paragraph(intro, 0, y + 15,
        'Boundaries. Unsupported XML can exclude an input. Unknown inheritance or references follow the local conservative guard. Ordinary hidden tables are not covered by hidden-tail admission. Alias-name cleanup does not currently prove equality of formula bodies. Selected-region native equality is not whole-workbook certification.', 519, 10)
    y = paragraph(intro, 0, y + 20,
        f'Reviewed source revision {revision}. Core verification completed with 298 passed, zero failed and one manual timing test ignored. The source repository was read and tested for this figure task. Cleanup behavior was not edited.', 519, 9)
    paragraph(intro, 0, y + 18,
        'The design follows Animated Vega-Lite Figure 2 and Bluefish Figure 2. This atlas uses newly drawn geometry, editable text, local diff highlighting and native control images.', 519, 9)
    draw_at(intro, pdf, 38, 70)
    pdf.showPage()
    # A variable-height two-column grid keeps every code panel at native 8.2 pt.
    row_top = 765
    for offset in range(0, len(drawings), 2):
        row = drawings[offset:offset + 2]
        height = max(item.height for item in row)
        if row_top - height < 45:
            if row_top != 765:
                pdf.showPage()
            page += 1
            header(pdf, 'Guarded XML transformations', page, 'Each panel is 252 pt wide. Code is 8.2 pt at this scale.')
            row_top = 765
        elif row_top == 765:
            page += 1
            header(pdf, 'Guarded XML transformations', page, 'Each panel is 252 pt wide. Code is 8.2 pt at this scale.')
        for x, item in zip([38, 305], row):
            draw_at(item, pdf, x, row_top - item.height)
        row_top -= height + 23
    pdf.showPage()
    page += 1
    header(pdf, 'One representative, one native layout effect', page,
           'Reduced fixture | installed Excel 16.0 build 20326 | full rebuild before each capture')
    draw_at(font, pdf, 38, 765 - font.height)
    note = Drawing(240, 500)
    y = paragraph(note, 0, 0, 'What was changed', 240, 11, INK, 'PaperSerif-Bold')
    y = paragraph(note, 0, y + 10, 'The three input packages differ only in sheet1.xml. The source has two empty Arial 18 cells. The one-witness package removes C2. The negative control removes B2 as well. All use the same fonts, text and columns.', 240, 10)
    y = paragraph(note, 0, y + 14, 'What Excel rendered', 240, 11, INK, 'PaperSerif-Bold')
    y = paragraph(note, 0, y + 10, 'The source and one-witness vector exports have identical decoded pixels at 600 dpi. Both have row 2 height 23.25 pt. The negative control has row 2 height 12.75 pt and moves the next visible row upward. Its exported range is shorter.', 240, 10)
    y = paragraph(note, 0, y + 14, 'Scope', 240, 11, INK, 'PaperSerif-Bold')
    paragraph(note, 0, y + 10, 'These are constructed native controls, not the Kulturfinanzbericht screenshots. They illustrate the font-witness mechanism independently of the documented corpus case. They are not a full validation run or a measurement of shrink throughput.', 240, 10)
    draw_at(note, pdf, 305, 265)
    pdf.showPage()
    page += 1
    header(pdf, 'A smaller file can expose a different fill', page,
           'Retained Enron case ccde4294bab9094d7d73 | PL-Consolidated!A25:H37')
    draw_at(fill, pdf, 38, 765 - fill.height)
    # At review scale the source images remain legible as full-width strips.
    top = 530
    for variant, label in [('Original', 'Original, rebuilt'), ('Default', 'Default, rebuilt'), ('Check-Performance', 'Check Performance, rebuilt')]:
        pdf.setFont('PaperSerif-Bold', 10)
        pdf.setFillColor(HexColor(RED if variant == 'Check-Performance' else INK))
        pdf.drawString(38, top, label)
        vector, preview = native_asset(EVIDENCE / 'enron-fill' / f'{variant}.png')
        with Image.open(preview) as bitmap:
            height = 516 * bitmap.height / bitmap.width
        bottom = top - height - 10
        pdf.drawImage(str(preview), 38, bottom, width=516, height=height)
        pdf._native_insets.append(dict(path=str(vector), page=pdf.getPageNumber() - 1,
                                      x=38, y=bottom, width=516))
        top -= height + 39
    pdf.showPage()
    page += 1
    header(pdf, 'What the rejected shortcuts would change', page,
           'Schematic illustrations | these are not Excel screenshots or measured coordinates')
    draw_at(schematic_effects(), pdf, 38, 765 - 318)
    pdf.showPage()
    page = coverage_pages(pdf, page)
    for start in range(0, len(cases), 6):
        page += 1
        header(pdf, 'Source index and evidence limits', page,
               f'Pinned implementation {revision[:12]} | full paths, anchors and hashes are in cases.json')
        top = 760
        for item in cases[start:start + 6]:
            d = Drawing(519, 110)
            text(d, 0, 0, item['id'] + '  ' + item['title'], 10, INK, 'PaperSerif-Bold')
            source = item['source']
            y = paragraph(d, 0, 18, source['path'], 519, 8, MUTED)
            y = paragraph(d, 0, y + 2, source['kind'] + ' | ' + source['test'], 519, 8, BLUE)
            y = paragraph(d, 0, y + 4, item['note'], 519, 8.6)
            draw_at(d, pdf, 38, top - 110)
            pdf.linkURL('https://github.com/Alexander-Aue-Johr/excel-shrink-rust/blob/' + revision + '/' + source['path'] + '#L' + str(source['line']), (38, top - 48, 557, top), relative=0, thickness=0)
            top -= 119
        pdf.showPage()
    pdf.save()
    merge_native_insets(OUTPUT / 'xml-cleanup-atlas.pdf', pdf._native_insets)


def verify_native_evidence() -> dict:
    """Compare decoded native captures and retain byte identities for every input."""
    font_dir = EVIDENCE / 'native-font'
    packages = {}
    for name in ['source', 'one-witness', 'no-witness']:
        with zipfile.ZipFile(font_dir / f'{name}.xlsx') as archive:
            assert archive.testzip() is None
            packages[name] = {part: archive.read(part) for part in archive.namelist()}
    source_parts = packages['source']
    sheet_part = 'xl/worksheets/sheet1.xml'
    for name, parts in packages.items():
        assert parts.keys() == source_parts.keys()
        assert all(parts[part] == source_parts[part] for part in parts if part != sheet_part)
    assert packages['one-witness'][sheet_part] == source_parts[sheet_part].replace(b'<c r="C2" s="1"/>', b'')
    assert packages['no-witness'][sheet_part] == packages['one-witness'][sheet_part].replace(b'<c r="B2" s="1"/>', b'')
    receipt = json.loads((font_dir / 'receipt.json').read_text(encoding='utf-8'))
    heights = {row['variant']: row['row_2_height_points'] for row in receipt}
    assert heights['source'] == heights['one-witness'] == 23.25
    assert heights['no-witness'] == 12.75
    with Image.open(font_dir / 'source.png') as source, Image.open(font_dir / 'one-witness.png') as kept:
        assert source.size == kept.size
        assert ImageChops.difference(source.convert('RGBA'), kept.convert('RGBA')).getbbox() is None
    with Image.open(EVIDENCE / 'enron-fill/Original.png') as source, Image.open(EVIDENCE / 'enron-fill/Default.png') as kept:
        assert source.size == kept.size
        assert ImageChops.difference(source.convert('RGBA'), kept.convert('RGBA')).getbbox() is None
    for group, original, kept in [('font', 'source', 'one-witness'), ('enron', 'Original', 'Default')]:
        folder = EVIDENCE / 'native-vector' / group
        with Image.open(folder / f'{original}.png') as left, Image.open(folder / f'{kept}.png') as right:
            assert left.size == right.size
            assert ImageChops.difference(left.convert('RGBA'), right.convert('RGBA')).getbbox() is None
    files = {str(path.relative_to(EVIDENCE)).replace('\\', '/'): hashlib.sha256(path.read_bytes()).hexdigest()
             for folder in ['native-font', 'enron-fill', 'native-vector'] for path in (EVIDENCE / folder).rglob('*') if path.is_file()}
    return {'native_control_packages_differ_only_in_stated_cells': True,
            'font_control_source_equals_one_witness_pixels': True,
            'enron_region_original_equals_default_pixels': True,
            'font_row_2_height_points': heights,
            'vector_export_control_pairs_equal_at_600_dpi': True,
            'band_minimum_width_points': MARKER_WIDTH,
            'band_sides': {'removed': 'right', 'added': 'left'},
            'input_sha256': files}


def main() -> None:
    register_fonts()
    OUTPUT.mkdir(parents=True, exist_ok=True)
    data = json.loads((EVIDENCE / 'cases.json').read_text(encoding='utf-8'))
    native = verify_native_evidence()
    drawings = [rule_card(case) for case in data['cases']]
    for case, drawing in zip(data['cases'], drawings):
        save_figure(drawing, case['id'])
    font = native_font_figure(next(case for case in data['cases'] if case['id'] == '02-font-witness'))
    fill = native_fill_figure()
    save_figure(font, 'font-witness-native')
    save_figure(fill, 'fill-override-native')
    save_figure(schematic_effects(), 'guard-effects-schematic')
    atlas(data['cases'], drawings, font, fill, data['source_revision'])
    result = {'source_revision': data['source_revision'], 'rule_count': len(drawings),
              'code_font_size_points': FONT_SIZE, 'column_width_points': COLUMN,
              'maximum_card_height_points': max(d.height for d in drawings), **native}
    (EVIDENCE / 'verification.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    from render_manuscript import render_manuscript
    render_manuscript(data['cases'])
    print(json.dumps({key: value for key, value in result.items() if key != 'input_sha256'}, indent=2))


if __name__ == '__main__':
    main()
