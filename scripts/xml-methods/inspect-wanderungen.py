#!/usr/bin/env python3
"""Record the presentation case from a supplied original Wanderungen workbook.

Only the small workbook catalogue uses ElementTree. Worksheet inspection works
on ZIP-member bytes. This records source cells and does not run SheetShrink.
"""
import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


def inspect(workbook: Path) -> dict:
    """Require all 21 documented L:XFD runs before publishing the source receipt."""
    namespace = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(workbook) as package:
        catalogue = ET.fromstring(package.read('xl/workbook.xml'))
        relationships = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
        targets = {item.attrib['Id']: item.attrib['Target'] for item in relationships}
        sheet = next(item for item in catalogue.find('m:sheets', namespace)
                     if item.attrib['name'] == 'Tabelle 3.4 ')
        relationship_id = sheet.attrib[
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        target = targets[relationship_id]
        part = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        worksheet = package.read(part)
    rows = []
    for match in re.finditer(rb'<row\b[^>]*>.*?</row>', worksheet, re.S):
        cells = re.findall(
            rb'<c\b[^>]*\br="([A-Z]+)(\d+)"[^>]*>\s*<v>0</v>\s*</c>',
            match.group())
        if len(cells) < 16_000:
            continue
        columns = []
        for letters, _ in cells:
            number = 0
            for letter in letters:
                number = number * 26 + letter - ord('A') + 1
            columns.append(number)
        if columns != list(range(12, 16_385)):
            raise ValueError('The source no longer matches the documented L:XFD run')
        row = int(cells[0][1])
        if any(int(coordinate_row) != row for _, coordinate_row in cells):
            raise ValueError('Cell coordinates disagree within a source row')
        rows.append(dict(row=row, first='L', last='XFD', zeros=len(cells)))
    if [item['row'] for item in rows] != list(range(11, 31)) + [32]:
        raise ValueError('The source no longer matches the documented 21 rows')
    return {
        'source_url': 'https://raw.githubusercontent.com/Alexander-Aue-Johr/'
                      'original_destatis_xlsx_files/main/wanderungen-2010120217005.xlsx',
        'source_sha256': hashlib.sha256(workbook.read_bytes()).hexdigest(),
        'sheet': 'Tabelle 3.4 ',
        'part': part,
        'rows': rows,
        'zero_total': sum(item['zeros'] for item in rows),
        'measurement': 'Source ZIP member inspection only. '
                       'No new shrink or native visual validation.',
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('receipt', type=Path)
    args = parser.parse_args()
    result = inspect(args.workbook)
    args.receipt.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print(f"Confirmed {len(result['rows'])} rows and {result['zero_total']:,} zeros.")


if __name__ == '__main__':
    main()
