#!/usr/bin/env python3
"""Crop recorded Excel print exports and generate high-resolution SVG previews."""
import argparse
import json
import subprocess
from pathlib import Path

from PIL import Image, ImageChops
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import RectangleObject

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / 'papers/sheetshrink/evidence/xml-methods/native-vector'


def crop_export(source: Path, target: Path, width: float, height: float) -> None:
    """Crop an unscaled zero-margin export to the recorded native range."""
    reader = PdfReader(source)
    if len(reader.pages) != 1:
        raise ValueError(f'Expected one unscaled print page: {source}')
    page = reader.pages[0]
    page_height = float(page.mediabox.height)
    if width > float(page.mediabox.width) or height > page_height:
        raise ValueError(f'Recorded range exceeds the print page: {source}')
    page.add_transformation(Transformation().translate(0, -(page_height - height)))
    page.mediabox = RectangleObject([0, 0, width, height])
    page.cropbox = page.mediabox
    writer = PdfWriter()
    writer.add_page(page)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open('wb') as stream:
        writer.write(stream)


def verify_positive_pair(group: str, original: str, retained: str) -> None:
    """Compare both positive controls at the same 600 dpi rendering scale."""
    with Image.open(OUTPUT / group / f'{original}.png') as left:
        with Image.open(OUTPUT / group / f'{retained}.png') as right:
            if left.size != right.size:
                raise ValueError(f'{group} positive-control geometry differs')
            if ImageChops.difference(left.convert('RGBA'), right.convert('RGBA')).getbbox():
                raise ValueError(f'{group} positive-control pixels differ')


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--export-directory', type=Path,
                        default=ROOT / 'tmp/pdfs/xml-methods/vector-export')
    args = parser.parse_args()
    receipts = json.loads((args.export_directory / 'receipt.json').read_text(encoding='utf-8-sig'))
    for item in receipts:
        target = OUTPUT / item['group'] / (item['variant'] + '.pdf')
        crop_export(args.export_directory / item['pdf'], target,
                    item['width_points'], item['height_points'])
        subprocess.run(['pdftoppm', '-r', '600', '-png', '-singlefile', str(target),
                        str(target.with_suffix(''))], check=True)
    verify_positive_pair('font', 'source', 'one-witness')
    verify_positive_pair('enron', 'Original', 'Default')
    (OUTPUT / 'receipt.json').write_text(json.dumps(receipts, indent=2) + '\n', encoding='utf-8')
    print(f'Prepared {len(receipts)} vector ranges. Both positive pairs match at 600 dpi.')


if __name__ == '__main__':
    main()
