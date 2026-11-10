"""Place cropped native Excel PDF ranges into the generated figures as vectors."""
from collections import defaultdict
from pathlib import Path

from pypdf import PageObject, PdfReader, PdfWriter, Transformation
from pypdf.generic import ContentStream, DecodedStreamObject, NameObject


def native_asset(path: Path) -> tuple[Path, Path]:
    """Resolve a retained capture to its new vector export and 600 dpi SVG preview."""
    group = 'font' if path.parent.name == 'native-font' else 'enron'
    vector = path.parent.parent / 'native-vector' / group / (path.stem + '.pdf')
    preview = vector.with_suffix('.png')
    if not vector.is_file() or not preview.is_file():
        raise FileNotFoundError(f'Native vector export or preview missing: {vector}')
    return vector, preview


def remove_preview_images(page, reader: PdfReader, expected: int) -> None:
    """Remove the declared PNG preview draws before inserting their PDF counterparts."""
    resources = page['/Resources']
    reference = resources.get('/XObject')
    objects = reference.get_object() if reference is not None else {}
    image_names = {
        name for name, resource in objects.items()
        if resource.get_object().get('/Subtype') == '/Image'
    }
    content = ContentStream(page.get_contents(), reader)
    kept = []
    removed = 0
    for operands, operator in content.operations:
        if operator == b'INLINE IMAGE' or (operator == b'Do' and operands[0] in image_names):
            removed += 1
        else:
            kept.append((operands, operator))
    if removed != expected:
        raise ValueError(f'Expected {expected} declared native images, found {removed}')
    content.operations = kept
    page[NameObject('/Contents')] = content
    for name in image_names:
        del objects[name]


def opaque_native_page(path: Path):
    """Keep each native range on its own white page, including otherwise unpainted cells."""
    source = PdfReader(path).pages[0]
    width, height = float(source.mediabox.width), float(source.mediabox.height)
    page = PageObject.create_blank_page(width=width, height=height)
    background = DecodedStreamObject()
    background.set_data(f'q 1 1 1 rg 0 0 {width} {height} re f Q'.encode('ascii'))
    page[NameObject('/Contents')] = background
    page.merge_page(source)
    return page


def merge_native_insets(path: Path, insets: list[dict]) -> None:
    """Replace every declared native bitmap with the corresponding vector PDF range.

    Code text and diagram geometry remain untouched. Unexpected bitmap draws abort
    the replacement. Only a completely written PDF replaces the original figure.
    """
    if not insets:
        return
    by_page = defaultdict(list)
    for inset in insets:
        by_page[inset['page']].append(inset)
    reader = PdfReader(path)
    writer = PdfWriter()
    for index, page in enumerate(reader.pages):
        placements = by_page.get(index, [])
        if placements:
            remove_preview_images(page, reader, len(placements))
        for inset in placements:
            source = opaque_native_page(Path(inset['path']))
            scale = inset['width'] / float(source.mediabox.width)
            transform = Transformation().scale(scale).translate(inset['x'], inset['y'])
            page.merge_transformed_page(source, transform)
        writer.add_page(page)
    if reader.metadata:
        writer.add_metadata({key: str(value) for key, value in reader.metadata.items()})
    temporary = path.with_suffix('.pending.pdf')
    try:
        with temporary.open('wb') as stream:
            writer.write(stream)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
