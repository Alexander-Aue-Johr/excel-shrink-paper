import re
from typing import List, NamedTuple, Optional

class XF(NamedTuple):
    fill_id: Optional[int]
    border_id: Optional[int]
    apply_border: bool
    apply_fill: bool
    has_fill_or_border: bool
    has_apply_fill_or_border: bool
    has_only_font_style: bool

fill_id_pattern = re.compile(b'fillId="([^"]+)"')

border_id_pattern = re.compile(b'borderId="([^"]+)"')

apply_border_pattern = re.compile(b'applyBorder="([^"]+)"')

apply_fill_pattern = re.compile(b'applyFill="([^"]+)"')

def parse_xfs(cell_styles: bytes) -> List[XF]:
    """
    Find all <xf ... /> and <xf ...>...</xf> in cell_styles,
    then parse out integer attributes into a list of XF named tuples.
    """
    xf_matches = re.findall(rb"<xf[^>]*/>|<xf[^>]*>", cell_styles)

    def get_attr(xf_bytes: bytes, pattern) -> Optional[int]:
        m = pattern.search(xf_bytes)
        return m.group(1) if m else None
    
    def get_int_attr(xf_bytes: bytes, pattern) -> Optional[int]:
        m = pattern.search(xf_bytes)
        return int(m.group(1)) if m else None

    all_xfs = []
    for xf in xf_matches:
        fill_id                  = get_int_attr(xf, fill_id_pattern)  
        border_id                = get_int_attr(xf, border_id_pattern) 
        apply_border             = get_attr(xf, fill_id_pattern)  == b'1'
        apply_fill               = get_attr(xf, apply_fill_pattern) == b'1'
        has_fill_or_border       = fill_id != 0 or border_id != 0
        has_apply_fill_or_border = apply_fill or apply_border
        has_only_font_style      = not has_fill_or_border and not has_apply_fill_or_border

        xf_info = XF(fill_id, border_id, apply_border, apply_fill, has_fill_or_border, has_apply_fill_or_border, has_only_font_style)

        all_xfs.append(xf_info)

    return all_xfs
