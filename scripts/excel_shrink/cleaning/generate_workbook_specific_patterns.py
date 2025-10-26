

import re
from typing import List, NamedTuple

from parsing.parse_xfs import XF

def generate_whitespace_cell_pattern(whitespace_indices_joined: bytes | None) -> re.Pattern | None:
    if not whitespace_indices_joined:
        return None

    whitespace_cell_pattern = re.compile(
        rb'(<c\b[^>]*?\bt="s"[^>]*?)><v>(?:'+ whitespace_indices_joined + b')</v></c>',
    )

    return whitespace_cell_pattern

def generate_whitespace_pattern(whitespace_indices_joined: bytes | None) -> re.Pattern | None:
    if not whitespace_indices_joined:
        return None
    
    whitespace_pattern = re.compile(
        rb'<v>(?:'+ whitespace_indices_joined + b')</v>',
    )

    return whitespace_pattern

def generate_whitespace_joined(whitespace_indices: List[int]) -> bytes | None:
    if whitespace_indices.count == 0:
        return None
    
    whitespace_indices_bytestrings = [str(idx).encode() for idx in whitespace_indices]
    
    whitespace_indices_joined = b"|".join(whitespace_indices_bytestrings)

    return whitespace_indices_joined

def find_no_style_but_font_ids(all_xfs: List[XF]) -> set[int]:
    """
    Identify which style indices have applyFill or applyBorder,
    and which have no style but do have a font.
    """
    return {
        i for i, xf in enumerate(all_xfs)
        if ((xf.fill_id is None or xf.fill_id == 0)
            and (xf.border_id is None or xf.border_id == 0)
            and not xf.apply_fill
            and not xf.apply_border)
    }

def generate_pattern_consecutive_only_font_styled_empty_cells(only_font_style_ids_joined : bytes | None) -> re.Pattern:
    if only_font_style_ids_joined is None:
        pattern_consecutive_styleless_empty_cells = re.compile(
            rb'(?:<c\b'
                # The tag has no s= attribute at all
                rb'(?![^>]*\bs=)[^>]*'
            rb'/>)+'
        )

        return pattern_consecutive_styleless_empty_cells

    pattern_consecutive_only_font_styled_empty_cells = re.compile(
        rb'(?:<c\b(?:'
        # The tag has no s= attribute at all
            rb'(?![^>]*\bs=)[^>]*'
            rb'|'
            # The tag has s="..." with an allowed value
            rb'[^>]*\bs="(?:' + only_font_style_ids_joined + rb')"[^>]*'
        rb')/>)+'
    )

    return pattern_consecutive_only_font_styled_empty_cells

def generate_pattern_remove_irrelevant_attributes_for_empty_rows(only_font_style_ids_joined : bytes | None) -> re.Pattern:
    if only_font_style_ids_joined is None:
        pattern_remove_irrelevant_attributes_for_empty_rows = re.compile(
            rb'\s*'
            rb'('
                rb'\bspans="[^"]+?"'
                rb'|'
                rb'\b[^\s:]*?:dyDescent="[^"]*"'
            rb')',      
            re.DOTALL
        )

        return pattern_remove_irrelevant_attributes_for_empty_rows


    pattern_remove_irrelevant_attributes_for_empty_rows = re.compile(
        rb'\s*'
        rb'('
            rb'\bspans="[^"]+?"'
            rb'|'
            rb'\b[^\s:]*?:dyDescent="[^"]*"'
            rb'|'
            rb'\bs="(?:' + only_font_style_ids_joined + rb')"'
        rb')',      
        re.DOTALL
    )

    return pattern_remove_irrelevant_attributes_for_empty_rows

def generate_only_font_style_ids_joined(xfs: List[XF]) -> bytes | None:
    if len(xfs) == 0:
        return None

    only_font_style_ids = find_no_style_but_font_ids(xfs)        

    only_font_style_bytestring_ids = [str(id).encode("ascii") for id in only_font_style_ids]
    only_font_style_ids_joined = b'|'.join(only_font_style_bytestring_ids)

    return only_font_style_ids_joined

def find_fill_or_border_ids(all_xfs: List[XF]) -> set[int]:
    return {
        i for i, xf in enumerate(all_xfs)
        if (xf.fill_id is not None and xf.fill_id != 0)
            or (xf.border_id is not None and xf.border_id != 0)
    }

def generate_fill_or_border_or_hidden_attribute_pattern(fill_or_border_bytestring_ids: List[bytes]) -> re.Pattern:
    fill_or_border_bytestring_ids_joined = b'|'.join(fill_or_border_bytestring_ids)

    fill_or_border_or_hidden_attribute_pattern = re.compile(
        rb'(?:s="(?:' + fill_or_border_bytestring_ids_joined + rb')"|"hidden="1")'
    )

    return fill_or_border_or_hidden_attribute_pattern

class WorkbookSpecificPatterns(NamedTuple):
    xfs_by_bytestring : dict[bytes, XF]
    pattern_remove_irrelevant_attributes_for_empty_rows : re.Pattern
    pattern_consecutive_only_font_styled_empty_cells : re.Pattern
    fill_or_border_or_hidden_attribute_pattern : re.Pattern
    whitespace_pattern : re.Pattern | None
    whitespace_cell_pattern : re.Pattern | None

def generate_workbook_specific_patterns(xfs: List[XF], whitespace_indices: List[int]) -> WorkbookSpecificPatterns:

    xfs_by_bytestring = {}
    for idx, xf in enumerate(xfs):
        key = str(idx).encode('ascii')
        xfs_by_bytestring[key] = xf

    fill_or_border_ids = find_fill_or_border_ids(xfs)
    fill_or_border_bytestring_ids = [str(id).encode("ascii") for id in fill_or_border_ids]

    only_font_style_ids_joined = generate_only_font_style_ids_joined(xfs)
    pattern_remove_irrelevant_attributes_for_empty_rows = generate_pattern_remove_irrelevant_attributes_for_empty_rows(only_font_style_ids_joined)
    pattern_consecutive_only_font_styled_empty_cells = generate_pattern_consecutive_only_font_styled_empty_cells(only_font_style_ids_joined)
    fill_or_border_or_hidden_attribute_pattern = generate_fill_or_border_or_hidden_attribute_pattern(fill_or_border_bytestring_ids)

    whitespace_joined = generate_whitespace_joined(whitespace_indices)
    whitespace_pattern = generate_whitespace_pattern(whitespace_joined)
    whitespace_cell_pattern = generate_whitespace_cell_pattern(whitespace_joined)

    return WorkbookSpecificPatterns(
        xfs_by_bytestring,
        pattern_remove_irrelevant_attributes_for_empty_rows,
        pattern_consecutive_only_font_styled_empty_cells,
        fill_or_border_or_hidden_attribute_pattern,
        whitespace_pattern,
        whitespace_cell_pattern
    )

