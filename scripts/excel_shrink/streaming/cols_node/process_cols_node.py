from common.named_tupples import CleanSheetParameters, CollectColsResult, ColumnDefinition, ColumnRangeMap
from common.sheet_streaming_state import SheetStreamingState
from common.slice_bytes import find_and_slice_before_tag
import re

pattern_col_with_style = re.compile(
    rb'<col\b'
    rb'(?=[^>]*?\bstyle="([^"]+)")?'
    rb'(?=[^>]*?\bmin="([^"]+)")'
    rb'(?=[^>]*?\bmax="([^"]+)")'
    rb'(?=[^>]*?\bhidden="([^"]+)")?'
    rb'[^/>]*/>'
)

row_cache = {}

def digits_convert_byte_string_to_integer(row_bytes: bytes) -> int:
    """
    Convert row bytes (e.g. b'80') directly into an integer
    without decoding to string. This avoids overhead.
    """
    row_index = 0
    for b in row_bytes:
        # Digits 0-9: 48-57
        row_index = row_index * 10 + (b - 48)
    return row_index

def digits_byte_string_to_integer(row_b: bytes) -> int:
    """Return a cached or newly computed row index for row_b."""
    if row_b not in row_cache:
        row_cache[row_b] = digits_convert_byte_string_to_integer(row_b)
    return row_cache[row_b]



def extract_column_definitions_and_apply_fill_border_column_styles(col_data : bytes | None, params: CleanSheetParameters) -> ColumnRangeMap:
    if col_data is None:
        return ColumnRangeMap()
    
    column_definitions = ColumnRangeMap()
    apply_fill_border_styles = ColumnRangeMap()

    cols_style_matches = pattern_col_with_style.findall(col_data)

    for style_bytes, min_index_bytes, max_index_bytes, hidden in cols_style_matches:
        min_idx = digits_byte_string_to_integer(min_index_bytes)
        max_idx = digits_byte_string_to_integer(max_index_bytes)
        
        style = params.workbook_specific_patterns.xfs_by_bytestring[style_bytes] if style_bytes else None
        column_definitions.add(ColumnDefinition(min_idx, max_idx, style, hidden))

    column_definitions.finalize()

    return column_definitions

def process_cols_node(buffer: bytes, params : CleanSheetParameters) -> CollectColsResult | None:
    found_closing_tag = find_and_slice_before_tag(buffer, b"</cols>")
    if found_closing_tag is None:
        return None
    
    column_definitions = extract_column_definitions_and_apply_fill_border_column_styles(found_closing_tag.until_tag, params)

    return CollectColsResult(
        new_state=SheetStreamingState.IN_WORKSHEET_NODE,
        chunk_to_add=found_closing_tag.until_tag,
        updated_buffer=found_closing_tag.after_tag,
        column_definitions=column_definitions
    )