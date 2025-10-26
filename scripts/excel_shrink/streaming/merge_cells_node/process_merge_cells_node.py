
from common.named_tupples import MergeCellsParsingResult
from common.sheet_streaming_state import SheetStreamingState
from common.slice_bytes import find_and_slice_before_tag

import re

from common.named_tupples import CellCoordinate

mergecell_pattern = re.compile(
    rb'<mergeCell\b'
    rb'[^/>]*\bref="([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)"'
    rb'[^/>]*/>'
)

def find_max_merge_cells_column_and_row_bytestrings(merge_cell_data : bytes) -> CellCoordinate | None:
    merge_cell_matches = mergecell_pattern.findall(merge_cell_data)

    if len(merge_cell_matches) == 0:
        return None

    distinct_columns = {m[0] for m in merge_cell_matches} | {m[2] for m in merge_cell_matches}
    distinct_rows = {m[1] for m in merge_cell_matches} | {m[3] for m in merge_cell_matches}

    max_merge_column = max(distinct_columns, key=lambda col: (len(col), col))
    max_merge_row = max(distinct_rows, key=lambda row: (len(row), row))

    return CellCoordinate(max_merge_column, max_merge_row)



def process_merge_cells_node(buffer: bytes) -> MergeCellsParsingResult | None:
    found_closing_tag = find_and_slice_before_tag(buffer, b"</mergeCells>")
    if found_closing_tag is None:
        return None
    
    max_merge_cell = find_max_merge_cells_column_and_row_bytestrings(found_closing_tag.until_tag)

    return MergeCellsParsingResult(
        new_state=SheetStreamingState.IN_WORKSHEET_NODE,
        chunk_to_add=found_closing_tag.until_tag,
        updated_buffer=found_closing_tag.after_tag,
        max_col_and_row=max_merge_cell,
    )

