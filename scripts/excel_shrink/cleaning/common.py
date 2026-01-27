from enum import Enum, auto
import re
from typing import List, NamedTuple, Set, Tuple

from common.named_tupples import (
    XF,
    CellCoordinate,
    CleanSheetParameters,
    ColumnRangeMap,
)


class RowAttributes(NamedTuple):
    row_attr: bytes
    ref: bytes
    hidden: bool
    style: XF | None
    has_height: bool
    is_empty: bool


class CellAttributes(NamedTuple):
    cell_attribute_string: bytes
    cell_coordinate: CellCoordinate
    style: XF | None


class RowState(Enum):
    empty = auto()
    hidden = auto()
    has_height = auto()
    has_style = auto()
    has_fill_or_border = auto()


class CellState(Enum):
    empty = auto()
    has_style = auto()
    has_fill_or_border = auto()
    overwrites_fill_or_border = auto()


class CleanSheetMetrics(NamedTuple):
    all_col_bytestrings_of_last_cell_in_row: Set[bytes]
    all_row_bytestrings_of_last_cell_in_row: Set[bytes]
    column_definitions: ColumnRangeMap
    apply_fill_border_styles_by_row: dict[int, Set[bytes]]
    kept_row_states: List[RowAttributes]
    deleted_row_states: List[RowAttributes]
    kept_cell_states: List[CellAttributes]
    deleted_cell_states: List[CellAttributes]
    max_merge_cells_column_and_row: CellCoordinate | None
    row_attributes: dict[bytes, RowAttributes]


def extract_empty_row_attributes(
    match: re.Match, params: CleanSheetParameters
) -> RowAttributes:
    row_attr, ref, hidden, style_id_bytes = match.groups()

    is_hidden = hidden is not None
    has_style = style_id_bytes is not None

    if has_style:
        style = params.workbook_specific_patterns.xfs_by_bytestring.get(style_id_bytes)
    else:
        style = None

    return RowAttributes(row_attr, ref, is_hidden, style, False, True)


def extract_row_attributes(
    match: re.Match, params: CleanSheetParameters
) -> Tuple[RowAttributes, bytes]:
    row_attr, ref, hidden, style_id_bytes, ht, original_row_content = match.groups()

    is_hidden = hidden is not None
    has_style = style_id_bytes is not None

    if has_style:
        style = params.workbook_specific_patterns.xfs_by_bytestring.get(style_id_bytes)
    else:
        style = None

    has_height = ht is not None
    is_empty = not original_row_content

    return (
        RowAttributes(row_attr, ref, is_hidden, style, has_height, is_empty),
        original_row_content,
    )


def previous_row_was_also_hidden(clean_sheet_metrics: CleanSheetMetrics):
    if not clean_sheet_metrics.kept_row_states:
        return False

    previous_row = clean_sheet_metrics.kept_row_states[-1]

    return previous_row.hidden


def extract_cell_attributes(
    match: re.Match, params: CleanSheetParameters
) -> CellAttributes:
    cell_attr, col_ref, row_ref, style_id = match.groups()

    style = (
        params.workbook_specific_patterns.xfs_by_bytestring[style_id]
        if style_id
        else None
    )

    return CellAttributes(cell_attr, CellCoordinate(col_ref, row_ref), style)
