from common.named_tupples import (
    CleanSheetDataResult,
    CleanSheetParameters,
    XF,
    CellCoordinate,
)
from common.sheet_streaming_state import SheetStreamingState
from common.slice_bytes import SliceAtTagResult, find_and_slice_before_tag
from cleaning.hidden_rows import remove_hidden_rows
import re
from cleaning.common import CleanSheetMetrics, extract_cell_attributes
from cleaning.generate_and_log_rows_and_cells import (
    generate_and_log_deleted_cell,
    generate_and_log_self_closing_cell,
)
from common.letters_byte_string_to_integer import letters_byte_string_to_integer

measure_remove_consecutive_hidden_self_closing_rows = False


def find_sheet_data_chunk(buffer: bytes) -> SliceAtTagResult | None:
    found_last_row_start_tag = find_and_slice_before_tag(
        buffer, b"<row", search_backwards=True
    )
    if found_last_row_start_tag:
        return found_last_row_start_tag

    return None


class CDataError(ValueError):
    pass


def pre_process_sheet_data_chunk(chunck: bytes, params: CleanSheetParameters) -> bytes:
    if b"<![CDATA[" in chunck:
        raise CDataError("CDATA sections are not supported in sheet data chunk.")

    chunck = remove_hidden_rows_if_present(chunck, params)

    chunck = replace_whitespace_cells(chunck, params)

    chunck = remove_consecutive_empty_cells(chunck, params)

    return chunck


def remove_hidden_rows_if_present(chunck, params):
    if b'hidden="1"' in chunck:
        chunck = remove_hidden_rows(chunck, params)
    return chunck


def remove_consecutive_empty_cells(chunck, params):
    if b"<c" in chunck:
        chunck = params.workbook_specific_patterns.pattern_consecutive_only_font_styled_empty_cells.sub(
            b"", chunck
        )
    return chunck


def replace_whitespace_cells(chunck, params):
    whitespace_pattern = params.workbook_specific_patterns.whitespace_pattern
    whitespace_cell_pattern = params.workbook_specific_patterns.whitespace_cell_pattern
    if whitespace_pattern and whitespace_cell_pattern:
        if whitespace_pattern.search(chunck):
            chunck = whitespace_cell_pattern.sub(rb"\1/>", chunck)
    return chunck


def cell_is_hidden(cell_coordinate: CellCoordinate, metrics: CleanSheetMetrics) -> bool:

    row_style = metrics.row_attributes.get(cell_coordinate.row, None)
    if row_style and row_style.hidden:
        return True

    column_index = letters_byte_string_to_integer(cell_coordinate.column)
    column = metrics.column_definitions.get(column_index)
    if column and column.hidden:
        return True

    return False


def cell_overrides_fill_or_border(
    cell_style: XF, cell_coordinate: CellCoordinate, metrics: CleanSheetMetrics
) -> bool:
    column_index = letters_byte_string_to_integer(cell_coordinate.column)

    column = metrics.column_definitions.get(column_index, None)
    column_style = column.cell_xf if column is not None else None

    row_attributes = metrics.row_attributes.get(cell_coordinate.row, None)
    row_style = row_attributes.style if row_attributes is not None else None
    overrides_fill_or_border = (
        (row_style and row_style.apply_fill and cell_style.fill_id != row_style.fill_id)
        or (
            column_style
            and column_style.apply_fill
            and cell_style.fill_id != column_style.fill_id
        )
    ) or (
        (
            row_style
            and row_style.apply_border
            and cell_style.border_id != row_style.border_id
        )
        or (
            column_style
            and column_style.apply_border
            and cell_style.border_id != column_style.border_id
        )
    )

    if overrides_fill_or_border:
        return True
    else:
        return False


def replace_cell(
    match: re.Match, params: CleanSheetParameters, metrics: CleanSheetMetrics
) -> bytes:
    cell_attributes = extract_cell_attributes(match, params)
    cell_coordinate, style = cell_attributes.cell_coordinate, cell_attributes.style

    if cell_is_hidden(cell_coordinate, metrics):
        return generate_and_log_deleted_cell(metrics, cell_attributes)

    if not style:
        return generate_and_log_deleted_cell(metrics, cell_attributes)

    if style.has_fill_or_border or style.has_apply_protection:
        return generate_and_log_self_closing_cell(metrics, cell_attributes)

    if not style.has_apply_fill_or_border:
        return generate_and_log_deleted_cell(metrics, cell_attributes)

    if cell_overrides_fill_or_border(style, cell_coordinate, metrics):
        return generate_and_log_self_closing_cell(metrics, cell_attributes)

    return generate_and_log_deleted_cell(metrics, cell_attributes)


def process_sheet_data_node(
    buffer: bytes, params: CleanSheetParameters
) -> CleanSheetDataResult | None:
    found_closing_sheet_data_tag = find_and_slice_before_tag(buffer, b"</sheetData>")
    if found_closing_sheet_data_tag:

        return CleanSheetDataResult(
            new_state=SheetStreamingState.IN_WORKSHEET_NODE,
            chunk_to_add=pre_process_sheet_data_chunk(
                found_closing_sheet_data_tag.until_tag, params
            ),
            updated_buffer=found_closing_sheet_data_tag.after_tag,
        )

    find_sheet_data_chunk_result = find_sheet_data_chunk(buffer)
    if not find_sheet_data_chunk_result:
        return None

    return CleanSheetDataResult(
        new_state=SheetStreamingState.IN_SHEET_DATA_NODE,
        chunk_to_add=pre_process_sheet_data_chunk(
            find_sheet_data_chunk_result.until_tag, params
        ),
        updated_buffer=find_sheet_data_chunk_result.after_tag,
    )
