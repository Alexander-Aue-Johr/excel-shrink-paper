
from typing import NamedTuple
from cleaning.common import CleanSheetMetrics
from common.named_tupples import CellCoordinate, CleanSheetParameters, ColumnRangeMap
from common.sheet_streaming_state import SheetStreamingState

from streaming.cols_node.process_cols_node import process_cols_node
from streaming.merge_cells_node.process_merge_cells_node import process_merge_cells_node
from streaming.sheet_data_node.process_sheet_data_node import process_sheet_data_node
from streaming.worksheet_node.process_worksheet_node import process_worksheet_node

class PreprocessedSheetXml(NamedTuple):
    before_cols_node_output_chunks : bytes
    in_cols_node_output_chunks : bytes
    before_sheet_data_node_output_chunks : bytes
    sheet_data : bytes
    before_merge_cells_node_output_chunks : bytes
    in_merge_cells_node_output_chunks : bytes
    until_worksheet_end_output_chunks : bytes
    clean_sheet_metrics : CleanSheetMetrics

def preprocess_sheet_xml_streamwise(sheet_stream, params : CleanSheetParameters, chunk_size : int):
    buffer = b""

    before_cols_node_output_chunks : bytes = b''
    in_cols_node_output_chunks : bytes = b''
    before_sheet_data_node_output_chunks : bytes = b''
    in_sheet_data_node_output_chunks : bytes = b''
    before_merge_cells_node_output_chunks : bytes = b''
    in_merge_cells_node_output_chunks : bytes = b''
    until_worksheet_end_output_chunks : bytes = b''
    max_merge_cells_column_and_row : CellCoordinate | None = None
    
    sheet_streaming_state = SheetStreamingState.IN_WORKSHEET_NODE
    processed_tags = set()

    column_definitions : ColumnRangeMap | None = None


    while sheet_streaming_state != SheetStreamingState.FOUND_WORKSHEET_END:
        chunk = sheet_stream.read(chunk_size)
        buffer += chunk
        match sheet_streaming_state:
            case SheetStreamingState.IN_WORKSHEET_NODE:
                collect_chunks_result = process_worksheet_node(buffer, sheet_streaming_state, processed_tags)
                if collect_chunks_result.new_state is None:
                    continue
                buffer = collect_chunks_result.updated_buffer
                processed_tags.add(collect_chunks_result.processed_tag)
                if collect_chunks_result.chunk_to_add is None:
                    raise ValueError(
                        "Chunk to add is None. This should not happen."
                    )
                match collect_chunks_result.new_state:
                    case SheetStreamingState.IN_COLS_NODE:
                        before_cols_node_output_chunks = collect_chunks_result.chunk_to_add
                    case SheetStreamingState.IN_SHEET_DATA_NODE:
                        before_sheet_data_node_output_chunks = collect_chunks_result.chunk_to_add
                    case SheetStreamingState.IN_MERGECELLS_NODE:
                        before_merge_cells_node_output_chunks = collect_chunks_result.chunk_to_add
                    case SheetStreamingState.FOUND_WORKSHEET_END:
                        until_worksheet_end_output_chunks = collect_chunks_result.chunk_to_add
                sheet_streaming_state = collect_chunks_result.new_state
                continue

            case SheetStreamingState.IN_COLS_NODE:
                result = process_cols_node(buffer, params)
                if result is None:
                    continue

                in_cols_node_output_chunks += result.chunk_to_add
                buffer = result.updated_buffer
                sheet_streaming_state = result.new_state
                column_definitions = result.column_definitions

                continue

            case SheetStreamingState.IN_SHEET_DATA_NODE:

                result = process_sheet_data_node(buffer, params)
                if result is None:
                    continue
                    
                in_sheet_data_node_output_chunks += result.chunk_to_add
                buffer = result.updated_buffer
                sheet_streaming_state = result.new_state
                continue

            case SheetStreamingState.IN_MERGECELLS_NODE:
                result = process_merge_cells_node(buffer)
                if result is None:
                    continue

                in_merge_cells_node_output_chunks += result.chunk_to_add
                buffer = result.updated_buffer
                max_merge_cells_column_and_row = result.max_col_and_row
                sheet_streaming_state = result.new_state
                continue
            case _:
                break
    

    clean_sheet_metrics = CleanSheetMetrics(
        all_row_bytestrings_of_last_cell_in_row=set(),
        all_col_bytestrings_of_last_cell_in_row=set(),
        kept_row_states=[],
        deleted_row_states=[],
        kept_cell_states=[],
        deleted_cell_states=[],
        apply_fill_border_styles_by_row={},
        column_definitions=column_definitions if column_definitions else ColumnRangeMap(),
        max_merge_cells_column_and_row=max_merge_cells_column_and_row,
        row_attributes=dict(),
    )

    preprocessed_sheet_xml = PreprocessedSheetXml(
        before_cols_node_output_chunks,
        in_cols_node_output_chunks,
        before_sheet_data_node_output_chunks,
        in_sheet_data_node_output_chunks,
        before_merge_cells_node_output_chunks,
        in_merge_cells_node_output_chunks,
        until_worksheet_end_output_chunks,
        clean_sheet_metrics
    )

    return preprocessed_sheet_xml