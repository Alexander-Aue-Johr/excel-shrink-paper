import logging
import re
from typing import Final, NamedTuple, Optional, Set
from cleaning.common import CleanSheetMetrics
from common.named_tupples import (
    CellCoordinate,
    CleanSheetParameters,
    CollectChunksResult,
    ColumnRangeMap,
    TagToProcess,
)
from common.sheet_streaming_state import SheetStreamingState

from common.slice_bytes import find_and_slice_after_tag_range
from streaming.cols_node.process_cols_node import process_cols_node
from streaming.merge_cells_node.process_merge_cells_node import process_merge_cells_node
from streaming.sheet_data_node.process_sheet_data_node import process_sheet_data_node
from streaming.worksheet_node.process_worksheet_node import process_worksheet_node


class PreprocessedSheetXml(NamedTuple):
    until_worksheet_start_output_chunks: bytes
    before_cols_node_output_chunks: bytes
    in_cols_node_output_chunks: bytes
    before_sheet_data_node_output_chunks: bytes
    sheet_data: bytes
    before_merge_cells_node_output_chunks: bytes
    in_merge_cells_node_output_chunks: bytes
    until_worksheet_end_output_chunks: bytes
    clean_sheet_metrics: CleanSheetMetrics


worksheet_tag: Final[TagToProcess] = TagToProcess(b"<worksheet", tag_end=b">")

root_start_pattern = re.compile(
    r"""
    ^\s*(?:\ufeff)?                              # optional BOM
    (?:<\?xml\b[^?]*\?>\s*)?                      # optional XML declaration
    (?:<\?[^?]*\?>\s*|<!--[\s\S]*?-->\s*|<!DOCTYPE[\s\S]*?>\s*)*  # optional PI/comments/doctype
    <\s*(?:(?P<prefix>[A-Za-z_][\w.-]*)\s*:\s*)?(?P<local>[A-Za-z_][\w.-]*)\b
    (?P<attrs>[\s\S]*?)>
    """,
    re.VERBOSE | re.DOTALL,
)


def _try_parse_root_start_tag(
    buffer: bytes, max_decode: int = 64 * 1024
) -> Optional[tuple[str, str, str]]:
    """
    Returns (prefix, local, attrs) or None if we can't parse yet.
    Decodes only up to max_decode bytes (UTF-8, ignore errors).
    """
    # Only look at a prefix of the buffer to stay cheap.
    head = buffer[:max_decode].decode("utf-8", errors="ignore")
    m = root_start_pattern.match(head)
    if not m:
        return None
    prefix = m.group("prefix") or ""
    local = m.group("local")
    attrs = m.group("attrs") or ""
    return prefix, local, attrs


def seek_worksheet_start(
    buffer: bytes, state: SheetStreamingState, params: CleanSheetParameters
) -> CollectChunksResult:

    if state != SheetStreamingState.SEEK_WORKSHEET_NODE:
        raise ValueError(f"Invalid state {state} for seek_worksheet_start")

    if not buffer:
        raise ValueError("Buffer is empty. Cannot process tags.")

    found_tag = find_and_slice_after_tag_range(
        buffer, worksheet_tag.tag_start, worksheet_tag.tag_end
    )

    if found_tag:
        return CollectChunksResult(
            new_state=SheetStreamingState.IN_WORKSHEET_NODE,
            chunk_to_add=found_tag.until_tag,
            updated_buffer=found_tag.after_tag,
            processed_tag=worksheet_tag,
        )

    # --- Fallback: root worksheet node with optional namespace prefix ---
    parsed = _try_parse_root_start_tag(buffer)
    if parsed is not None:
        prefix, local, _attrs = parsed

        if local != "worksheet" or prefix != "":
            logging.error(
                f"Found invalid root node with local='{local}' and prefix='{prefix}' in sheet {params.zip_info.filename} in {params.xlsx_file_path}."
            )
            raise ValueError("Found invalid root node.")

    return CollectChunksResult(
        new_state=None,
    )


def preprocess_sheet_xml_streamwise(
    sheet_stream, params: CleanSheetParameters, chunk_size: int
) -> PreprocessedSheetXml:
    buffer: bytes = b""

    before_cols_node_output_chunks: bytes = b""
    in_cols_node_output_chunks: bytes = b""
    before_sheet_data_node_output_chunks: bytes = b""
    in_sheet_data_node_output_chunks: bytes = b""
    before_merge_cells_node_output_chunks: bytes = b""
    in_merge_cells_node_output_chunks: bytes = b""
    until_worksheet_end_output_chunks: bytes = b""
    max_merge_cells_column_and_row: CellCoordinate | None = None

    sheet_streaming_state = SheetStreamingState.SEEK_WORKSHEET_NODE
    processed_tags = set()
    reached_end_of_file: bool = False

    column_definitions: ColumnRangeMap | None = None

    while sheet_streaming_state != SheetStreamingState.FOUND_WORKSHEET_END:
        chunk = sheet_stream.read(chunk_size)
        if chunk == b"":
            reached_end_of_file = True

        buffer += chunk
        match sheet_streaming_state:
            case SheetStreamingState.SEEK_WORKSHEET_NODE:
                collect_chunks_result = seek_worksheet_start(
                    buffer, sheet_streaming_state, params
                )
                if collect_chunks_result.new_state is None:
                    if reached_end_of_file:
                        raise ValueError(
                            "Reached end of file while processing worksheet node."
                        )
                    continue

                if (
                    collect_chunks_result.new_state
                    is SheetStreamingState.IN_WORKSHEET_NODE
                ):
                    until_worksheet_start_output_chunks = (
                        collect_chunks_result.chunk_to_add
                    )
                    sheet_streaming_state = collect_chunks_result.new_state
                    buffer = collect_chunks_result.updated_buffer

                continue

            case SheetStreamingState.IN_WORKSHEET_NODE:
                collect_chunks_result = process_worksheet_node(
                    buffer, sheet_streaming_state, processed_tags
                )
                if collect_chunks_result.new_state is None:
                    continue
                buffer = collect_chunks_result.updated_buffer
                processed_tags.add(collect_chunks_result.processed_tag)
                if collect_chunks_result.chunk_to_add is None:
                    raise ValueError("Chunk to add is None. This should not happen.")
                match collect_chunks_result.new_state:
                    case SheetStreamingState.IN_COLS_NODE:
                        before_cols_node_output_chunks = (
                            collect_chunks_result.chunk_to_add
                        )
                    case SheetStreamingState.IN_SHEET_DATA_NODE:
                        before_sheet_data_node_output_chunks = (
                            collect_chunks_result.chunk_to_add
                        )
                    case SheetStreamingState.IN_MERGECELLS_NODE:
                        before_merge_cells_node_output_chunks = (
                            collect_chunks_result.chunk_to_add
                        )
                    case SheetStreamingState.FOUND_WORKSHEET_END:
                        until_worksheet_end_output_chunks = (
                            collect_chunks_result.chunk_to_add
                        )
                sheet_streaming_state = collect_chunks_result.new_state
                continue

            case SheetStreamingState.IN_COLS_NODE:
                result = process_cols_node(buffer, params)
                if result is None:
                    if reached_end_of_file:
                        raise ValueError(
                            "Reached end of file while processing cols node."
                        )
                    continue

                in_cols_node_output_chunks += result.chunk_to_add
                buffer = result.updated_buffer
                sheet_streaming_state = result.new_state
                column_definitions = result.column_definitions

                continue

            case SheetStreamingState.IN_SHEET_DATA_NODE:

                result = process_sheet_data_node(buffer, params)
                if result is None:
                    if reached_end_of_file:
                        raise ValueError(
                            "Reached end of file while processing sheet data node."
                        )
                    continue

                in_sheet_data_node_output_chunks += result.chunk_to_add
                buffer = result.updated_buffer
                sheet_streaming_state = result.new_state
                continue

            case SheetStreamingState.IN_MERGECELLS_NODE:
                result = process_merge_cells_node(buffer)
                if result is None:
                    if reached_end_of_file:
                        raise ValueError(
                            "Reached end of file while processing merge cells node."
                        )
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
        column_definitions=(
            column_definitions if column_definitions else ColumnRangeMap()
        ),
        max_merge_cells_column_and_row=max_merge_cells_column_and_row,
        row_attributes=dict(),
    )

    preprocessed_sheet_xml = PreprocessedSheetXml(
        until_worksheet_start_output_chunks,
        before_cols_node_output_chunks,
        in_cols_node_output_chunks,
        before_sheet_data_node_output_chunks,
        in_sheet_data_node_output_chunks,
        before_merge_cells_node_output_chunks,
        in_merge_cells_node_output_chunks,
        until_worksheet_end_output_chunks,
        clean_sheet_metrics,
    )

    return preprocessed_sheet_xml
