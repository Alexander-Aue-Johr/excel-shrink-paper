from typing import Final, List, Set
from common.named_tupples import CollectChunksResult, TagToProcess
from common.sheet_streaming_state import SheetStreamingState
from common.slice_bytes import find_and_slice_after_tag, find_and_slice_after_tag_range

cols_tag : Final[TagToProcess] = TagToProcess(b"<cols", tag_end=b'>')
sheet_data_tag : Final[TagToProcess] = TagToProcess(b"<sheetData", tag_end=b'>')
merge_cells_tag : Final[TagToProcess] = TagToProcess(b"<mergeCells", tag_end=b'>')
worksheet_end_tag : Final[TagToProcess] = TagToProcess(b"</worksheet>")

def process_worksheet_node(
    buffer: bytes, 
    state: SheetStreamingState,
    processed_tags: Set[bytes]
) -> CollectChunksResult:
    if state != SheetStreamingState.IN_WORKSHEET_NODE:
        raise ValueError(
            f"Invalid state {state} for collect_chunks_until_child_node_is_found"
        )
    
    if not buffer:
        raise ValueError(
            "Buffer is empty. Cannot process tags."
        )
    

    tags_to_process_in_order : List[TagToProcess] = [
        cols_tag, sheet_data_tag, merge_cells_tag, worksheet_end_tag
    ]
    remaining_tags_to_process : List[TagToProcess] = [tag for tag in tags_to_process_in_order if tag not in processed_tags]
    for tag in remaining_tags_to_process:
        
        if not tag.tag_end:
            found_tag = find_and_slice_after_tag(buffer, tag.tag_start)
        else:
            found_tag = find_and_slice_after_tag_range(buffer, tag.tag_start, tag.tag_end)
        
        if found_tag:
            if found_tag.is_self_closing_tag:
                new_state= SheetStreamingState.IN_WORKSHEET_NODE
            else:
                new_state= SheetStreamingState.IN_COLS_NODE if tag == cols_tag else \
                            SheetStreamingState.IN_SHEET_DATA_NODE if tag == sheet_data_tag else \
                            SheetStreamingState.IN_MERGECELLS_NODE if tag == merge_cells_tag else \
                            SheetStreamingState.FOUND_WORKSHEET_END if tag == worksheet_end_tag else \
                            None
            
            if new_state is None:
                raise ValueError(
                    f"Invalid tag {tag} found in worksheet node."
                )
            return CollectChunksResult(
                new_state=new_state,
                chunk_to_add=found_tag.until_tag,
                updated_buffer=found_tag.after_tag,
                processed_tag=tag
            )
        
    return CollectChunksResult(
        new_state=None,
    )