
from typing import Optional, Tuple


def extract_xml_block(
    xml_bytes: bytes,
    start_tag: bytes,
    end_tag: bytes,
    start_index: int = 0
) -> Tuple[Optional[bytes], int, int, int]:
    """
    Searches in xml_bytes for the region delimited by start_tag and end_tag,
    beginning at start_index. If found, returns a tuple of:
    
        (extracted_block, new_index, block_start, block_end)

    where:
      - extracted_block is the raw bytes between <start_tag> and </end_tag>.
      - new_index is the position right after the end_tag.
      - block_start is the offset in xml_bytes where extracted_block begins.
      - block_end is the offset in xml_bytes where extracted_block ends.

    If not found, returns (None, start_index, -1, -1).
    """
    # Find the start tag
    start_idx = xml_bytes.find(start_tag, start_index)
    if start_idx == -1:
        return None, start_index, -1, -1

    # Find the end of the start tag (the '>' character)
    start_tag_end = xml_bytes.find(b">", start_idx)
    if start_tag_end == -1:
        return None, start_index, -1, -1

    # Find the end tag
    end_idx = xml_bytes.find(end_tag, start_tag_end)
    if end_idx == -1:
        return None, start_index, -1, -1

    # The content is between (start_tag_end + 1) and end_idx
    block_start = start_tag_end + 1
    block_end   = end_idx
    extracted_data = xml_bytes[block_start:block_end]

    # Position after the end_tag
    new_search_index = end_idx + len(end_tag)

    return extracted_data, new_search_index, block_start, block_end
