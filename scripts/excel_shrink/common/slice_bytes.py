from enum import Enum, auto
from typing import NamedTuple, Optional

class WhereToSlice(Enum):
    before_tag = auto()
    after_tag = auto()

class SliceAtTagResult(NamedTuple):
    until_tag: bytes
    after_tag: bytes
    is_self_closing_tag: bool = False

def find_and_slice_after_tag_range(buffer: bytes, tag_start, tag_end, search_backwards: bool = False) -> Optional[SliceAtTagResult]:
    if search_backwards:
        index_tag = buffer.rfind(tag_start)
    else:
        index_tag = buffer.find(tag_start)

    if index_tag == -1:
        return None
    
    after_tag_start = index_tag + len(tag_start)
    index_start_tag_end = buffer.find(tag_end, after_tag_start)
    if index_start_tag_end == -1:
        return None
    
    is_self_closing_tag = buffer[index_start_tag_end - 1] == b'/'[0]
    after_tag_end = index_start_tag_end + len(tag_end)  
   
    return SliceAtTagResult(
        until_tag= buffer[:after_tag_end],
        after_tag= buffer[after_tag_end:],
        is_self_closing_tag= is_self_closing_tag
    )

def find_and_slice_at_tag(buffer: bytes, tag : bytes, where_to_slice : WhereToSlice, search_backwards: bool = False) -> Optional[SliceAtTagResult]:
    if search_backwards:
        index_tag = buffer.rfind(tag)
    else:
        index_tag = buffer.find(tag)

    if index_tag == -1:
        return None
    
    if where_to_slice == WhereToSlice.before_tag:
        return SliceAtTagResult(
            until_tag= buffer[:index_tag],
            after_tag= buffer[index_tag:]
        )
    else:
        after_tag_name = index_tag + len(tag)
        return SliceAtTagResult(
            until_tag= buffer[:after_tag_name],
            after_tag= buffer[after_tag_name:]
        )

def find_and_slice_before_tag(buffer: bytes, tag_start_pattern, search_backwards: bool = False) -> Optional[SliceAtTagResult]:
    return find_and_slice_at_tag(buffer, tag_start_pattern, WhereToSlice.before_tag, search_backwards)


def find_and_slice_after_tag(buffer: bytes, tag_start_pattern, search_backwards: bool = False) -> Optional[SliceAtTagResult]:
    return find_and_slice_at_tag(buffer, tag_start_pattern, WhereToSlice.after_tag, search_backwards)