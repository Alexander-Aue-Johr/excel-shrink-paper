import re
from typing import List

t_node_content_capture_group = re.compile(rb"<t[^>]*>(.*?)</t>", re.DOTALL)


def find_whitespace_sst_indices(si_matches: list) -> List[int]:
    """
    Given the list of <si> blocks, find which indices are purely whitespace.
    """
    whitespace_indices = []
    for idx, si_block in enumerate(si_matches):
        # capture text from <t>...</t>
        t_open_close = t_node_content_capture_group.findall(si_block)

        text_parts = []
        for part in t_open_close:
            trimmed = part.strip()
            if trimmed:
                text_parts.append(trimmed)

        combined_text = b"".join(text_parts).strip()
        if len(combined_text) == 0:
            whitespace_indices.append(idx)

    return whitespace_indices
