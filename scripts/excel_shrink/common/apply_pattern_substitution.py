import re
import time
from typing import Callable, Union
import logging

from common.named_tupples import CleanSheetParameters, CleanWorkbookParameters


def apply_pattern_substitution(
    data: bytes,
    pattern: re.Pattern,
    replacement: Union[bytes, Callable[[re.Match], bytes]],
    description: str,
    params: CleanSheetParameters | CleanWorkbookParameters,
    duration_logging_threshold: float = 5.0,
) -> bytes:
    old_len = len(data)
    start_time = time.time()
    new_data = pattern.sub(replacement, data)
    end_time = time.time()
    difference = old_len - len(new_data)
    reduction_percent = (difference / old_len * 100) if old_len > 0 else 0.0
    duration = end_time - start_time

    if duration > duration_logging_threshold:
        logging.info(
            f"{description}: reduced from {old_len} to {len(new_data)} by {difference} bytes, reduction {reduction_percent:.4f}% in {duration:.4f} seconds, in file {params.xlsx_file_path} : {params.zip_info.filename}"
        )
    return new_data
