import logging
import re

common_col_pattern = rb"<col\b" rb'(?=[^/>]*(\bmax="([^"]+?)"))' rb"[^/>]*/>"

pattern_col = re.compile(common_col_pattern)


def collect_col_chunks_until_first_extraneous_column(
    cols_content: bytes, last_max_col: int
) -> tuple[re.Match[bytes] | None, bytes]:
    col_chunks: bytes = b""

    for match in pattern_col.finditer(cols_content):
        max_raw = match.group(2)

        col_max = int(max_raw)

        if col_max <= last_max_col:
            col_chunks += match.group(0)
            continue
        else:
            return match, col_chunks

    return None, col_chunks


pattern_last_cols = re.compile(common_col_pattern + rb"$")


def merge_extraneous_columns(cols_content: bytes, last_max_col: int | None) -> bytes:
    if last_max_col is None:
        logging.debug("no last_max_col")
        return cols_content

    first_extraneous_column_match, col_chunks = (
        collect_col_chunks_until_first_extraneous_column(cols_content, last_max_col)
    )

    if not first_extraneous_column_match:
        logging.debug("no extraneous columns found")
        return cols_content

    last_column_match = pattern_last_cols.search(
        cols_content, first_extraneous_column_match.end()
    )
    if not last_column_match:
        logging.debug("the last column already is the first extraneous column")
        return cols_content

    first_extraneous_column_max_attribute = first_extraneous_column_match.group(2)
    first_extraneous_column_max_attribute_string = (
        b'max="' + first_extraneous_column_max_attribute + b'"'
    )

    last_column_max_attribute_string = b'max="' + last_column_match.group(2) + b'"'

    first_extraneous_column_string = first_extraneous_column_match.group(0)
    updated_first_extraneous_column_string = first_extraneous_column_string.replace(
        first_extraneous_column_max_attribute_string,
        last_column_max_attribute_string,
        1,
    )

    new_cols = col_chunks + updated_first_extraneous_column_string
    return new_cols
