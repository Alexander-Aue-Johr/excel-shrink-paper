

import re
from cleaning.common import CleanSheetMetrics, extract_row_attributes, previous_row_was_also_hidden
from cleaning.generate_and_log_rows_and_cells import generate_and_log_row
from cleaning.generate_and_log_rows_and_cells import generate_and_log_deleted_row, generate_and_log_self_closing_row
from common.apply_pattern_substitution import apply_pattern_substitution
from common.named_tupples import CleanSheetParameters
from streaming.sheet_data_node.process_sheet_data_node import replace_cell


pattern_process_cells = re.compile(
    rb'<c\b'
        rb'('
            rb'(?:[^>]*\br="([A-Z]+)(\d+)")?'
            rb'(?:[^>]*\bs="(\d+?)")?'
            rb'[^/>]*?'
        rb')'
        rb'/>'
    , re.DOTALL
)

def clean_cells(params, metrics, row_content):
    if row_content:
        row_content = apply_pattern_substitution(
            row_content,
            pattern_process_cells,
            lambda match: replace_cell(match, params, metrics),
            "pattern_process_cells + replace_cell",
            params
        )
        
    return row_content

def add_custom_height_if_absent(row_attr_str: bytes):
    if b' customHeight="1"' in row_attr_str:
        return row_attr_str
    
    if b' ht="' not in row_attr_str:
        return row_attr_str

    row_attr_str += b' customHeight="1"'

    return row_attr_str

def replace_row(match: re.Match, params : CleanSheetParameters, metrics : CleanSheetMetrics) -> bytes:
    row_attributes, row_content = extract_row_attributes(match, params)   
    row_attr_str, ref, is_hidden, row_style, has_height, is_empty = row_attributes
    metrics.row_attributes[ref] = row_attributes

    was_already_empty = not row_content
    if not was_already_empty:
        row_content = clean_cells(params, metrics, row_content)
    is_empty = not row_content 
    became_empty_by_cleaning = not was_already_empty and is_empty

    if is_empty:
        if is_hidden: 
            if previous_row_was_also_hidden(metrics):
                return generate_and_log_deleted_row(metrics, row_attributes)
            else:
                return generate_and_log_self_closing_row(row_attr_str, params, metrics, row_attributes)

        if (not row_style or not row_style.has_fill_or_border) and not has_height:
            return generate_and_log_deleted_row(metrics, row_attributes)

        if became_empty_by_cleaning:
            add_custom_height_if_absent(row_attr_str)
        return generate_and_log_self_closing_row(row_attr_str, params, metrics, row_attributes)
        

        
    return generate_and_log_row(row_attr_str, row_content, metrics, row_attributes)