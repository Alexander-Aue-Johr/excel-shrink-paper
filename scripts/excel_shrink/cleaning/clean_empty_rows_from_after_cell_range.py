import re
from cleaning.common import CleanSheetMetrics, extract_empty_row_attributes, previous_row_was_also_hidden
from common.apply_pattern_substitution import apply_pattern_substitution
from common.named_tupples import CleanSheetParameters
from cleaning.generate_and_log_rows_and_cells import generate_and_log_deleted_row, generate_and_log_self_closing_row


pattern_process_empty_rows_ignoring_height = re.compile(
    rb'<row\b'
        rb'('
            rb'(?:[^>]*\br="([^"]+?)")?'
            rb'(?:[^>]*\bhidden="(1)")?'
            rb'(?:[^>]*\bs="([^"]+?)")?'
            rb'[^/>]*?'
        rb')'
    rb'(?:'      
        rb'/>'   
        rb'|'
        rb'></row>'
    rb')',      
    re.DOTALL
)

def replace_empty_row(match: re.Match, params : CleanSheetParameters, metrics : CleanSheetMetrics) -> bytes:
    row_attributes = extract_empty_row_attributes(match, params)   
    row_attr_str, ref, is_hidden, row_style, _, _ = row_attributes
    metrics.row_attributes[ref] = row_attributes

    if is_hidden: 
        if previous_row_was_also_hidden(metrics):
            return generate_and_log_deleted_row(metrics, row_attributes)
        else:
            return generate_and_log_self_closing_row(row_attr_str, params, metrics, row_attributes)

    if (not row_style or not row_style.has_fill_or_border):
        return generate_and_log_deleted_row(metrics, row_attributes)

    return generate_and_log_self_closing_row(row_attr_str, params, metrics, row_attributes)

def clean_empty_rows_from_after_cell_range(sheet_data : bytes, params: CleanSheetParameters, clean_sheet_metrics : CleanSheetMetrics) -> bytes:
    
    sheet_data = apply_pattern_substitution(
        sheet_data,
        pattern_process_empty_rows_ignoring_height,
        lambda match: replace_empty_row(match, params, clean_sheet_metrics),
        "pattern_process_empty_rows_after_cell_removal + replace_row",
        params
    )

    return sheet_data