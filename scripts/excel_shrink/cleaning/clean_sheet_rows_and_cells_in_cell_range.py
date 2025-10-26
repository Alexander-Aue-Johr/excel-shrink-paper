import re
from cleaning.common import CleanSheetMetrics
from cleaning.replace_row import replace_row
from common.apply_pattern_substitution import apply_pattern_substitution
from common.named_tupples import CleanSheetParameters

pattern_process_rows = re.compile(
    rb'<row\b'
        rb'('
            rb'(?:[^>]*\br="([^"]+?)")?'
            rb'(?:[^>]*\bhidden="(1)")?'
            rb'(?:[^>]*\bs="([^"]+?)")?'
            rb'(?:[^>]*\bht="([^"]+?)")?'
            rb'[^/>]*?'
        rb')'
    rb'(?:'      
        rb'/>'   
        rb'|'
        rb'>'
        rb'(.*?)'
        rb'</row>'
    rb')',      
    re.DOTALL
)

def clean_sheet_rows_and_cells_in_cell_range(sheet_data : bytes, params: CleanSheetParameters, clean_sheet_metrics : CleanSheetMetrics) -> bytes:
           
    sheet_data = apply_pattern_substitution(
        sheet_data,
        pattern_process_rows,
        lambda match: replace_row(match, params, clean_sheet_metrics),
        "pattern_process_rows + replace_row",
        params
    )

    return sheet_data