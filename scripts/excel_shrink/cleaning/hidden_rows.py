
import re

from common.apply_pattern_substitution import apply_pattern_substitution
from common.named_tupples import CleanSheetParameters

try_catastrophic_backtracking = False

pattern_hidden_self_closing_rows_ignoring_first_non_empty_all_lazy = re.compile(
    rb'(<row\b[^>]+?hidden="1"[^/>]+?/>)(?:<row\b[^>]+?hidden="1"[^/>]+?/>)+'
)
pattern_hidden_self_closing_rows_ignoring_first_non_empty_last_not_lazy = re.compile(
    rb'(<row\b[^>]+?hidden="1"[^/>]+?/>)(?:<row\b[^>]+?hidden="1"[^/>]+/>)+'
)
pattern_hidden_self_closing_rows_ignoring_first_non_empty = re.compile(
    rb'(<row\b[^>]*?hidden="1"[^/>]++/>)(?:<row\b[^>]+?hidden="1"[^/>]++/>)+'
)

pattern_hidden_rows = re.compile(
    rb'(<row\b[^>]*?hidden="1"[^/>]++/>)(?:<row\b[^>]+?hidden="1"[^/>]++/>)+'
)

pattern_hidden_self_closing_rows_ignoring_first_non_empty2 = re.compile(
    rb'('
        rb'<row\b'
        rb'(?=[^>]*\bhidden="1")'
        rb'[^/>]*/>'
    rb')'
    rb'(?:'
        rb'<row\b'
        rb'(?=[^>]*\bhidden="1")'
        rb'[^/>]*/>'
    rb'/>)+'
)      

def remove_consecutive_hidden_self_closing_rows_all_lazy(sheet_data : bytes, params: CleanSheetParameters) -> bytes:
    return apply_pattern_substitution(sheet_data, pattern_hidden_self_closing_rows_ignoring_first_non_empty_all_lazy, rb"\1", "pattern_hidden_rows", params)
def remove_consecutive_hidden_self_closing_rows_last_not_lazy(sheet_data : bytes, params: CleanSheetParameters) -> bytes:
    return apply_pattern_substitution(sheet_data, pattern_hidden_self_closing_rows_ignoring_first_non_empty_last_not_lazy, rb"\1", "pattern_hidden_rows", params)
def remove_consecutive_hidden_self_closing_rows(sheet_data : bytes, params: CleanSheetParameters) -> bytes:
    return apply_pattern_substitution(sheet_data, pattern_hidden_self_closing_rows_ignoring_first_non_empty, rb"\1", "pattern_hidden_rows", params)

pattern_hidden_self_closing_rows_not_ignoring_first_non_empty_tempered_dot = re.compile(
    rb'(<row\b[^/>]*?\bhidden="1"[^/>]*?(?:/>|>(?:(?!</row>).)*?</row>))(?:<row\b[^/>]*\bhidden="1"[^/>]*/>)+'
)
def remove_consecutive_hidden_self_closing_rows_not_ignoring_first_non_empty_tempered_dot(sheet_data : bytes, params: CleanSheetParameters) -> bytes:
    return apply_pattern_substitution(sheet_data, pattern_hidden_self_closing_rows_not_ignoring_first_non_empty_tempered_dot, rb"\1", "pattern_hidden_self_closing_rows_not_ignoring_first_non_empty_tempered_dot", params)

pattern_hidden_rows_lazy = re.compile(
    rb'(<row\b[^/>]*?\bhidden="1"[^/>]+(?:/>|>(?:(?!(?:</row>|<v>)).)+</row>))(?:<row\b[^/>]*?\bhidden="1"[^/>]+(?:/>|>(?:(?!(?:</row>|<v>)).)+</row>))+'
)
def remove_pattern_hidden_rows_lazy(sheet_data : bytes, params: CleanSheetParameters) -> bytes:
    return apply_pattern_substitution(sheet_data, pattern_hidden_rows_lazy, rb"\1", "pattern_hidden_rows_lazy", params)

pattern_hidden_rows_posessive = re.compile(
    rb'(<row\b[^/>]*?\bhidden="1"[^/>]++(?:/>|>(?:(?!(?:</row>|<v>)).)++</row>))(?:<row\b[^/>]*?\bhidden="1"[^/>]++(?:/>|>(?:(?!(?:</row>|<v>)).)++</row>))+'
)
def remove_hidden_rows(sheet_data : bytes, params: CleanSheetParameters) -> bytes:
    return apply_pattern_substitution(sheet_data, pattern_hidden_rows_posessive, rb"\1", "pattern_hidden_rows_posessive", params)
