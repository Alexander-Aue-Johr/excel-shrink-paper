
from typing import Set
from cleaning.common import CellAttributes, CellState, CleanSheetMetrics, RowAttributes, RowState
from common.named_tupples import CleanSheetParameters

def generate_self_closing_row_xml(cleaned_row_attr):
    return b'<row' + cleaned_row_attr + b'/>'

def generate_row_xml(row_attr, row_content):
    return b'<row' + row_attr + b'>' + row_content + b'</row>'

def generate_self_closing_cell_xml(cell_attr):
    return b'<c' + cell_attr + b'/>'

def generate_and_log_deleted_row(metrics : CleanSheetMetrics, row_attributes : RowAttributes) -> bytes:
    metrics.deleted_row_states.append(row_attributes)
    return b''

def generate_and_log_self_closing_row(row_attr: bytes, params : CleanSheetParameters, metrics : CleanSheetMetrics, row_attributes : RowAttributes) -> bytes:
    metrics.kept_row_states.append(row_attributes)

    cleaned_row_attr = params.workbook_specific_patterns.pattern_remove_irrelevant_attributes_for_empty_rows.sub(b'', row_attr)
    
    return generate_self_closing_row_xml(cleaned_row_attr)

def generate_and_log_row(row_attr: bytes, row_content : bytes, metrics : CleanSheetMetrics, row_attributes : RowAttributes) -> bytes:
    metrics.kept_row_states.append(row_attributes)

    return generate_row_xml(row_attr, row_content)

def generate_and_log_self_closing_cell(metrics : CleanSheetMetrics, cell_attributes: CellAttributes) -> bytes:
    metrics.kept_cell_states.append(cell_attributes)

    cell_attribute_string = cell_attributes.cell_attribute_string

    return generate_self_closing_cell_xml(cell_attribute_string)

def generate_and_log_deleted_cell(metrics : CleanSheetMetrics, cell_attributes: CellAttributes) -> bytes:
    metrics.deleted_cell_states.append(cell_attributes)
    return b''



