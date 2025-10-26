#!/usr/bin/env python3
import os
import zipfile
import re
from typing import IO, List, NamedTuple
import time
import logging

from cleaning.clean_empty_rows_from_after_cell_range import clean_empty_rows_from_after_cell_range
from cleaning.columns.merge_invisible_columns import merge_extraneous_columns
from cleaning.common import CellAttributes, CleanSheetMetrics, RowAttributes
from cleaning.generate_workbook_specific_patterns import generate_workbook_specific_patterns
from common.letters_byte_string_to_integer import letters_byte_string_to_integer
from parsing.parse_xfs import parse_xfs
from cleaning.clean_sheet_rows_and_cells_in_cell_range import clean_sheet_rows_and_cells_in_cell_range
from common.extract_xml_block import extract_xml_block
from common.named_tupples import XF, CleanSheetParameters, CleanWorkbookParameters, CleanXlsxParameters

from optimized_regex_excel_shrink.find_whitespace_sst_indices import find_whitespace_sst_indices
from optimized_regex_excel_shrink.clean_workbook import clean_workbook
from streaming.clean_sheet_xml_streamwise import preprocess_sheet_xml_streamwise


def clean_xml_data(params: CleanWorkbookParameters | CleanSheetParameters, chunk_size : int) -> bytes:
    logging.info(f"Processing {params.zip_info.filename} of {params.xlsx_file_path}")

    if logging.getLogger().isEnabledFor(logging.INFO):
        start_time = time.time()

    with zipfile.ZipFile(params.xlsx_file_path, 'r') as z:

        if isinstance(params, CleanSheetParameters):

            with z.open(params.zip_info.filename) as xml_data_stream:
                preprocessed_sheet_xml = preprocess_sheet_xml_streamwise(
                    xml_data_stream, params, chunk_size
                )
                metrics = preprocessed_sheet_xml.clean_sheet_metrics

                ranges = retrieve_sheet_data_range_containing_cells(preprocessed_sheet_xml.sheet_data, params)

                cleaned_cells_range = clean_sheet_rows_and_cells_in_cell_range(ranges.cells_range, params, metrics)
                cleaned_visible_rows_range = clean_empty_rows_from_after_cell_range(ranges.rows_range, params, metrics)

                final_sheet_data = cleaned_cells_range + cleaned_visible_rows_range

                max_column = get_max_column(metrics)
                in_cols_node_output_chunks = merge_extraneous_columns(preprocessed_sheet_xml.in_cols_node_output_chunks, max_column)

                final_data = (
                    preprocessed_sheet_xml.before_cols_node_output_chunks +
                    preprocessed_sheet_xml.in_cols_node_output_chunks +
                    preprocessed_sheet_xml.before_sheet_data_node_output_chunks +
                    final_sheet_data +
                    preprocessed_sheet_xml.before_merge_cells_node_output_chunks +
                    preprocessed_sheet_xml.in_merge_cells_node_output_chunks +
                    preprocessed_sheet_xml.until_worksheet_end_output_chunks
                )
                
                return final_data

                
            #with z.open(clean_parameters.zip_info.filename) as xml_data_stream:
            #    clean_xml_result = clean_sheet_xml(
            #        xml_data_stream.read(), clean_parameters
            #    )

        elif isinstance(params, CleanWorkbookParameters):
            with z.open(params.zip_info.filename) as xml_data_stream:
                xml_data = xml_data_stream.read()

                clean_xml_result = clean_workbook(
                    xml_data, params
                )

        if logging.getLogger().isEnabledFor(logging.INFO):
            clean_size = len(clean_xml_result)
            original_size = params.zip_info.file_size
            reduction_bytes = original_size - clean_size
            reduction_ratio = reduction_bytes / original_size if original_size else 0

            end_time = time.time()
            elapsed_time = end_time - start_time

            logging.info(f"Processed Workbook of {params.xlsx_file_path}: Original size {params.zip_info.file_size} bytes, Clean size {clean_size} bytes, "
                f"reduction {reduction_bytes} bytes ({reduction_ratio:.2%} reduction) completed in {elapsed_time:.2f} seconds.")

    return clean_xml_result


def get_max_merge_cells_column_index(metrics: CleanSheetMetrics) -> int | None:
    max_merge_cells_column_and_row = metrics.max_merge_cells_column_and_row

    if max_merge_cells_column_and_row:
        return letters_byte_string_to_integer(max_merge_cells_column_and_row.column)
    
    return None

def get_max_visible_column_index(metrics: CleanSheetMetrics) -> int | None:
    last_visible_column_range = metrics.column_definitions.get_last_visible_range()

    if last_visible_column_range:
        return last_visible_column_range.max
    
    return None

def get_max_column(metrics: CleanSheetMetrics) -> int | None:
    max_cell_column = get_max_cell_column(metrics)
    max_merge_cells_column = get_max_merge_cells_column_index(metrics)
    max_visible_column_index = get_max_visible_column_index(metrics)

    candidates = [
        column_index for column_index in (
            max_cell_column,
            max_merge_cells_column,
            max_visible_column_index
        ) if column_index is not None
    ]
    
    return max(candidates) if candidates else None




def get_max_cell_column(metrics: CleanSheetMetrics)-> int | None:
    kept_cells = metrics.kept_cell_states
    if not kept_cells:
        return None

    all_distinct_columns = {cell.cell_coordinate.column for cell in kept_cells}
    max_column = max(all_distinct_columns, key=lambda column: len(column))
    
    return letters_byte_string_to_integer(max_column)



class VisibleSheetDataRanges(NamedTuple):
    cells_range: bytes
    rows_range: bytes


def find_last_or_border_style_or_hidden_attribute(sheet_data : bytes, params : CleanSheetParameters, last_row_with_cell_end_pos : int) -> int:
    fill_or_border_or_hidden_attribute_pattern = params.workbook_specific_patterns.fill_or_border_or_hidden_attribute_pattern
    last_pos = last_row_with_cell_end_pos

    last_visible_attribute_match = None
    for match in fill_or_border_or_hidden_attribute_pattern.finditer(sheet_data, last_row_with_cell_end_pos):
        last_visible_attribute_match = match

    if last_visible_attribute_match:
        last_pos = max(last_pos, last_visible_attribute_match.start())

    return last_pos


def find_last_visible_row_end_position(sheet_data : bytes, last_visible_attribute_position : int) -> int:
    row_end_tag_pos = sheet_data.find(b"</row>", last_visible_attribute_position)
    self_closing_tag_pos = sheet_data.find(b"/>", last_visible_attribute_position)

    if row_end_tag_pos != -1 and self_closing_tag_pos != -1:
        if row_end_tag_pos < self_closing_tag_pos:
            return row_end_tag_pos + len(b"</row>")
        else:
            return self_closing_tag_pos + len(b"/>")

    if row_end_tag_pos == -1:	
        return self_closing_tag_pos + len(b"/>")
    if self_closing_tag_pos == -1:
        return row_end_tag_pos + len(b"</row>")

    raise ValueError("No </row> or /> tag found after last visible style position")


def find_last_row_with_cell_end_pos(sheet_data : bytes) -> int:
    last_cell_pos = sheet_data.rfind(b"<c ")
    if last_cell_pos == -1:
        return -1
    last_row_with_cell_end_pos = sheet_data.find(b"</row>", last_cell_pos) + len(b"</row>")
    if last_row_with_cell_end_pos == -1:
        assert False, "No </row> tag found after last cell position"
    
    return last_row_with_cell_end_pos

def retrieve_sheet_data_range_containing_cells(sheet_data : bytes, params : CleanSheetParameters) -> VisibleSheetDataRanges:
    last_row_with_cell_end_pos = find_last_row_with_cell_end_pos(sheet_data)

    last_or_border_style_or_hidden_attribute = find_last_or_border_style_or_hidden_attribute(sheet_data, params, last_row_with_cell_end_pos)
    last_visible_row_end_position = find_last_visible_row_end_position(sheet_data, last_or_border_style_or_hidden_attribute)

    if last_row_with_cell_end_pos == -1:
        cells_range = b''
    else:
        cells_range = sheet_data[:last_row_with_cell_end_pos]

    if last_visible_row_end_position == -1:
        rows_range = b''
    else:
        rows_range_start = max(0, last_row_with_cell_end_pos)
        rows_range = sheet_data[rows_range_start:last_visible_row_end_position]
    
    return VisibleSheetDataRanges(cells_range, rows_range)



def find_apply_fill_border_ids(all_xfs: List[XF]) -> set[int]:
    """
    Identify which style indices have applyFill or applyBorder,
    and which have no style but do have a font.
    """
    return {
        i for i, xf in enumerate(all_xfs)
        if xf.apply_fill or xf.apply_border
    }



si_pattern = re.compile(rb"<si>.*?<\/si>", flags=re.DOTALL)

def generate_clean_sheet_params_for_xlsx_file(file_path: str | IO[bytes], sheet_names: List[str] | None = None, only_clean_workbook : bool = False) -> CleanXlsxParameters:
    clean_sheet_parameters_list = []

    with zipfile.ZipFile(file_path, 'r') as zin:
        infolist = zin.infolist()

        workbook_zip_info = None
        sheets : List[zipfile.ZipInfo] = []
        xml_files_excluded_from_cleaning : List[zipfile.ZipInfo] = []

        for item in infolist:
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                if only_clean_workbook:
                    xml_files_excluded_from_cleaning.append(item)
                    continue
                if sheet_names and os.path.basename(item.filename) not in sheet_names:
                    xml_files_excluded_from_cleaning.append(item)
                    continue

                sheets.append(item)
            elif item.filename == "xl/workbook.xml":
                workbook_zip_info = item
                clean_workbook_parameters = CleanWorkbookParameters(file_path, item)
            else:
                xml_files_excluded_from_cleaning.append(item)

        with zin.open("xl/sharedStrings.xml") as shared_strings_file:
            shared_strings_xml = shared_strings_file.read()
            si_matches = si_pattern.findall(shared_strings_xml)
        
        whitespace_indices = find_whitespace_sst_indices(si_matches)

        with zin.open("xl/styles.xml") as styles_file:
            styles_xml = styles_file.read()
            cell_styles, _, _, _ = extract_xml_block(styles_xml, b"<cellXfs", b"</cellXfs>")
            if cell_styles is None:
                raise ValueError("Cell styles XML block not found in styles.xml")
            all_xfs = parse_xfs(cell_styles)


        workbook_specific_patterns = generate_workbook_specific_patterns(all_xfs, whitespace_indices)

        for sheet_zip_info in sheets:          
            clean_sheet_parameters = CleanSheetParameters(
                xlsx_file_path=file_path,
                zip_info=sheet_zip_info,
                workbook_specific_patterns=workbook_specific_patterns)
            clean_sheet_parameters_list.append(clean_sheet_parameters)

    if workbook_zip_info is None:
        raise ValueError("Workbook XML file not found in the zip archive.")

    return CleanXlsxParameters(file_path, clean_workbook_parameters, clean_sheet_parameters_list, workbook_zip_info, sheets, xml_files_excluded_from_cleaning)