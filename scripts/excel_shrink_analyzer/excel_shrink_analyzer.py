from dataclasses import dataclass, field
import os
import zipfile
from pathlib import Path
from lxml import etree
import re
import csv
from typing import Dict, Tuple, Optional, List
import time


@dataclass
class MergeCell:
    ref: str  # B20:E38

    @staticmethod
    def from_xml(element: etree._Element) -> "MergeCell":
        ref = element.attrib.get("ref")
        return MergeCell(ref=ref)


@dataclass
class MergeCells:
    merge_cells: List[MergeCell]

    @staticmethod
    def from_xml(root: etree._Element) -> "MergeCells":
        merge_cells_list = []
        for merge_cell_el in root.findall("{*}mergeCell"):
            merge_cell = MergeCell.from_xml(merge_cell_el)
            merge_cells_list.append(merge_cell)
        return MergeCells(merge_cells=merge_cells_list)

    def find_max_column(self) -> int:
        max_x = 0
        for merge_cell in self.merge_cells:
            _, end_ref = merge_cell.ref.split(":")
            col_index, _ = coordinate_to_indices(end_ref)
            max_x = max(max_x, col_index)
        return max_x

    def find_max_row(self) -> int:
        max_y = 0
        for merge_cell in self.merge_cells:
            _, end_ref = merge_cell.ref.split(":")
            _, row_index = coordinate_to_indices(end_ref)
            max_y = max(max_y, row_index)
        return max_y


@dataclass
class BorderColor:
    indexed: Optional[str]


@dataclass
class BorderStyle:
    style: Optional[str]
    color: Optional[BorderColor]


@dataclass
class Border:
    left: Optional[BorderStyle]
    right: Optional[BorderStyle]
    top: Optional[BorderStyle]
    bottom: Optional[BorderStyle]
    diagonal: Optional[BorderStyle]

    def has_border(self) -> bool:
        return any(
            [
                self.left is not None and self.left.style is not None,
                self.right is not None and self.right.style is not None,
                self.top is not None and self.top.style is not None,
                self.bottom is not None and self.bottom.style is not None,
                self.diagonal is not None and self.diagonal.style is not None,
            ]
        )


@dataclass
class Borders:
    @staticmethod
    def parse_border_side(
        border_side_el: Optional[etree._Element],
    ) -> Optional[BorderStyle]:
        if border_side_el is None:
            return None

        # Get style attribute if available
        style = border_side_el.attrib.get("style", None)

        # Get color information if available
        color_el = border_side_el.find("{*}color")
        color = (
            BorderColor(indexed=color_el.attrib.get("indexed", None))
            if color_el is not None
            else None
        )

        return BorderStyle(style=style, color=color)


@dataclass
class PatternFill:
    patternType: Optional[str]
    fgColor: Optional[Dict[str, str]]
    bgColor: Optional[Dict[str, str]]


@dataclass
class Fill:
    patternFill: Optional[PatternFill]


@dataclass
class Fills:
    fills: List[Fill]

    @staticmethod
    def from_xml(fills_el: etree._Element) -> "Fills":
        fills_list = []
        for fill_el in fills_el.findall("{*}fill"):
            pattern_fill_el = fill_el.find("{*}patternFill")
            if pattern_fill_el is not None:
                pattern_type = pattern_fill_el.attrib.get("patternType")
                fg_color_el = pattern_fill_el.find("{*}fgColor")
                bg_color_el = pattern_fill_el.find("{*}bgColor")

                fg_color = (
                    {k: v for k, v in fg_color_el.attrib.items()}
                    if fg_color_el is not None
                    else None
                )
                bg_color = (
                    {k: v for k, v in bg_color_el.attrib.items()}
                    if bg_color_el is not None
                    else None
                )

                pattern_fill = PatternFill(
                    patternType=pattern_type, fgColor=fg_color, bgColor=bg_color
                )
            else:
                pattern_fill = None

            fills_list.append(Fill(patternFill=pattern_fill))

        return Fills(fills=fills_list)


@dataclass
class CellXf:
    num_fmt_id: int
    font_id: int
    apply_fill: bool
    fill_id: str
    fill: Fill
    apply_border: bool
    border_id: int
    xf_id: Optional[int]

    @staticmethod
    def from_xml(root: etree._Element, fills: Fills) -> "SharedStrings":
        num_fmt_id = int(root.attrib["numFmtId"])
        font_id = int(root.attrib["fontId"])
        apply_fill = root.attrib.get("applyFill") == "1"
        fill_id = int(root.attrib["fillId"])
        fill = fills.fills[int(fill_id)]
        apply_border = root.attrib.get("applyBorder") == "1"
        border_id = int(root.attrib["borderId"])
        xf_id = root.attrib.get("xfId", None)

        return CellXf(
            num_fmt_id,
            font_id,
            apply_fill,
            fill_id,
            fill,
            apply_border,
            border_id,
            xf_id,
        )

    def has_none_fill_pattern(self) -> bool:
        return not self.fill.patternFill.patternType == "none"


@dataclass
class StyleSheet:
    numFmts: List
    fonts: List
    fills: List
    borders: List[Border]
    cellXfs: List[CellXf]

    @staticmethod
    def from_xml(root: etree._Element) -> "StyleSheet":
        # Borders
        borders_list = []
        borders_el = root.find("{*}borders")
        if borders_el is not None:
            for border_el in borders_el.findall("{*}border"):
                # Parse each border side
                left_el = border_el.find("{*}left")
                right_el = border_el.find("{*}right")
                top_el = border_el.find("{*}top")
                bottom_el = border_el.find("{*}bottom")
                diagonal_el = border_el.find("{*}diagonal")

                # Create BorderStyle objects for each side
                left = Borders.parse_border_side(left_el)
                right = Borders.parse_border_side(right_el)
                top = Borders.parse_border_side(top_el)
                bottom = Borders.parse_border_side(bottom_el)
                diagonal = Borders.parse_border_side(diagonal_el)

                # Append a new Border instance to the list
                borders_list.append(
                    Border(
                        left=left,
                        right=right,
                        top=top,
                        bottom=bottom,
                        diagonal=diagonal,
                    )
                )

        fills_el = root.find("{*}fills")
        fills = Fills.from_xml(fills_el)

        # CellXfs
        cellXfs_list = []
        cellXfs_el = root.find("{*}cellXfs")
        if cellXfs_el is not None:
            for xf_el in cellXfs_el.findall("{*}xf"):
                cellXfs_list.append(CellXf.from_xml(xf_el, fills))

        return StyleSheet(
            numFmts=[], fonts=[], fills=[], borders=borders_list, cellXfs=cellXfs_list
        )


@dataclass
class SharedStrings:
    strings: List[str]

    @staticmethod
    def from_xml(root: etree._Element) -> "SharedStrings":
        strings = []
        for si in root.findall("{*}si"):
            text_elements = si.findall(".//{*}t")
            full_text = "".join(
                [
                    text_element.text
                    for text_element in text_elements
                    if text_element.text
                ]
            )
            strings.append(full_text)

        return SharedStrings(strings=strings)

    def get(self, index: int) -> Optional[str]:
        if 0 <= index < len(self.strings):
            return self.strings[index]
        return None


@dataclass(frozen=True)
class Row:
    element: etree._Element = field(hash=False, compare=False)
    row_index: int
    height: Optional[str] = field(hash=False, compare=False)
    cell_xf: Optional[CellXf] = field(hash=False, compare=False)

    @staticmethod
    def from_xml(row_element: etree._Element, style_sheet: StyleSheet) -> "Row":
        row_index = int(row_element.attrib.get("r", ""))
        height = row_element.attrib.get("ht")
        style_id_str = row_element.attrib.get("s", None)
        style_id = int(style_id_str) if style_id_str is not None else None

        cell_xf = None
        if style_id is not None:
            cell_xf = style_sheet.cellXfs[style_id]

        return Row(row_element, row_index, height, cell_xf)

    def has_no_children(self) -> bool:
        return len(list(self.element)) == 0

    def has_height(self) -> bool:
        return self.height is not None

    def has_no_height(self) -> bool:
        return self.height is None

    def is_in_value_range(self, max_row: int) -> bool:
        return self.row_index <= max_row

    def is_out_of_value_range(self, max_row: int) -> bool:
        return self.row_index > max_row


@dataclass
class Column:
    element: etree._Element
    min: int
    max: int
    width: Optional[float]
    style: Optional[int]
    custom_width: bool
    hidden: bool
    cell_xf: Optional[CellXf]

    @staticmethod
    def from_xml(col_element: etree._Element, style_sheet: StyleSheet) -> "Column":
        min = int(col_element.attrib.get("min"))
        max = int(col_element.attrib.get("max"))
        width = (
            float(col_element.attrib.get("width"))
            if "width" in col_element.attrib
            else None
        )
        style = (
            int(col_element.attrib.get("style"))
            if "style" in col_element.attrib
            else None
        )
        custom_width = col_element.attrib.get("customWidth") == "1"
        hidden = col_element.attrib.get("hidden") == "1"

        cell_xf = style_sheet.cellXfs[style] if style is not None else None

        return Column(
            col_element, min, max, width, style, custom_width, hidden, cell_xf
        )


@dataclass
class ColumnDefinitions:
    columns: List[Column]

    @staticmethod
    def from_xml(
        cols_element: etree._Element, style_sheet: StyleSheet
    ) -> "ColumnDefinitions":
        columns = []

        if cols_element is None:
            return ColumnDefinitions(columns)

        for col_el in cols_element.findall("{*}col"):
            column = Column.from_xml(col_el, style_sheet)
            columns.append(column)
        return ColumnDefinitions(columns=columns)

    def get_cell_xf_by_col_index(self, col_index: int) -> Optional[CellXf]:
        for column in self.columns:
            if column.min <= col_index <= column.max:
                return column.cell_xf
        return None


@dataclass
class Cell:
    element: etree._Element
    reference: str
    value: str
    data_type: str
    style_id: str
    cell_xf: CellXf
    border: Border
    apply_fill: bool
    col_xf: CellXf
    row: Row
    row_xf: CellXf

    def __repr__(self):
        return (
            f"Cell(reference='{self.reference}', value='{self.value}', "
            f"data_type='{self.data_type}', style_id='{self.style_id}', "
            f"border='{self.border}')"
        )

    @staticmethod
    def from_xml(
        cell_element: etree._Element,
        shared_strings: SharedStrings,
        style_sheet: StyleSheet,
        row: Row,
        column_definitions: ColumnDefinitions,
    ) -> "Cell":
        reference = cell_element.attrib.get("r", "")
        data_type = cell_element.attrib.get("t", "")
        style_id_str = cell_element.attrib.get("s", None)
        style_id = int(style_id_str) if style_id_str is not None else 0

        value = ""
        if data_type == "s":
            value_index = int(cell_element.find("{*}v").text)
            value = shared_strings.get(value_index)
        elif data_type == "n":
            value = cell_element.find("{*}v").text
        else:
            v_el = cell_element.find("{*}v")
            if v_el is not None:
                value = v_el.text

        cell_xf = None
        border = None
        apply_fill = None

        cell_xf = style_sheet.cellXfs[style_id]
        borderId = cell_xf.border_id
        border = style_sheet.borders[borderId]

        apply_fill = cell_xf.apply_fill

        col_index, _ = coordinate_to_indices(reference)
        col_xf = column_definitions.get_cell_xf_by_col_index(col_index)
        row_xf = row.cell_xf

        return Cell(
            cell_element,
            reference,
            value,
            data_type,
            style_id,
            cell_xf,
            border,
            apply_fill,
            col_xf,
            row,
            row_xf,
        )

    def has_value(self) -> bool:
        return self.value is not None and self.value != ""

    def has_no_value(self) -> bool:
        return self.value is None or self.value == ""

    def has_only_whitespaces(self) -> bool:
        return self.value is not None and self.value.strip() == ""

    def has_border(self) -> bool:
        if (
            self.row_xf is not None
            and self.row_xf.apply_border
            and self.col_xf is not None
            and self.col_xf.apply_border
        ):
            if (
                self.cell_xf.border_id != self.row_xf.border_id
                or self.cell_xf.border_id != self.col_xf.border_id
            ):
                return True
            else:
                return False

        if self.row_xf is not None and self.row_xf.apply_border:
            if self.cell_xf.border_id != self.row_xf.border_id:
                return True
            else:
                return False

        if self.col_xf is not None and self.col_xf.apply_border:
            if self.cell_xf.border_id != self.col_xf.border_id:
                return True
            else:
                return False

        return self.border is not None and self.border.has_border()

    def has_no_border(self) -> bool:
        return not self.has_border()

    def has_fill(self) -> bool:
        if (
            self.row_xf is not None
            and self.row_xf.apply_fill
            and self.col_xf is not None
            and self.col_xf.apply_fill
        ):
            if (
                self.cell_xf.fill_id != self.row_xf.fill_id
                or self.cell_xf.fill_id != self.col_xf.fill_id
            ):
                return True
            else:
                return False

        if self.row_xf is not None and self.row_xf.apply_fill:
            if self.cell_xf.fill_id != self.row_xf.fill_id:
                return True
            else:
                return False

        if self.col_xf is not None and self.col_xf.apply_fill:
            if self.cell_xf.fill_id != self.col_xf.fill_id:
                return True
            else:
                return False

        has_fill = (
            self.cell_xf.fill is not None
            and self.cell_xf.fill.patternFill is not None
            and self.cell_xf.fill.patternFill.patternType != "none"
        )

        return has_fill

    def has_no_fill(self) -> bool:
        return not self.has_fill()

    def has_fill_pattern_none(self) -> bool:
        if self.cell_xf.fill.patternFill is None:
            return False
        return self.cell_xf.fill.patternFill.patternType == "none"

    def remove(self) -> bool:
        self.row.element.remove(self.element)


def coordinate_to_indices(coordinate: str) -> Tuple[int, int]:
    column_str = "".join(filter(str.isalpha, coordinate))
    row_str = "".join(filter(str.isdigit, coordinate))

    column_index = sum(
        (ord(char.upper()) - ord("A") + 1) * (26**i)
        for i, char in enumerate(reversed(column_str))
    )

    row_index = int(row_str)

    return column_index, row_index


@dataclass
class FindExcessColumnsResult:
    total_cols: int
    columns_to_remove: List[Column]

    def as_list(self) -> List:
        return [self.total_cols, len(self.columns_to_remove)]

    @staticmethod
    def csv_header() -> List:
        return ["Total Column Count", "Removed Columns Count"]


def find_excess_columns(
    column_definitions: ColumnDefinitions, max_column: int
) -> FindExcessColumnsResult:
    total_cols = len(column_definitions.columns)

    columns_to_remove: List[Column] = []

    hidden_columns_exist = any(column.hidden for column in column_definitions.columns)
    if hidden_columns_exist:
        return FindExcessColumnsResult(total_cols, [])

    for column in column_definitions.columns:
        if column.min > max_column:
            columns_to_remove.append(column)

    return FindExcessColumnsResult(total_cols, columns_to_remove)


@dataclass
class FindExcessCellsResult:
    max_row: int
    max_column: int
    total_cells: int
    empty_cells_to_remove: List[Cell]
    whitespace_cells_to_remove: List[Cell]

    def as_list(self) -> List:
        return [
            self.total_cells,
            len(self.empty_cells_to_remove),
            len(self.whitespace_cells_to_remove),
        ]

    @staticmethod
    def csv_header() -> List:
        return ["Total Cell Count", "Empty Cells Count", "Whitespace Cells Count"]


def find_excess_cells(cells: List[Cell]) -> FindExcessCellsResult:
    max_row, max_column = (0, 0)
    empty_cells_to_remove: List[Cell] = []
    whitespace_cells_to_remove: List[Cell] = []

    total_cells = len(cells)

    for cell in cells:
        if cell.has_no_value() and cell.has_no_border() and cell.has_no_fill():
            empty_cells_to_remove.append(cell)
        elif (
            cell.has_only_whitespaces() and cell.has_no_border() and cell.has_no_fill()
        ):
            whitespace_cells_to_remove.append(cell)
        else:
            column_index, row_index = coordinate_to_indices(cell.reference)
            max_column = max(max_column, column_index)
            max_row = max(max_row, row_index)

    return FindExcessCellsResult(
        max_row,
        max_column,
        total_cells,
        empty_cells_to_remove,
        whitespace_cells_to_remove,
    )


@dataclass
class FindExcessRowsResult:
    total_rows: int
    rows_outside_of_value_range: List[Row]
    rows_inside_of_value_range: List[Row]

    def as_list(self) -> List:
        return [self.total_rows, len(self.rows_outside_of_value_range)]

    @staticmethod
    def csv_header() -> List:
        return ["Total Row Count", "Rows Outside of Value Range Count"]


def find_excess_rows(rows: List[Row], max_row: int) -> FindExcessRowsResult:
    total_rows = len(rows)

    rows_outside_of_value_range: List[Row] = []
    rows_inside_of_value_range: List[Row] = []

    for row in rows:
        if row.has_no_children() and row.is_out_of_value_range(max_row=max_row):
            rows_outside_of_value_range.append(row)
        elif (
            row.has_no_children()
            and row.has_height()
            and row.is_in_value_range(max_row=max_row)
        ):
            rows_inside_of_value_range.append(row)

    return FindExcessRowsResult(
        total_rows, rows_outside_of_value_range, rows_inside_of_value_range
    )


@dataclass
class CleanXMLResult:
    cleaned_xml: bytes
    find_excess_cells_result: Optional[FindExcessCellsResult]
    find_excess_rows_result: Optional[FindExcessRowsResult]
    find_excess_columns_result: Optional[FindExcessColumnsResult]
    original_size: int
    cleaned_size: int
    empty_cell_size: int
    whitespace_size: int
    out_of_bounds_rows_size: int
    cols_size: int
    reduction_ratio: int
    empty_cell_reduction_ratio: int
    whitespace_reduction_ratio: int
    out_of_bounds_rows_reduction_ratio: int
    cols_reduction_ratio: int

    def as_list(self) -> List:
        find_excess_cells_result_list = (
            self.find_excess_cells_result.as_list()
            if self.find_excess_cells_result
            else [0, 0, 0]
        )
        find_excess_rows_result_list = (
            self.find_excess_rows_result.as_list()
            if self.find_excess_rows_result
            else [0, 0, 0]
        )
        find_excess_columns_result_list = (
            self.find_excess_columns_result.as_list()
            if self.find_excess_columns_result
            else [0, 0]
        )
        size_list = [
            self.original_size,
            self.cleaned_size,
            self.empty_cell_size,
            self.whitespace_size,
            self.out_of_bounds_rows_size,
            self.cols_size,
        ]
        reduction_list = [
            self.reduction_ratio,
            self.empty_cell_reduction_ratio,
            self.whitespace_reduction_ratio,
            self.out_of_bounds_rows_reduction_ratio,
            self.cols_reduction_ratio,
        ]

        return (
            find_excess_cells_result_list
            + find_excess_rows_result_list
            + find_excess_columns_result_list
            + size_list
            + reduction_list
        )

    @staticmethod
    def csv_header() -> List:
        size_csv_header = [
            "Original Size",
            "Cleaned Size",
            "Empty Cell Size",
            "Whitespace Cell Size",
            "Rows Outside of Value Range Size",
            "Removed Columns Size",
        ]
        reduction_csv_header = [
            "Reduction Ratio",
            "Empty Cell Reduction Ratio",
            "Whitespace Cell Reduction Ratio",
            "Rows Outside of Value Range Reduction Ratio",
            "Removed Columns Reduction Ratio",
        ]

        return (
            FindExcessCellsResult.csv_header()
            + FindExcessRowsResult.csv_header()
            + FindExcessColumnsResult.csv_header()
            + size_csv_header
            + reduction_csv_header
        )


def clean_sheet_xml(
    root: etree._Element,
    sheet_name: str,
    shared_strings: SharedStrings,
    style_sheet: StyleSheet,
    collect_log_data: bool,
) -> CleanXMLResult:
    if collect_log_data is not None:
        original_size = len(
            etree.tostring(root, encoding="utf-8", xml_declaration=True)
        )

    merge_cells = None
    merge_cells_el = root.find("{*}mergeCells")
    if merge_cells_el is not None:
        merge_cells = MergeCells.from_xml(merge_cells_el)

    column_definitions = ColumnDefinitions.from_xml(root.find("{*}cols"), style_sheet)

    cols = root.find("{*}cols")
    sheet_data = root.find("{*}sheetData")

    # parse all rows and cells
    rows: List[Row] = []
    cells: List[Cell] = []
    if sheet_data is not None:
        for row_element in sheet_data.findall("{*}row"):
            row = Row.from_xml(row_element, style_sheet)
            rows.append(row)
            for cell_el in row_element.findall("{*}c"):
                cell = Cell.from_xml(
                    cell_el,
                    shared_strings=shared_strings,
                    style_sheet=style_sheet,
                    row=row,
                    column_definitions=column_definitions,
                )

                cells.append(cell)

    find_excess_cells_result = find_excess_cells(cells)

    if find_excess_cells_result is not None:
        for cell in find_excess_cells_result.empty_cells_to_remove:
            cell.remove()
        for row in {
            cell.row for cell in find_excess_cells_result.empty_cells_to_remove
        }:
            if row.has_height():
                row.element.set("customHeight", "1")

    size_after_empty_cell_removal, empty_cell_size = calculate_size_and_difference(
        root, original_size, collect_log_data
    )

    if find_excess_cells_result is not None:
        for cell in find_excess_cells_result.whitespace_cells_to_remove:
            cell.remove()
        for row in set(
            [cell.row for cell in find_excess_cells_result.whitespace_cells_to_remove]
        ):
            if row.has_height():
                row.element.set("customHeight", "1")

    size_after_whitespace_removal, whitespace_size = calculate_size_and_difference(
        root, size_after_empty_cell_removal, collect_log_data
    )

    max_row = (
        max(merge_cells.find_max_row(), find_excess_cells_result.max_row)
        if merge_cells is not None
        else find_excess_cells_result.max_row
    )

    find_excess_rows_result = find_excess_rows(rows, max_row)

    if find_excess_rows_result is not None:
        for row in find_excess_rows_result.rows_outside_of_value_range:
            sheet_data.remove(row.element)

    size_after_out_of_bounds_row_removal, out_of_bounds_rows_size = (
        calculate_size_and_difference(
            root, size_after_whitespace_removal, collect_log_data
        )
    )

    max_column = (
        max(merge_cells.find_max_column(), find_excess_cells_result.max_column)
        if merge_cells is not None
        else find_excess_cells_result.max_column
    )

    find_excess_columns_result = find_excess_columns(column_definitions, max_column)
    if (
        find_excess_columns_result is not None
        and len(find_excess_columns_result.columns_to_remove) > 0
    ):
        columns_to_remove = find_excess_columns_result.columns_to_remove

        first_col = columns_to_remove[0]

        last_col = columns_to_remove[-1]
        last_col.element.set("min", str(first_col.min))

        all_but_last_col = columns_to_remove[:-1]
        for col in all_but_last_col:
            cols.remove(col.element)

    cleaned_xml = etree.tostring(root, encoding="utf-8", xml_declaration=True)

    if collect_log_data is not None:
        cleaned_size = len(cleaned_xml)
        cols_size = size_after_out_of_bounds_row_removal - cleaned_size
    else:
        cleaned_size = 0
        cols_size = 0

    return CleanXMLResult(
        cleaned_xml,
        find_excess_cells_result,
        find_excess_rows_result,
        find_excess_columns_result,
        original_size,
        cleaned_size,
        empty_cell_size,
        whitespace_size,
        out_of_bounds_rows_size,
        cols_size,
        reduction_ratio=(original_size - cleaned_size) / original_size,
        empty_cell_reduction_ratio=empty_cell_size / original_size,
        whitespace_reduction_ratio=whitespace_size / original_size,
        out_of_bounds_rows_reduction_ratio=out_of_bounds_rows_size / original_size,
        cols_reduction_ratio=cols_size / original_size,
    )


def calculate_size_and_difference(
    root: etree._Element, size_before: int, collect_log_data: bool
) -> Tuple[int, int]:
    if not collect_log_data:
        return 0, 0

    shrinked_size = len(etree.tostring(root, encoding="utf-8", xml_declaration=True))
    empty_cell_size = size_before - shrinked_size
    return shrinked_size, empty_cell_size


def clean_workbook_xml(root: etree._Element) -> Tuple[bytes, int, int]:
    definedNames = root.find("{*}definedNames")
    removed_print_areas = 0
    removed_print_titles = 0

    if definedNames is not None:
        # Create dictionaries to track "_xlnm.Print_Area" and "Print_Area" nodes for each localSheetId
        xlnm_print_area_tracker = {}
        print_area_tracker = {}
        xlnm_print_titles_tracker = {}
        print_titles_tracker = {}

        for definedName in definedNames.findall("{*}definedName"):
            localSheetId = definedName.attrib.get("localSheetId")
            name = definedName.attrib.get("name")

            if localSheetId is not None:
                if name == "_xlnm.Print_Area":
                    xlnm_print_area_tracker[localSheetId] = definedName
                elif name == "Print_Area":
                    print_area_tracker[localSheetId] = definedName
                elif name == "_xlnm.Print_Titles":
                    xlnm_print_titles_tracker[localSheetId] = definedName
                elif name == "Print_Titles":
                    print_titles_tracker[localSheetId] = definedName

        # Remove "Print_Area" nodes if a corresponding "_xlnm.Print_Area" node is found
        for localSheetId, print_area_node in print_area_tracker.items():
            if localSheetId in xlnm_print_area_tracker:
                definedNames.remove(print_area_node)
                removed_print_areas += 1

        # Remove "Print_Titles" nodes if a corresponding "_xlnm.Print_Titles" node is found
        for localSheetId, print_titles_node in print_titles_tracker.items():
            if localSheetId in xlnm_print_titles_tracker:
                definedNames.remove(print_titles_node)
                removed_print_titles += 1

    return (
        etree.tostring(root, encoding="utf-8", xml_declaration=True),
        removed_print_areas,
        removed_print_titles,
    )


def styles_xml(style_content: etree._Element) -> Dict[int, str]:
    root = style_content
    # Dictionary to store mapping of sheet names to sheet XML files
    sheet_map = {}
    # Find all the sheet elements under the <sheets> tag
    sheets = root.findall("{*}sheets/{*}sheet")

    for sheet in sheets:
        # Get the sheet name and the r:id which links to sheetX.xml
        sheet_name = sheet.attrib["name"]
        sheet_id_string = sheet.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]  # Handle namespace for r:id
        sheet_id = int(
            sheet_id_string[3:]
        )  # This extracts 1 from 'rId1', 2 from 'rId2', etc.

        # Add the mapping to the dictionary
        sheet_map[sheet_id] = sheet_name

    return sheet_map


def map_sheet_names_to_ids(workbook_content: etree._Element) -> Dict[int, str]:
    root = workbook_content
    # Dictionary to store mapping of sheet names to sheet XML files
    sheet_map = {}
    # Find all the sheet elements under the <sheets> tag
    sheets = root.findall("{*}sheets/{*}sheet")

    for sheet in sheets:
        # Get the sheet name and the r:id which links to sheetX.xml
        sheet_name = sheet.attrib["name"]
        sheet_id_string = sheet.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]  # Handle namespace for r:id
        sheet_id = int(
            sheet_id_string[3:]
        )  # This extracts 1 from 'rId1', 2 from 'rId2', etc.

        # Add the mapping to the dictionary
        sheet_map[sheet_id] = sheet_name

    return sheet_map


def extract_sheet_number(filename: str) -> Optional[int]:
    # Use a regular expression to find the number in the filename
    match = re.search(r"sheet(\d+)\.xml", filename)

    if match:
        return int(match.group(1))  # Extract the matched number as an integer
    return None


def process_xlsx_in_place(
    file_path: str, output_path: str, collect_log_data: bool
) -> list:
    with zipfile.ZipFile(file_path, "r") as zin, zipfile.ZipFile(
        output_path, "w"
    ) as zout:
        sheet_mapping = {}

        with zin.open("xl/workbook.xml") as workbook:
            workbook_content = workbook.read()  # Decode bytes to string
            workbook_xml = etree.XML(workbook_content)

            sheet_mapping = map_sheet_names_to_ids(workbook_xml)

        shared_strings = None
        with zin.open("xl/sharedStrings.xml") as shared_strings_file:
            shared_strings_content = (
                shared_strings_file.read()
            )  # Decode bytes to string
            shared_strings_el = etree.XML(shared_strings_content)
            shared_strings = SharedStrings.from_xml(shared_strings_el)

        styles = None
        with zin.open("xl/styles.xml") as styles_file:
            styles_content = styles_file.read()  # Decode bytes to string
            styles_el = etree.XML(styles_content)
            styles = StyleSheet.from_xml(styles_el)

        # Create a list to store the log data
        log_data = []

        # Iterate through all files in the archive
        sheets = [
            item
            for item in zin.infolist()
            if item.filename.startswith("xl/worksheets/sheet")
            and item.filename.endswith(".xml")
        ]
        other_files = [item for item in zin.infolist() if item not in sheets]

        for index, item in enumerate(sheets):
            with zin.open(item.filename) as sheet:
                xml_content = sheet.read()
                xml = etree.XML(xml_content)

                sheet_number = extract_sheet_number(item.filename)
                if sheet_number is not None:

                    sheet_name = sheet_mapping[sheet_number]
                    print(
                        f"Processing sheet {index + 1} of {len(sheets)}: {sheet_name}",
                        flush=True,
                    )
                    clean_xml_result = clean_sheet_xml(
                        xml, sheet_name, shared_strings, styles, collect_log_data
                    )

                    if collect_log_data:
                        log_data.append(
                            [file_path, sheet_name]
                            + clean_xml_result.as_list()
                            + [0] * 2
                        )

                    # Write the cleaned sheet back to the zip
                    zout.writestr(item, clean_xml_result.cleaned_xml)

        for index, item in enumerate(other_files):
            if item.filename == "xl/workbook.xml":
                with zin.open(item.filename) as workbook:
                    workbook_content = workbook.read()
                    workbook_xml = etree.XML(workbook_content)

                    cleaned_workbook, removed_print_areas, removed_print_titles = (
                        clean_workbook_xml(workbook_xml)
                    )

                    if collect_log_data:
                        log_data.append(
                            [file_path, "Workbook"]
                            + [0] * 18
                            + [removed_print_areas, removed_print_titles]
                        )

                    zout.writestr(item, cleaned_workbook)
            else:
                # For all other files, copy them as is
                zout.writestr(item, zin.read(item.filename))

    return log_data


def process_directory(
    directory_path: str,
    output_dir: str,
    collect_log_data: bool,
    recursive: bool = False,
) -> list:
    xlsx_files = []
    log_data = []

    if recursive:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                if file.endswith(".xlsx"):
                    xlsx_files.append(os.path.join(root, file))
    else:
        xlsx_files = [
            os.path.join(directory_path, file)
            for file in os.listdir(directory_path)
            if file.endswith(".xlsx")
        ]

    for index, file_path in enumerate(xlsx_files):
        print(
            f"Processing file {index + 1} of {len(xlsx_files)}: {file_path}", flush=True
        )

        output_file = Path(output_dir) / Path(file_path).name
        try:
            log_data_for_file = process_xlsx_in_place(
                file_path, output_file, collect_log_data
            )

            for log_data_to_add in log_data_for_file:
                log_data.append(log_data_to_add)
        except zipfile.BadZipFile:
            print(f"Error: {file_path} is not a valid xlsx file.", flush=True)
        except PermissionError:
            print(f"Error: Permission denied for file {file_path}.", flush=True)

    return log_data


def main(
    path: str, output_path: str, recursive: bool = False, log_file: Optional[str] = None
) -> None:
    start_time = time.time()

    # Check if the output directory exists, and create it if necessary
    if not os.path.exists(output_path):
        if path.endswith(".xlsx"):
            parent_dir = os.path.dirname(output_path)
            if parent_dir and not os.path.exists(parent_dir):
                os.makedirs(parent_dir)
        else:
            os.makedirs(output_path)

    log_data = []

    if os.path.isfile(path) and path.endswith(".xlsx"):
        # If it's a file, process it
        log_data = process_xlsx_in_place(path, output_path, log_file is not None)
    elif os.path.isdir(path):
        # If it's a directory, process all xlsx files in it
        log_data = process_directory(path, output_path, log_file is not None, recursive)
    else:
        print("Invalid path provided.", flush=True)

    # Write the log data to the CSV file if log_file is provided
    if log_file:
        log_file_path = Path(output_path) / log_file
        with open(log_file_path, "a", newline="") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(
                ["File", "Sheet / Workbook"]
                + CleanXMLResult.csv_header()
                + ["Removed Print Areas", "Removed Print Titles"]
            )
            writer.writerows(log_data)

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds.", flush=True)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Process and clean xlsx files.")
    parser.add_argument(
        "path", help="Path to xlsx file or directory containing xlsx files."
    )
    parser.add_argument("output", help="Path to output directory or file.")
    parser.add_argument(
        "--recursive", action="store_true", help="Recursively process directories."
    )
    parser.add_argument("--log", help="Path to save the log CSV file.", default=None)

    args = parser.parse_args()
    main(args.path, args.output, args.recursive, args.log)
