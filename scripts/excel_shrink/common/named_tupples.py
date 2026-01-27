from bisect import bisect_right
from dataclasses import dataclass
from enum import Enum, auto
from typing import NamedTuple, Pattern, Set, Dict, List, Optional, Any
from zipfile import ZipFile, ZipInfo

from cleaning.generate_workbook_specific_patterns import WorkbookSpecificPatterns
from common.sheet_streaming_state import SheetStreamingState
from common.slice_bytes import WhereToSlice
from parsing.parse_xfs import XF


class ColumnDefinition(NamedTuple):
    min: int
    max: int
    cell_xf: Optional[XF]
    hidden: bool


class ColumnRangeMap:
    exact: Dict[int, ColumnDefinition]
    intervals: List[ColumnDefinition]
    _interval_starts: List[int]

    def __init__(self):
        self.exact = {}
        self.intervals = []
        self._interval_starts = []

    def add(self, col_def: ColumnDefinition):
        if col_def.min == col_def.max:
            self.exact[col_def.min] = col_def
        else:
            self.intervals.append(col_def)

    def finalize(self):
        self.intervals.sort(key=lambda cd: cd.min)
        self._interval_starts = [cd.min for cd in self.intervals]

    def __getitem__(self, col: int) -> ColumnDefinition:
        if col in self.exact:
            return self.exact[col]
        idx = bisect_right(self._interval_starts, col) - 1
        if idx >= 0:
            cd = self.intervals[idx]
            if cd.min <= col <= cd.max:
                return cd
        raise KeyError(col)

    def get(
        self, col: int, default: Optional[ColumnDefinition] = None
    ) -> Optional[ColumnDefinition]:
        if col in self.exact:
            return self.exact[col]
        idx = bisect_right(self._interval_starts, col) - 1
        if idx >= 0:
            cd = self.intervals[idx]
            if cd.min <= col <= cd.max:
                return cd
        return None

    def __contains__(self, col: int) -> bool:
        if col in self.exact:
            return True
        idx = bisect_right(self._interval_starts, col) - 1
        if idx >= 0:
            cd = self.intervals[idx]
            return cd.min <= col <= cd.max
        return False

    def __len__(self):
        return len(self.exact) + len(self.intervals)

    def get_last_visible_range(self) -> Optional[ColumnDefinition]:

        last_visible: ColumnDefinition | None = None

        for cd in self.exact.values():
            if (
                cd.hidden
                or cd.cell_xf
                and (cd.cell_xf.has_fill_or_border or cd.cell_xf.has_apply_protection)
            ):
                if last_visible is None or cd.max > last_visible.max:
                    last_visible = cd

        for cd in self.intervals:
            if (
                cd.hidden
                or cd.cell_xf
                and (cd.cell_xf.has_fill_or_border or cd.cell_xf.has_apply_protection)
            ):
                if last_visible is None or cd.max > last_visible.max:
                    last_visible = cd

        return last_visible


@dataclass(frozen=True)
class TagToProcess:
    tag_start: bytes
    tag_end: bytes | None = None
    where_to_slice: WhereToSlice = WhereToSlice.before_tag


class CollectChunksResult(NamedTuple):
    new_state: SheetStreamingState | None
    chunk_to_add: bytes | None = None
    updated_buffer: bytes | None = None
    processed_tag: TagToProcess | None = None


class CollectColsResult(NamedTuple):
    new_state: SheetStreamingState | None
    chunk_to_add: bytes
    updated_buffer: bytes
    column_definitions: ColumnRangeMap


class CleanSheetDataResult(NamedTuple):
    chunk_to_add: bytes
    updated_buffer: bytes
    new_state: SheetStreamingState | None = None


class CellCoordinate(NamedTuple):
    column: bytes
    row: bytes


class MergeCellsParsingResult(NamedTuple):
    new_state: SheetStreamingState | None
    chunk_to_add: bytes
    updated_buffer: bytes
    max_col_and_row: CellCoordinate | None


class DefinedName(NamedTuple):
    name: bytes
    local_sheet_id: bytes


class CleanSheetParameters(NamedTuple):
    xlsx_file_path: str
    zip_info: ZipInfo
    workbook_specific_patterns: WorkbookSpecificPatterns


class XlsxFileToWrite(NamedTuple):
    zip_handle: ZipFile
    remaining_files_to_write: List[str]
    successfully_written_sheet_files: List[str]
    failed_sheet_files: List[CleanSheetParameters]


class CleanWorkbookParameters(NamedTuple):
    xlsx_file_path: str
    zip_info: ZipInfo


class CleanXlsxParameters(NamedTuple):
    file_path: str
    clean_workbook_parameters: CleanWorkbookParameters
    clean_sheet_parameters: List[CleanSheetParameters]
    workbook_zip_info: ZipInfo
    sheet_zip_files: List[ZipInfo]
    other_zip_files: List[ZipInfo]
    only_clean_workbook: bool
