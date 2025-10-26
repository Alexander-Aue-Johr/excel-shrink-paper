from enum import Enum, auto


class SheetStreamingState(Enum):
    IN_WORKSHEET_NODE = auto()
    IN_COLS_NODE = auto()
    IN_SHEET_DATA_NODE = auto()
    IN_MERGECELLS_NODE = auto()
    FOUND_WORKSHEET_END = auto()
    END_OF_FILE = auto()
