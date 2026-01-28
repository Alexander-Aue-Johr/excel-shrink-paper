import re
from typing import Set

from common.apply_pattern_substitution import apply_pattern_substitution
from common.extract_xml_block import extract_xml_block
from common.named_tupples import CleanWorkbookParameters, DefinedName


def replace_xml_block(
    xml_bytes: bytes, new_block: bytes | None, block_start: int, block_end: int
) -> bytes:
    """
    Replace the region [block_start:block_end] in xml_bytes
    with new_block, returning the updated bytes.
    """

    if new_block is None:
        return xml_bytes

    return xml_bytes[:block_start] + new_block + xml_bytes[block_end:]


search_duplicate_defined_name_pattern = re.compile(
    rb'<definedName\b[^>]*?name="([^"]*?)"'
    rb'(?:[^>]*?localSheetId="([^"]*?)")?'
    rb"[^>]*?>",
    re.DOTALL,
)

replace_duplicate_defined_name_pattern = re.compile(
    rb'<definedName\b[^>]*?name="([^"]*?)"'
    rb'(?:[^>]*?localSheetId="([^"]*?)")?'
    rb"[^>]*?>.*?</definedName>",
    re.DOTALL,
)


def clean_workbook(workbook_xml_bytes: bytes, params: CleanWorkbookParameters) -> bytes:
    # Four dictionaries to store nodes, keyed by the integer localSheetId.
    print_area_nodes = {}
    print_title_nodes = {}
    filter_database_nodes = {}
    database_nodes = {}
    xlnm_print_area_nodes = {}
    xlnm_print_title_nodes = {}
    xlnm_filter_database_nodes = {}
    xlnm_database_nodes = {}

    defined_names_content, next_index, start, end = extract_xml_block(
        workbook_xml_bytes, b"<definedNames", b"</definedNames>"
    )
    if defined_names_content is None:
        return workbook_xml_bytes

    matches = search_duplicate_defined_name_pattern.findall(
        defined_names_content
    )  # This line is not necessary, as we already iterate over all matches above.
    for m in matches:
        name, local_sheet_id = m

        if local_sheet_id is None:
            local_sheet_id = b""

        if name == b"Print_Area":
            print_area_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)
        elif name == b"_xlnm.Print_Area":
            xlnm_print_area_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)
        elif name == b"Print_Titles":
            print_title_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)
        elif name == b"_xlnm.Print_Titles":
            xlnm_print_title_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)
        elif name == b"_FilterDatabase":
            filter_database_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)
        elif name == b"_xlnm._FilterDatabase":
            xlnm_filter_database_nodes[local_sheet_id] = DefinedName(
                name, local_sheet_id
            )
        elif name == b"Database":
            database_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)
        elif name == b"_xlnm.Database":
            xlnm_database_nodes[local_sheet_id] = DefinedName(name, local_sheet_id)

    defined_names_to_delete: Set[DefinedName] = set()

    for local_sheet_id, print_area_defined_name in print_area_nodes.items():
        if local_sheet_id in xlnm_print_area_nodes:
            defined_names_to_delete.add(print_area_defined_name)

    for local_sheet_id, print_title_defined_name in print_title_nodes.items():
        if local_sheet_id in xlnm_print_title_nodes:
            defined_names_to_delete.add(print_title_defined_name)

    for local_sheet_id, filter_database_defined_name in filter_database_nodes.items():
        if local_sheet_id in xlnm_filter_database_nodes:
            defined_names_to_delete.add(filter_database_defined_name)

    for local_sheet_id, database_defined_name in database_nodes.items():
        if local_sheet_id in xlnm_database_nodes:
            defined_names_to_delete.add(database_defined_name)

    if not defined_names_to_delete:
        return workbook_xml_bytes

    def remove_duplicate_defined_names(match: re.Match) -> bytes:
        name = match.group(1)
        local_sheet_id = match.group(2) or b""
        defined_name = DefinedName(name, local_sheet_id)

        if defined_name in defined_names_to_delete:
            return b""
        return match.group(0)

    defined_names_content = apply_pattern_substitution(
        defined_names_content,
        replace_duplicate_defined_name_pattern,
        remove_duplicate_defined_names,
        "Remove duplicate defined names",
        params,
    )

    workbook_xml_bytes = replace_xml_block(
        workbook_xml_bytes, defined_names_content, start, end
    )

    return workbook_xml_bytes
