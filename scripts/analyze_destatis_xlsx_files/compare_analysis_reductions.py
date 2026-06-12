from __future__ import annotations

import argparse
import shutil
import subprocess
import os
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd


DEFAULT_PYTHON_ANALYSIS_CSV = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "analysis_output.csv",
)
DEFAULT_LONG_ANALYSIS_CSV = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "excel_shrink_rust_analysis.csv",
)
DEFAULT_OUTPUT_CSV = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "analysis_reduction_comparison.csv",
)
DEFAULT_DIFF_OUTPUT = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "analysis_reduction_top_diffs.diff",
)
DEFAULT_PYTHON_XLSX_DIR = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "shrinked_by_analyzer",
)
DEFAULT_LONG_XLSX_DIR = Path(
    "scripts",
    "analyze_destatis_xlsx_files",
    "shrinked_by_excel_shrink_rust",
)

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def normalize_sheet_name(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value)


def normalize_file_name(value: object) -> str:
    if pd.isna(value):
        return ""
    return Path(str(value)).name


def normalize_zip_path(path: str) -> str:
    return posixpath.normpath(path.replace("\\", "/")).lstrip("/")


def workbook_target_to_zip_path(target: str) -> str:
    if target.startswith("/"):
        return normalize_zip_path(target)
    return normalize_zip_path(posixpath.join("xl", target))


def build_sheet_path_to_name_map(xlsx_path: Path) -> dict[str, str]:
    try:
        with zipfile.ZipFile(xlsx_path) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
            workbook_rels_xml = archive.read("xl/_rels/workbook.xml.rels")
    except (KeyError, OSError, zipfile.BadZipFile):
        return {}

    rel_root = ET.fromstring(workbook_rels_xml)
    rel_targets = {
        rel.attrib.get("Id"): workbook_target_to_zip_path(rel.attrib.get("Target", ""))
        for rel in rel_root.findall(f"{{{REL_NS}}}Relationship")
    }

    workbook_root = ET.fromstring(workbook_xml)
    result: dict[str, str] = {}

    for sheet in workbook_root.findall(f".//{{{WORKBOOK_NS}}}sheet"):
        sheet_name = sheet.attrib.get("name")
        relationship_id = sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id")

        if not sheet_name or not relationship_id:
            continue

        sheet_path = rel_targets.get(relationship_id)
        if sheet_path:
            result[sheet_path] = sheet_name

    return result


def read_python_analysis(path: Path) -> pd.DataFrame:
    data = pd.read_csv(path)

    for column in ["Original Size", "Cleaned Size"]:
        data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0)

    result = pd.DataFrame(
        {
            "file": data["File"].map(normalize_file_name),
            "sheet": data["Sheet / Workbook"].map(normalize_sheet_name),
            "python_original_bytes": data["Original Size"],
            "python_cleaned_bytes": data["Cleaned Size"],
        }
    )
    result["python_reduction_bytes"] = (
        result["python_original_bytes"] - result["python_cleaned_bytes"]
    ).clip(lower=0)

    return result


def read_long_analysis(path: Path) -> pd.DataFrame:
    data = pd.read_csv(path)

    for column in [
        "original_sheet_xml_bytes",
        "cleaned_sheet_xml_bytes",
        "removed_bytes_uncompressed_xml",
        "added_bytes_uncompressed_xml",
    ]:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0)

    group_keys = ["input_file", "sheet_file"]
    result = (
        data.groupby(group_keys, as_index=False)
        .agg(
            long_original_bytes=("original_sheet_xml_bytes", "max"),
            long_cleaned_bytes=("cleaned_sheet_xml_bytes", "min"),
            long_cause_removed_bytes=("removed_bytes_uncompressed_xml", "sum"),
            long_cause_added_bytes=("added_bytes_uncompressed_xml", "sum"),
        )
    )
    result["long_reduction_bytes"] = (
        result["long_original_bytes"] - result["long_cleaned_bytes"]
    ).clip(lower=0)
    result["file"] = result["input_file"].map(normalize_file_name)

    sheet_name_cache: dict[str, dict[str, str]] = {}

    def resolve_sheet(row: pd.Series) -> str:
        sheet_file = normalize_zip_path(str(row["sheet_file"]))
        if sheet_file == "xl/workbook.xml":
            return "Workbook"

        input_file = Path(str(row["input_file"]))
        cache_key = str(input_file)
        if cache_key not in sheet_name_cache:
            sheet_name_cache[cache_key] = build_sheet_path_to_name_map(input_file)

        return sheet_name_cache[cache_key].get(sheet_file, sheet_file)

    result["sheet"] = result.apply(resolve_sheet, axis=1)

    return result


def compare_reductions(
    python_data: pd.DataFrame,
    long_data: pd.DataFrame,
) -> pd.DataFrame:
    merged = python_data.merge(
        long_data,
        on=["file", "sheet"],
        how="outer",
        indicator=True,
    )

    for column in [
        "python_original_bytes",
        "python_cleaned_bytes",
        "python_reduction_bytes",
        "long_original_bytes",
        "long_cleaned_bytes",
        "long_reduction_bytes",
    ]:
        merged[column] = pd.to_numeric(merged[column], errors="coerce")

    merged["reduction_delta_bytes"] = (
        merged["long_reduction_bytes"] - merged["python_reduction_bytes"]
    )
    merged["analyzer_extra_reduction_bytes"] = (
        merged["python_reduction_bytes"] - merged["long_reduction_bytes"]
    )
    merged["rust_extra_reduction_bytes"] = (
        merged["long_reduction_bytes"] - merged["python_reduction_bytes"]
    )
    merged["abs_reduction_delta_bytes"] = merged["reduction_delta_bytes"].abs()
    merged["original_delta_bytes"] = (
        merged["long_original_bytes"] - merged["python_original_bytes"]
    )
    merged["cleaned_delta_bytes"] = (
        merged["long_cleaned_bytes"] - merged["python_cleaned_bytes"]
    )
    merged["same_reduction_bytes"] = merged["reduction_delta_bytes"].fillna(0) == 0

    return merged.sort_values(
        by=["analyzer_extra_reduction_bytes", "file", "sheet"],
        ascending=[False, True, True],
        na_position="last",
    )


def top_analyzer_more_entries(comparison: pd.DataFrame, count: int) -> pd.DataFrame:
    return comparison[
        (comparison["_merge"] == "both")
        & (comparison["analyzer_extra_reduction_bytes"].fillna(0) > 0)
    ].head(count).copy()


def extract_xlsx_member_for_diff(
    xlsx_path: Path,
    member_path: str,
    output_root: Path,
) -> Path:
    normalized_member_path = normalize_zip_path(member_path)
    output_path = output_root / normalized_member_path
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(xlsx_path) as archive:
        with archive.open(normalized_member_path) as source:
            output_path.write_bytes(source.read())

    return output_path


def format_xml_for_diff(xml_path: Path) -> None:
    text = xml_path.read_text(encoding="utf-8", errors="replace")

    try:
        root = ET.fromstring(text)
        ET.indent(root, space="  ")
        formatted = strip_namespaces_from_xml(ET.tostring(root, encoding="unicode"))
        xml_path.write_text(formatted + "\n", encoding="utf-8", newline="\n")
        return
    except ET.ParseError:
        pass

    formatted = re.sub(r"><(row\b)", r">\n<\1", text)
    formatted = re.sub(r"><(col\b)", r">\n<\1", formatted)
    formatted = re.sub(r"><(c\b)", r">\n<\1", formatted)
    xml_path.write_text(
        strip_namespaces_from_xml(formatted),
        encoding="utf-8",
        newline="\n",
    )


def strip_namespaces_from_xml(xml_text: str) -> str:
    result = re.sub(r'\s+xmlns(:\w+)?="[^"]*"', "", xml_text)
    result = re.sub(r"\s+\w+:(\w+)=", r" \1=", result)
    result = re.sub(r"</\w+:", "</", result)
    result = re.sub(r"<(\w+):", "<", result)
    return result


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    if ":" in tag:
        return tag.rsplit(":", 1)[1]
    return tag


def run_git_diff(left_path: Path, right_path: Path) -> str:
    command = [
        "git",
        "-c",
        "core.autocrlf=false",
        "-c",
        "core.safecrlf=false",
        "diff",
        "--no-index",
        "--no-ext-diff",
        "--",
        str(left_path),
        str(right_path),
    ]
    completed = subprocess.run(
        command,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )

    if completed.returncode not in (0, 1):
        return (
            f"git diff failed with exit code {completed.returncode}\n"
            f"{completed.stdout}"
        )

    return completed.stdout


def changed_cells_from_diff(
    diff_text: str,
) -> Iterable[tuple[str, str, str | None]]:
    cell_pattern = re.compile(r'^[+-]\s*<[^:>\s]*(?::)?c\b([^>]*)>')

    for line in diff_text.splitlines():
        if line.startswith(("+++", "---")):
            continue

        match = cell_pattern.search(line)
        if match:
            attributes = match.group(1)
            cell_reference_match = re.search(r'\br="([^"]+)"', attributes)
            if cell_reference_match is None:
                continue

            style_match = re.search(r'\bs="([^"]+)"', attributes)
            style_id = style_match.group(1) if style_match is not None else None
            yield line[0], cell_reference_match.group(1), style_id


def first_changed_cell_from_diff(diff_text: str) -> tuple[str, str, str | None] | None:
    return next(iter(changed_cells_from_diff(diff_text)), None)


def column_letters_to_index(letters: str) -> int:
    result = 0
    for character in letters.upper():
        if not "A" <= character <= "Z":
            break
        result = result * 26 + (ord(character) - ord("A") + 1)
    return result


def cell_reference_parts(cell_reference: str) -> tuple[int, str] | None:
    match = re.match(r"^\$?([A-Z]+)\$?([0-9]+)$", cell_reference, re.IGNORECASE)
    if not match:
        return None

    column_letters, row_number = match.groups()
    return int(row_number), column_letters.upper()


def cell_reference_to_indexes(cell_reference: str) -> tuple[int, int] | None:
    parts = cell_reference_parts(cell_reference)
    if parts is None:
        return None

    row_index, column_letters = parts
    return row_index, column_letters_to_index(column_letters)


def sqref_token_contains_cell(token: str, cell_reference: str) -> bool | None:
    target = cell_reference_to_indexes(cell_reference)
    if target is None:
        return None

    target_row, target_column = target
    references = token.split(":")
    if len(references) == 1:
        reference = cell_reference_to_indexes(references[0])
        if reference is None:
            return None
        return reference == target

    if len(references) == 2:
        start = cell_reference_to_indexes(references[0])
        end = cell_reference_to_indexes(references[1])
        if start is None or end is None:
            return None

        start_row, start_column = start
        end_row, end_column = end
        return (
            min(start_row, end_row) <= target_row <= max(start_row, end_row)
            and min(start_column, end_column)
            <= target_column
            <= max(start_column, end_column)
        )

    return None


def sqref_contains_cell(sqref: str, cell_reference: str) -> bool | None:
    tokens = sqref.split()
    if not tokens:
        return None

    saw_unparseable = False
    for token in tokens:
        contains = sqref_token_contains_cell(token, cell_reference)
        if contains is True:
            return True
        if contains is None:
            saw_unparseable = True

    return None if saw_unparseable else False


def element_xml(element: ET.Element) -> str:
    return strip_namespaces_from_xml(
        ET.tostring(element, encoding="unicode", short_empty_elements=True)
    )


REFERENCED_ELEMENT_ATTRIBUTES = {
    "conditionalFormatting": "sqref",
    "dataValidation": "sqref",
    "hyperlink": "ref",
    "mergeCell": "ref",
    "sortCondition": "ref",
    "selection": "sqref",
}

REFERENCE_CONTAINER_CHILDREN = {
    "conditionalFormattings": "conditionalFormatting",
    "dataValidations": "dataValidation",
    "hyperlinks": "hyperlink",
    "mergeCells": "mergeCell",
}


def referenced_element_text(element: ET.Element) -> str | None:
    element_name = local_name(element.tag)
    attribute_name = REFERENCED_ELEMENT_ATTRIBUTES.get(element_name)
    if attribute_name is None:
        return None

    reference = element.attrib.get(attribute_name)
    if reference:
        return reference

    for child in element:
        if local_name(child.tag) == attribute_name:
            text = child.text.strip() if child.text is not None else ""
            if text:
                return text

    return None


def referenced_elements_for_cell(
    root: ET.Element,
    cell_reference: str,
) -> dict[str, dict[str, list[str]]]:
    result = {
        element_name: {"matches": [], "unparseable": []}
        for element_name in REFERENCED_ELEMENT_ATTRIBUTES
    }

    for element in root.iter():
        element_name = local_name(element.tag)
        if element_name not in REFERENCED_ELEMENT_ATTRIBUTES:
            continue

        reference = referenced_element_text(element)
        if reference is None:
            result[element_name]["unparseable"].append(element_xml(element))
            continue

        contains = sqref_contains_cell(reference, cell_reference)
        if contains is True:
            result[element_name]["matches"].append(element_xml(element))
        elif contains is None:
            result[element_name]["unparseable"].append(element_xml(element))

    return result


def remove_sheet_rows_and_cols(
    sheet_xml_path: Path,
    output_path: Path,
    relevant_cell_reference: str | None = None,
) -> None:
    text = sheet_xml_path.read_text(encoding="utf-8", errors="replace")

    try:
        root = ET.fromstring(text)
        for parent in root.iter():
            children_to_keep = []
            for child in list(parent):
                child_name = local_name(child.tag)
                if child_name in {"row", "col"}:
                    continue

                if child_name in REFERENCED_ELEMENT_ATTRIBUTES:
                    if relevant_cell_reference is None:
                        continue

                    reference = referenced_element_text(child)
                    contains = (
                        None
                        if reference is None
                        else sqref_contains_cell(reference, relevant_cell_reference)
                    )
                    if contains is False:
                        continue

                children_to_keep.append(child)

            parent[:] = children_to_keep

        prune_empty_reference_containers(root)
        ET.indent(root, space="  ")
        stripped = ET.tostring(root, encoding="unicode", short_empty_elements=True)
    except ET.ParseError:
        stripped = re.sub(r"<(?:\w+:)?row\b.*?</(?:\w+:)?row>", "", text, flags=re.S)
        stripped = re.sub(r"<(?:\w+:)?row\b[^>]*/>", "", stripped)
        stripped = re.sub(r"<(?:\w+:)?col\b[^>]*/>", "", stripped)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        strip_namespaces_from_xml(stripped) + "\n",
        encoding="utf-8",
        newline="\n",
    )


def prune_empty_reference_containers(element: ET.Element) -> bool:
    kept_children = []
    for child in list(element):
        if prune_empty_reference_containers(child):
            kept_children.append(child)

    element[:] = kept_children
    element_name = local_name(element.tag)
    child_name = REFERENCE_CONTAINER_CHILDREN.get(element_name)
    if child_name is None:
        return True

    matching_children = [
        child for child in element if local_name(child.tag) == child_name
    ]
    if not matching_children:
        return False

    if "count" in element.attrib:
        element.attrib["count"] = str(len(matching_children))

    return True


def read_sheet_context(
    xlsx_path: Path,
    sheet_file: str,
    cell_reference: str,
) -> dict[str, str | None]:
    parts = cell_reference_parts(cell_reference)
    if parts is None:
        return {
            "cell_type": None,
            "cell_value": None,
            "cell_style_id": None,
            "shared_string_text": None,
            "shared_string_xml": None,
            "row_xml": None,
            "row_style_id": None,
            "row_style_xf": None,
            "col_xml": None,
            "col_style_id": None,
            "col_style_xf": None,
            "referenced_elements": None,
        }

    row_index, column_letters = parts
    column_index = column_letters_to_index(column_letters)

    with zipfile.ZipFile(xlsx_path) as archive:
        sheet_xml = archive.read(sheet_file)

    root = ET.fromstring(sheet_xml)
    referenced_elements = referenced_elements_for_cell(
        root,
        cell_reference,
    )
    row_element = root.find(f".//{{{SPREADSHEET_NS}}}row[@r='{row_index}']")
    cell_element = root.find(f".//{{{SPREADSHEET_NS}}}c[@r='{cell_reference}']")
    col_element: ET.Element | None = None

    for candidate in root.findall(f".//{{{SPREADSHEET_NS}}}col"):
        min_text = candidate.attrib.get("min")
        max_text = candidate.attrib.get("max")
        if min_text is None or max_text is None:
            continue

        try:
            min_index = int(min_text)
            max_index = int(max_text)
        except ValueError:
            continue

        if min_index <= column_index <= max_index:
            col_element = candidate
            break

    row_style_id = row_element.attrib.get("s") if row_element is not None else None
    col_style_id = col_element.attrib.get("style") if col_element is not None else None
    shared_string_text = read_cell_shared_string(xlsx_path, cell_element)

    return {
        "cell_type": cell_element.attrib.get("t") if cell_element is not None else None,
        "cell_value": read_cell_value(cell_element),
        "cell_style_id": cell_element.attrib.get("s") if cell_element is not None else None,
        "shared_string_text": shared_string_text,
        "shared_string_xml": read_cell_shared_string_xml(xlsx_path, cell_element),
        "row_xml": element_xml(row_element) if row_element is not None else None,
        "row_style_id": row_style_id,
        "row_style_xf": read_cell_xf_from_styles(xlsx_path, row_style_id)
        if row_style_id is not None
        else None,
        "col_xml": element_xml(col_element) if col_element is not None else None,
        "col_style_id": col_style_id,
        "col_style_xf": read_cell_xf_from_styles(xlsx_path, col_style_id)
        if col_style_id is not None
        else None,
        "referenced_elements": referenced_elements,
    }


def read_cell_value(cell_element: ET.Element | None) -> str | None:
    if cell_element is None:
        return None

    value_element = cell_element.find(f"{{{SPREADSHEET_NS}}}v")
    if value_element is None:
        return None

    return value_element.text


def shared_string_text(shared_string_item: ET.Element) -> str:
    text_parts = [
        text_element.text or ""
        for text_element in shared_string_item.findall(f".//{{{SPREADSHEET_NS}}}t")
    ]
    return "".join(text_parts)


def read_cell_shared_string(
    xlsx_path: Path,
    cell_element: ET.Element | None,
) -> str | None:
    if cell_element is None or cell_element.attrib.get("t") != "s":
        return None

    value_text = read_cell_value(cell_element)
    if value_text is None:
        return None

    try:
        shared_string_index = int(value_text)
    except ValueError:
        return None

    try:
        with zipfile.ZipFile(xlsx_path) as archive:
            shared_strings_xml = archive.read("xl/sharedStrings.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None

    root = ET.fromstring(shared_strings_xml)
    shared_string_items = list(root.findall(f"{{{SPREADSHEET_NS}}}si"))
    if shared_string_index < 0 or shared_string_index >= len(shared_string_items):
        return None

    return shared_string_text(shared_string_items[shared_string_index])


def read_cell_shared_string_xml(
    xlsx_path: Path,
    cell_element: ET.Element | None,
) -> str | None:
    shared_string_item = read_cell_shared_string_item(xlsx_path, cell_element)
    if shared_string_item is None:
        return None
    return element_xml(shared_string_item)


def read_cell_shared_string_item(
    xlsx_path: Path,
    cell_element: ET.Element | None,
) -> ET.Element | None:
    if cell_element is None or cell_element.attrib.get("t") != "s":
        return None

    value_text = read_cell_value(cell_element)
    if value_text is None:
        return None

    try:
        shared_string_index = int(value_text)
    except ValueError:
        return None

    try:
        with zipfile.ZipFile(xlsx_path) as archive:
            shared_strings_xml = archive.read("xl/sharedStrings.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None

    root = ET.fromstring(shared_strings_xml)
    shared_string_items = list(root.findall(f"{{{SPREADSHEET_NS}}}si"))
    if shared_string_index < 0 or shared_string_index >= len(shared_string_items):
        return None

    return shared_string_items[shared_string_index]


def read_cell_xf_from_styles(xlsx_path: Path, style_index_text: str) -> str | None:
    try:
        style_index = int(style_index_text)
    except ValueError:
        return None

    try:
        with zipfile.ZipFile(xlsx_path) as archive:
            styles_xml = archive.read("xl/styles.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None

    root = ET.fromstring(styles_xml)
    cell_xfs = root.find(f"{{{SPREADSHEET_NS}}}cellXfs")
    if cell_xfs is None:
        return None

    xfs = list(cell_xfs.findall(f"{{{SPREADSHEET_NS}}}xf"))
    if style_index < 0 or style_index >= len(xfs):
        return None

    xf = xfs[style_index]
    return element_xml(xf)


def cell_style_applies_protection(xlsx_path: Path, style_id: str | None) -> bool:
    if not style_id:
        return False

    xf = read_cell_xf_from_styles(xlsx_path, style_id)
    if xf is None:
        return False

    return re.search(r'\bapplyProtection="1"', xf) is not None


def selected_unprotected_changed_cell_from_diff(
    diff_text: str,
    python_xlsx: Path,
    long_xlsx: Path,
    sheet_file: str,
) -> tuple[str, str, str | None] | None:
    selected: tuple[str, str, str | None] | None = None
    selected_column_index = -1
    selected_row_index: int | None = None

    for side, cell_ref, style_id in changed_cells_from_diff(diff_text):
        source_xlsx = long_xlsx if side == "+" else python_xlsx
        if style_id is None:
            context = read_sheet_context(source_xlsx, sheet_file, cell_ref)
            style_id = context["cell_style_id"]

        if cell_style_applies_protection(source_xlsx, style_id):
            continue

        parts = cell_reference_parts(cell_ref)
        if parts is None:
            continue

        row_index, column_letters = parts
        column_index = column_letters_to_index(column_letters)

        if selected_row_index is None:
            selected_row_index = row_index

        if row_index != selected_row_index:
            continue

        if column_index >= selected_column_index:
            selected = side, cell_ref, style_id
            selected_column_index = column_index

    return selected


def changed_cell_style_summary(
    diff_text: str,
    python_xlsx: Path,
    long_xlsx: Path,
    sheet_file: str,
    changed_cell: tuple[str, str, str | None] | None = None,
) -> str:
    if changed_cell is None:
        changed_cell = selected_unprotected_changed_cell_from_diff(
            diff_text,
            python_xlsx,
            long_xlsx,
            sheet_file,
        )
    if changed_cell is None:
        return "Selected changed cell: not found after ignoring applyProtection=\"1\" cells\n"

    side, cell_ref, style_id = changed_cell
    source_name = "long-analysis" if side == "+" else "python"
    source_xlsx = long_xlsx if side == "+" else python_xlsx
    context = read_sheet_context(source_xlsx, sheet_file, cell_ref)
    style_id = style_id or context["cell_style_id"]
    xf = read_cell_xf_from_styles(source_xlsx, style_id) if style_id else None
    shared_string = context["shared_string_text"]
    if shared_string == "":
        shared_string = "<empty string>"
    elif shared_string is None:
        shared_string = "none"

    lines = [
        f"Selected changed cell: {cell_ref}",
        "Selected changed cell rule: last changed cell in the first changed row",
        f"Selected changed cell side: {source_name}",
        f"Selected changed cell type: {context['cell_type'] or 'none'}",
        f"Selected changed cell value: {context['cell_value'] or 'none'}",
        f"Selected changed cell shared string: {shared_string}",
        f"Selected changed cell shared string xml: {context['shared_string_xml'] or 'none'}",
        f"Selected changed cell style id: {style_id or 'none'}",
    ]

    if xf is None:
        lines.append("Selected changed cell style xf: not found")
    else:
        lines.append(f"Selected changed cell style xf: {xf}")

    lines.extend(
        [
            f"Selected changed cell row xml: {context['row_xml'] or 'not found'}",
            f"Selected changed cell row style id: {context['row_style_id'] or 'none'}",
            f"Selected changed cell row style xf: {context['row_style_xf'] or 'none'}",
            f"Selected changed cell col xml: {context['col_xml'] or 'not found'}",
            f"Selected changed cell col style id: {context['col_style_id'] or 'none'}",
            f"Selected changed cell col style xf: {context['col_style_xf'] or 'none'}",
        ]
    )

    referenced_elements = context["referenced_elements"] or {}
    for element_name in REFERENCED_ELEMENT_ATTRIBUTES:
        element_result = referenced_elements.get(
            element_name,
            {"matches": [], "unparseable": []},
        )
        matches = element_result["matches"]
        unparseable = element_result["unparseable"]
        reference_name = REFERENCED_ELEMENT_ATTRIBUTES[element_name]

        lines.append(
            f"Selected changed cell matching {element_name} count: {len(matches)}"
        )
        if matches:
            lines.append(f"Selected changed cell matching {element_name}:")
            lines.extend(str(item) for item in matches)

        lines.append(
            f"Unparseable {element_name} {reference_name} count: "
            f"{len(unparseable)}"
        )
        if unparseable:
            lines.append(f"Unparseable {element_name} {reference_name} entries:")
            lines.extend(str(item) for item in unparseable)

    return "\n".join(lines) + "\n"


def first_lines(text: str, max_lines: int) -> tuple[str, int]:
    if max_lines <= 0:
        return text, 0

    lines = text.splitlines()
    omitted = max(0, len(lines) - max_lines)
    return "\n".join(lines[:max_lines]) + ("\n" if lines else ""), omitted


def diff_window_around_first_changed_cell(
    diff_text: str,
    max_lines: int,
    cell_reference: str | None = None,
) -> tuple[str, int, int]:
    if max_lines <= 0:
        return diff_text, 0, 0

    lines = diff_text.splitlines()
    if len(lines) <= max_lines:
        return diff_text + ("\n" if lines else ""), 0, 0

    changed_index = None
    if cell_reference:
        cell_pattern = re.compile(
            r'^[+-]\s*<[^:>\s]*(?::)?c\b[^>]*\br="'
            + re.escape(cell_reference)
            + r'"'
        )
    else:
        cell_pattern = re.compile(r'^[+-]\s*<[^:>\s]*(?::)?c\b[^>]*\br="[^"]+"')
    for index, line in enumerate(lines):
        if line.startswith(("+++", "---")):
            continue
        if cell_pattern.search(line):
            changed_index = index
            break

    if changed_index is None:
        start = 0
    else:
        start = max(0, changed_index - (max_lines // 2))
        start = min(start, max(0, len(lines) - max_lines))

    end = min(len(lines), start + max_lines)
    before_omitted = start
    after_omitted = len(lines) - end
    return "\n".join(lines[start:end]) + "\n", before_omitted, after_omitted


def write_limited_text_block(
    output,
    text: str,
    max_lines: int,
    block_name: str,
) -> None:
    limited_text, omitted = first_lines(text, max_lines)
    output.write(limited_text)
    if omitted:
        output.write(
            f"... omitted {omitted} more lines from {block_name} "
            f"(showing first {max_lines})\n"
        )


def write_limited_diff_block(
    output,
    diff_text: str,
    max_lines: int,
    block_name: str,
    cell_reference: str | None = None,
) -> None:
    limited_text, before_omitted, after_omitted = diff_window_around_first_changed_cell(
        diff_text,
        max_lines,
        cell_reference,
    )
    if before_omitted:
        output.write(
            f"... omitted {before_omitted} lines before selected changed cell "
            f"from {block_name}\n"
        )
    output.write(limited_text)
    if after_omitted:
        output.write(
            f"... omitted {after_omitted} lines after shown window "
            f"from {block_name} (showing {max_lines} lines around selected changed cell)\n"
        )


def write_top_file_diffs(
    comparison: pd.DataFrame,
    python_xlsx_dir: Path,
    long_xlsx_dir: Path,
    output_path: Path,
    count: int,
    max_diff_lines: int,
) -> None:
    entries = top_analyzer_more_entries(comparison, count)

    os.makedirs(output_path.parent, exist_ok=True)

    with output_path.open("w", encoding="utf-8", newline="\n") as output:
        output.write(
            "Top sheet XML diffs where the analyzer removed more than Rust\n"
            "Generated by compare_analysis_reductions.py\n"
            f"Python XLSX dir: {python_xlsx_dir}\n"
            f"Long-analysis XLSX dir: {long_xlsx_dir}\n"
            f"Requested sheet count: {count}\n"
            f"Shown lines per XML block/diff window: {max_diff_lines}\n"
            "Sheet XML git diff windows are centered around the selected changed cell when possible.\n"
            "Selected changed cell rule: last changed cell in the first changed row.\n"
            "Changed cells with style xf applyProtection=\"1\" are ignored for diagnostics and windowing.\n"
            f"Work dir: {output_path.parent / 'analysis_reduction_diff_work'}\n"
            "\n"
        )

        if entries.empty:
            output.write("No sheets found where the analyzer removed more than Rust.\n")
            return

        work_dir = output_path.parent / "analysis_reduction_diff_work"
        shutil.rmtree(work_dir, ignore_errors=True)
        work_dir.mkdir(parents=True, exist_ok=True)

        for index, (_, row) in enumerate(entries.iterrows(), start=1):
            file_name = str(row["file"])
            sheet_name = str(row["sheet"])
            sheet_file = normalize_zip_path(str(row.get("sheet_file", "")))

            if sheet_name == "Workbook":
                sheet_file = "xl/workbook.xml"

            python_xlsx = python_xlsx_dir / file_name
            long_xlsx = long_xlsx_dir / file_name

            output.write("\n")
            output.write("=" * 100)
            output.write("\n")
            output.write(f"{index}. {file_name} :: {sheet_name}\n")
            output.write(f"Sheet XML: {sheet_file}\n")
            output.write(
                f"Python reduction bytes: {row.get('python_reduction_bytes')}\n"
            )
            output.write(f"Long reduction bytes: {row.get('long_reduction_bytes')}\n")
            output.write(f"Delta bytes: {row.get('reduction_delta_bytes')}\n")
            output.write(
                "Analyzer extra reduction bytes: "
                f"{row.get('analyzer_extra_reduction_bytes')}\n"
            )
            output.write(f"Python XLSX: {python_xlsx}\n")
            output.write(f"Long-analysis XLSX: {long_xlsx}\n")
            output.write("=" * 100)
            output.write("\n\n")

            if not sheet_file or sheet_file == "nan":
                output.write("Missing sheet XML path in comparison row.\n")
                continue

            if not python_xlsx.is_file():
                output.write(f"Missing Python XLSX file: {python_xlsx}\n")
                continue

            if not long_xlsx.is_file():
                output.write(f"Missing long-analysis XLSX file: {long_xlsx}\n")
                continue

            entry_work_dir = work_dir / f"{index:02d}_{Path(file_name).stem}"
            left_root = entry_work_dir / "python"
            right_root = entry_work_dir / "long"

            try:
                left_xml = extract_xlsx_member_for_diff(
                    python_xlsx,
                    sheet_file,
                    left_root,
                )
                right_xml = extract_xlsx_member_for_diff(
                    long_xlsx,
                    sheet_file,
                    right_root,
                )
            except KeyError as error:
                output.write(f"Missing sheet XML in XLSX container: {error}\n")
                continue
            except zipfile.BadZipFile as error:
                output.write(f"Cannot extract XLSX: {error}\n")
                continue

            format_xml_for_diff(left_xml)
            format_xml_for_diff(right_xml)

            diff_text = run_git_diff(left_xml, right_xml)
            if diff_text:
                changed_cell = selected_unprotected_changed_cell_from_diff(
                    diff_text,
                    python_xlsx,
                    long_xlsx,
                    sheet_file,
                )
                changed_cell_ref = changed_cell[1] if changed_cell is not None else None
                left_naked_xml = entry_work_dir / "python_naked_sheet.xml"
                right_naked_xml = entry_work_dir / "long_naked_sheet.xml"
                remove_sheet_rows_and_cols(
                    left_xml,
                    left_naked_xml,
                    changed_cell_ref,
                )
                remove_sheet_rows_and_cols(
                    right_xml,
                    right_naked_xml,
                    changed_cell_ref,
                )

                output.write(
                    changed_cell_style_summary(
                        diff_text,
                        python_xlsx,
                        long_xlsx,
                        sheet_file,
                        changed_cell,
                    )
                )
                output.write("\n")
                output.write(
                    "Python naked sheet XML without row/col elements "
                    "and with referenced nodes filtered to the changed cell:\n"
                )
                output.write("-" * 100)
                output.write("\n")
                write_limited_text_block(
                    output,
                    left_naked_xml.read_text(encoding="utf-8"),
                    max_diff_lines,
                    "Python naked sheet XML",
                )
                output.write("\n")
                output.write(
                    "Long-analysis naked sheet XML without row/col elements "
                    "and with referenced nodes filtered to the changed cell:\n"
                )
                output.write("-" * 100)
                output.write("\n")
                write_limited_text_block(
                    output,
                    right_naked_xml.read_text(encoding="utf-8"),
                    max_diff_lines,
                    "Long-analysis naked sheet XML",
                )
                output.write("\n")
                output.write("Sheet XML git diff:\n")
                output.write("-" * 100)
                output.write("\n")
                write_limited_diff_block(
                    output,
                    strip_namespaces_from_xml(diff_text),
                    max_diff_lines,
                    "sheet XML git diff",
                    changed_cell_ref,
                )
            else:
                output.write("No sheet XML diff found.\n")


def print_summary(comparison: pd.DataFrame, limit: int) -> None:
    both = comparison[comparison["_merge"] == "both"]
    left_only = comparison[comparison["_merge"] == "left_only"]
    right_only = comparison[comparison["_merge"] == "right_only"]
    changed = both[both["abs_reduction_delta_bytes"] > 0]
    analyzer_more = both[both["analyzer_extra_reduction_bytes"] > 0]
    rust_more = both[both["rust_extra_reduction_bytes"] > 0]

    print("Comparison summary")
    print("=" * 80)
    print(f"Matched entries: {len(both)}")
    print(f"Only in Python analysis: {len(left_only)}")
    print(f"Only in long analysis: {len(right_only)}")
    print(f"Matched entries with changed reduction bytes: {len(changed)}")
    print(f"Analyzer removed more than Rust: {len(analyzer_more)}")
    print(f"Rust removed more than analyzer: {len(rust_more)}")

    if len(both) > 0:
        print(
            "Total Python reduction bytes:",
            int(both["python_reduction_bytes"].fillna(0).sum()),
        )
        print(
            "Total long-analysis reduction bytes:",
            int(both["long_reduction_bytes"].fillna(0).sum()),
        )
        print(
            "Total delta bytes:",
            int(
                both["long_reduction_bytes"].fillna(0).sum()
                - both["python_reduction_bytes"].fillna(0).sum()
            ),
        )

    print()
    print(f"Top {limit} cases where analyzer removed more than Rust")
    print("=" * 80)

    columns = [
        "file",
        "sheet",
        "python_reduction_bytes",
        "long_reduction_bytes",
        "analyzer_extra_reduction_bytes",
        "reduction_delta_bytes",
        "abs_reduction_delta_bytes",
        "python_original_bytes",
        "long_original_bytes",
        "_merge",
    ]
    print(
        comparison[
            (comparison["_merge"] == "both")
            & (comparison["analyzer_extra_reduction_bytes"].fillna(0) > 0)
        ][columns]
        .head(limit)
        .to_string(index=False)
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compare wide Python Excel Shrink Analyzer CSV output with a long-form "
            "analysis CSV such as the Rust analysis output."
        )
    )
    parser.add_argument(
        "--python-analysis-csv",
        type=Path,
        default=DEFAULT_PYTHON_ANALYSIS_CSV,
    )
    parser.add_argument(
        "--long-analysis-csv",
        type=Path,
        default=DEFAULT_LONG_ANALYSIS_CSV,
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=DEFAULT_OUTPUT_CSV,
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=50,
    )
    parser.add_argument(
        "--top-diff-count",
        type=int,
        default=100,
        help="Number of top changed file/sheet entries to diff as sheet XML.",
    )
    parser.add_argument(
        "--max-diff-lines",
        type=int,
        default=1000,
        help=(
            "Maximum lines shown for each naked sheet XML block and for the sheet XML "
            "diff window. The diff window is centered around the selected changed cell "
            "when possible. Extracted temp files remain complete."
        ),
    )
    parser.add_argument(
        "--diff-out",
        type=Path,
        default=DEFAULT_DIFF_OUTPUT,
        help="Output path for concatenated git diffs of top changed sheet XML files.",
    )
    parser.add_argument(
        "--python-xlsx-dir",
        type=Path,
        default=DEFAULT_PYTHON_XLSX_DIR,
        help="Directory containing XLSX files produced by the Python/analyzer pipeline.",
    )
    parser.add_argument(
        "--long-xlsx-dir",
        type=Path,
        default=DEFAULT_LONG_XLSX_DIR,
        help="Directory containing XLSX files produced by the long-analysis pipeline.",
    )
    parser.add_argument(
        "--skip-top-diffs",
        action="store_true",
        help="Only write the comparison CSV; do not create concatenated sheet XML diffs.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    python_data = read_python_analysis(args.python_analysis_csv)
    long_data = read_long_analysis(args.long_analysis_csv)
    comparison = compare_reductions(python_data, long_data)

    os.makedirs(args.out.parent, exist_ok=True)
    comparison.to_csv(args.out, index=False)

    print_summary(comparison, args.limit)
    print()
    print(f"Saved comparison CSV: {args.out}")

    if not args.skip_top_diffs:
        write_top_file_diffs(
            comparison=comparison,
            python_xlsx_dir=args.python_xlsx_dir,
            long_xlsx_dir=args.long_xlsx_dir,
            output_path=args.diff_out,
            count=args.top_diff_count,
            max_diff_lines=args.max_diff_lines,
        )
        print(f"Saved top sheet XML diffs: {args.diff_out}")


if __name__ == "__main__":
    main()
