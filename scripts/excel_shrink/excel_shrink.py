#!/usr/bin/env python3
import os
import zipfile
from typing import List, Dict
import time
import logging
import concurrent.futures
import hashlib
import urllib.request
from urllib.parse import urlparse
import certifi

from common.named_tupples import CleanSheetParameters, CleanWorkbookParameters, CleanXlsxParameters, XlsxFileToWrite

from excel_shrink_workers import clean_xml_data, generate_clean_sheet_params_for_xlsx_file

verbose: bool = False

def gather_xlsx_files_in_directory(
    path: str, recursive: bool = False
) -> List[str]:
    if os.path.isfile(path) and path.endswith(".xlsx"):
        return [path]
    
    if os.path.isfile(path) and not path.endswith(".xlsx"):
        raise ValueError(f"Path {path} is not a valid xlsx file or directory.")

    if not os.path.isdir(path):
        raise ValueError(f"Path {path} is not a valid xlsx file or directory.")

    xlsx_files = []
    if recursive:
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.endswith(".xlsx"):
                    xlsx_files.append(os.path.join(root, file))
    else:
        xlsx_files = [
            os.path.join(path, file)
            for file in os.listdir(path)
            if file.endswith(".xlsx")
        ]
    return xlsx_files

def generate_clean_sheet_params_for_xlsx_files_multi_threaded(
    xlsx_files: List[str],
    max_workers: int = 1,
    sheet_names: List[str] | None = None,
    only_clean_workbook : bool = False) -> List[CleanXlsxParameters]:
    clean_xlsx_parameters = []

    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_file_path = {}
        for index, file_path in enumerate(xlsx_files):
            logging.info(f"Analysing: {file_path}")

            clean_xlsx_params_future = executor.submit(
                generate_clean_sheet_params_for_xlsx_file, file_path, sheet_names, only_clean_workbook
            )
            future_to_file_path[clean_xlsx_params_future] = file_path

        for clean_xlsx_params_future in concurrent.futures.as_completed(future_to_file_path.keys()):
            file_path = future_to_file_path[clean_xlsx_params_future]
            try:
                clean_xlsx_params = clean_xlsx_params_future.result()
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
            else:
                clean_xlsx_parameters.append(clean_xlsx_params)
                if len(future_to_file_path) % 100 == 0:
                    print(f"Remaining xlsx files to analyse: {len(future_to_file_path)}")
                del future_to_file_path[clean_xlsx_params_future]
                del clean_xlsx_params
                del clean_xlsx_params_future

    return clean_xlsx_parameters

def create_cleaned_xlsx_zip_handle(
    clean_xlsx_params: CleanXlsxParameters, output_path: str
) -> XlsxFileToWrite:
    if output_path.endswith(".xlsx"):
        output_zip_path = output_path
    else:
        output_zip_path = os.path.join(output_path, os.path.basename(clean_xlsx_params.file_path))

    try:
        zip_handle = zipfile.ZipFile(output_zip_path, "w")
    except PermissionError as e:
        raise PermissionError(f"Cannot write to {output_zip_path}. Check permissions or file locks.") from e
    with zipfile.ZipFile(clean_xlsx_params.file_path, "r") as zin:
        for info in clean_xlsx_params.other_zip_files:
            data = zin.read(info.filename)
            zip_handle.writestr(info.filename, data, info.compress_type, info.compress_level)

    remaining_files_to_write = [sheet_zip_file.filename for sheet_zip_file in clean_xlsx_params.sheet_zip_files] + [clean_xlsx_params.workbook_zip_info.filename]
    xlsx_file_to_write = XlsxFileToWrite(zip_handle, remaining_files_to_write)
    return xlsx_file_to_write

def create_output_directory(path: str, output_path: str) -> None:
    if not os.path.exists(output_path):
        parent_dir = os.path.dirname(output_path) if output_path.endswith(".xlsx") else output_path
        if parent_dir and not os.path.exists(parent_dir):
            os.makedirs(parent_dir)

def process_xlsx_files_multi_threaded(
    clean_xlsx_parameters_list: List[CleanXlsxParameters], output_path: str, max_workers: int = 1, chunk_size: int = 1024,
) -> None:
    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_clean_workbook_or_sheet_params: Dict[concurrent.futures.Future, CleanWorkbookParameters | CleanSheetParameters] = {}
        clean_params_to_xlsx_params: Dict[concurrent.futures.Future, CleanXlsxParameters] = {}
        clean_xlsx_params_to_writer: Dict[str, XlsxFileToWrite] = {}

        def retrieve_or_create_new_writer(clean_xlsx_params: CleanXlsxParameters) -> XlsxFileToWrite:
            if clean_xlsx_params.file_path not in clean_xlsx_params_to_writer:
                clean_xlsx_params_to_writer[clean_xlsx_params.file_path] = create_cleaned_xlsx_zip_handle(
                    clean_xlsx_params, output_path
                )
            return clean_xlsx_params_to_writer[clean_xlsx_params.file_path]

        for clean_xlsx_params in clean_xlsx_parameters_list:
            logging.info(f"Processing: {clean_xlsx_params.file_path}")

            clean_workbook_params = clean_xlsx_params.clean_workbook_parameters

            cleaned_workbook_future = executor.submit(clean_xml_data, clean_workbook_params, chunk_size)
            future_to_clean_workbook_or_sheet_params[cleaned_workbook_future] = clean_workbook_params
            clean_params_to_xlsx_params[cleaned_workbook_future] = clean_xlsx_params
            
            clean_sheet_parameters = clean_xlsx_params.clean_sheet_parameters
            for clean_data_params in clean_sheet_parameters:
                cleaned_data_future = executor.submit(clean_xml_data, clean_data_params, chunk_size)
                future_to_clean_workbook_or_sheet_params[cleaned_data_future] = clean_data_params
                clean_params_to_xlsx_params[cleaned_data_future] = clean_xlsx_params

        for cleaned_data_future in concurrent.futures.as_completed(future_to_clean_workbook_or_sheet_params):
            clean_data_params = future_to_clean_workbook_or_sheet_params.pop(cleaned_data_future)
            clean_xlsx_params = clean_params_to_xlsx_params.pop(cleaned_data_future)

            try:
                xlsx_file_writer = retrieve_or_create_new_writer(clean_xlsx_params)

                cleaned_data = cleaned_data_future.result()

                xlsx_file_writer.zip_handle.writestr(
                    clean_data_params.zip_info.filename,
                    cleaned_data,
                    clean_data_params.zip_info.compress_type,
                    clean_data_params.zip_info.compress_level,
                )

                xlsx_file_writer.remaining_files_to_write.remove(clean_data_params.zip_info.filename)
                if not xlsx_file_writer.remaining_files_to_write:
                    writer = clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path)
                    writer.zip_handle.close()
            except Exception as e:
                print(
                    f"Error processing sheet {clean_data_params.zip_info.filename} "
                    f"in file {clean_xlsx_params.file_path}: {e}"
                )
            else:
                del cleaned_data
                del cleaned_data_future

            len_remaining_inner_files = len(future_to_clean_workbook_or_sheet_params)
            if len_remaining_inner_files % 1000 == 0:
                logging.info(f"Remaining Xml Documents (sheet.xml or workbook.xml) {len_remaining_inner_files}")


def generate_clean_sheet_params_for_xlsx_files_in_directory_single_threaded(xlsx_files: List[str], sheet_names: List[str] | None = None, only_clean_workbook : bool = False) -> List[CleanXlsxParameters]:
    clean_xlsx_parameters = []

    for file_path in xlsx_files:
        try:
            logging.info(f"Analysing: {file_path}")
            clean_xlsx_params = generate_clean_sheet_params_for_xlsx_file(file_path, sheet_names, only_clean_workbook)
            clean_xlsx_parameters.append(clean_xlsx_params)
        except Exception as e:
            print(f"Error processing {file_path}: {e}")

    return clean_xlsx_parameters




def process_xlsx_files_single_threaded(clean_xlsx_parameters_list: List[CleanXlsxParameters], output_path: str, chunk_size: int) -> None:
    """
    Single-threaded version of the XLSX processing function.
    """
    clean_xlsx_params_to_writer: Dict[str, XlsxFileToWrite] = {}

    def retrieve_or_create_new_writer(clean_xlsx_params: CleanXlsxParameters) -> XlsxFileToWrite:
        if clean_xlsx_params.file_path not in clean_xlsx_params_to_writer:
            clean_xlsx_params_to_writer[clean_xlsx_params.file_path] = create_cleaned_xlsx_zip_handle(
                clean_xlsx_params,
                output_path
            )
        return clean_xlsx_params_to_writer[clean_xlsx_params.file_path]

    # We’ll track how many total XML segments (1 for the workbook + 1 for each sheet) need processing
    total_tasks = sum(
        1 + len(params.clean_sheet_parameters) for params in clean_xlsx_parameters_list
    )
    tasks_completed = 0

    for clean_xlsx_params in clean_xlsx_parameters_list:
        logging.info(f"Processing: {clean_xlsx_params.file_path}")
        workbook_params = clean_xlsx_params.clean_workbook_parameters
        try:
            cleaned_data = clean_xml_data(workbook_params, chunk_size)
        except Exception as e:
            print(
                f"Error processing sheet {workbook_params.zip_info.filename} "
                f"in file {clean_xlsx_params.file_path}: {e}"
            )
        else:
            writer = retrieve_or_create_new_writer(clean_xlsx_params)
            writer.zip_handle.writestr(
                workbook_params.zip_info.filename,
                cleaned_data,
                workbook_params.zip_info.compress_type,
                workbook_params.zip_info.compress_level,
            )
            writer.remaining_files_to_write.remove(workbook_params.zip_info.filename)

            # Close the ZIP if there are no more files to write
            if not writer.remaining_files_to_write:
                clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path)
                writer.zip_handle.close()

        tasks_completed += 1
        tasks_remaining = total_tasks - tasks_completed
        if tasks_remaining % 1000 == 0:
            logging.info(f"Remaining Xml Documents (sheet.xml or workbook.xml): {tasks_remaining}")

        # Now clean each sheet
        for clean_data_params in clean_xlsx_params.clean_sheet_parameters:
            try:
                cleaned_data = clean_xml_data(clean_data_params, chunk_size)
            except Exception as e:
                print(
                    f"Error processing sheet {clean_data_params.zip_info.filename} "
                    f"in file {clean_xlsx_params.file_path}: {e}"
                )
            else:
                writer = retrieve_or_create_new_writer(clean_xlsx_params)
                writer.zip_handle.writestr(
                    clean_data_params.zip_info.filename,
                    cleaned_data,
                    clean_data_params.zip_info.compress_type,
                    clean_data_params.zip_info.compress_level,
                )
                writer.remaining_files_to_write.remove(clean_data_params.zip_info.filename)

                # Close the ZIP if there are no more files to write
                if not writer.remaining_files_to_write:
                    clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path)
                    writer.zip_handle.close()

            tasks_completed += 1
            if tasks_completed % 1000 == 0:
                logging.info(f"Remaining Xml Documents (sheet.xml or workbook.xml): {total_tasks - tasks_completed}")

def is_url(path: str) -> bool:
    return path.startswith("http://") or path.startswith("https://")

def get_cache_dir() -> str:
    cache_dir = os.path.expanduser("~/.cache/excel-shrink")
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir

def download_file_to_cache(url: str) -> str:
    parsed = urlparse(url)
    # Eindeutiger Dateiname: host + Pfad + Hash
    basename = os.path.basename(parsed.path)
    hash_digest = hashlib.md5(url.encode("utf-8")).hexdigest()
    filename = f"{basename}_{hash_digest}.xlsx"
    cache_path = os.path.join(get_cache_dir(), filename)

    if not os.path.exists(cache_path):
        logging.info(f"Downloading {url} to cache...")
        os.environ["SSL_CERT_FILE"] = certifi.where()
        urllib.request.urlretrieve(url, cache_path)
    else:
        logging.info(f"Using cached file for {url}")

    return cache_path

def main(path: str, output_path: str, recursive: bool = False, disable_mp: bool = False, chunk_size: int = 1024, sheet_names: List[str] | None = None, only_clean_workbook : bool = False) -> None:
    start_time = time.time()

    if is_url(path):
        path = download_file_to_cache(path)

    create_output_directory(path, output_path)

    xlsx_files = gather_xlsx_files_in_directory(path, recursive=recursive)

    if disable_mp:
        clean_xlsx_parameters_list = generate_clean_sheet_params_for_xlsx_files_in_directory_single_threaded(xlsx_files, sheet_names, only_clean_workbook)
        process_xlsx_files_single_threaded(clean_xlsx_parameters_list, output_path, chunk_size)
    else:
        max_workers = os.cpu_count() or 1

        clean_xlsx_parameters_list = generate_clean_sheet_params_for_xlsx_files_multi_threaded(
            xlsx_files, max_workers, sheet_names, only_clean_workbook
        ) if len(xlsx_files) > 1 else generate_clean_sheet_params_for_xlsx_files_in_directory_single_threaded(xlsx_files, sheet_names, only_clean_workbook)

        process_xlsx_files_multi_threaded(clean_xlsx_parameters_list, output_path, max_workers=max_workers, chunk_size=chunk_size)

    end_time = time.time()
    elapsed_time = end_time - start_time
    logging.info(f"Script completed in {elapsed_time:.2f} seconds.")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Clean xlsx files.")
    parser.add_argument("path", help="Path to xlsx file or directory containing xlsx files.")
    parser.add_argument("output", help="Path to output directory or file.")
    parser.add_argument("--recursive", action="store_true", help="Recursively process directories.")
    parser.add_argument(
        "--disable-multiprocessing",
        action="store_true",
        help="Run in single-threaded mode (no multiprocessing).",
    )
    parser.add_argument("--logging-level", type=str, default="INFO", help="Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)")
    parser.add_argument("--chunk-size", type=int, default=16777216, help="Set the chunk size for processing XML data.")
    parser.add_argument("--sheets", nargs='+', default=None, help="List of sheet names to process (e.g. sheet1.xml sheet2.xml)")
    parser.add_argument("--only-clean-workbook", action="store_true", help="Only clean the workbook.xml file and not the sheet.xml files."
    )

    args = parser.parse_args()
    numeric_level = getattr(logging, args.logging_level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f"Invalid logging level: {args.logging_level}")
    logging.basicConfig(level=numeric_level)

    main(args.path, args.output, args.recursive, args.disable_multiprocessing, args.chunk_size, args.sheets, args.only_clean_workbook)
