#!/usr/bin/env python3
from collections import defaultdict
import os
import shutil
import threading
import zipfile
from typing import List, Dict, Deque, Iterable, Tuple

import time
import logging
import concurrent.futures
import hashlib
import urllib.request
from urllib.parse import urlparse
import certifi
import argparse

from common.named_tupples import (
    CleanSheetParameters,
    CleanWorkbookParameters,
    CleanXlsxParameters,
    XlsxFileToWrite,
)

from excel_shrink_workers import (
    clean_xml_data,
    generate_clean_sheet_params_for_xlsx_file,
)

verbose: bool = False


def gather_xlsx_files_in_directory(
    path: str, recursive: bool = False, force_overwrite: bool = False
) -> List[str]:
    if os.path.isfile(path) and path.lower().endswith(".xlsx"):
        return [path]

    if os.path.isfile(path) and not path.lower().endswith(".xlsx"):
        raise ValueError(f"Path {path} is not a valid xlsx file or directory.")

    if not os.path.isdir(path):
        raise ValueError(f"Path {path} is not a valid xlsx file or directory.")

    xlsx_files = []
    if recursive:
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.lower().endswith(".xlsx"):
                    xlsx_files.append(os.path.join(root, file))
    else:
        xlsx_files = [
            os.path.join(path, file)
            for file in os.listdir(path)
            if file.lower().endswith(".xlsx")
        ]
    return xlsx_files


def generate_clean_sheet_params_for_xlsx_files_multi_threaded(
    xlsx_files: List[str],
    max_workers: int = 1,
    sheet_names: List[str] | None = None,
    only_clean_workbook: bool = False,
    remove_whitespaces: bool = False,
) -> Dict[str, CleanXlsxParameters]:
    clean_xlsx_parameters: Dict[str, CleanXlsxParameters] = {}

    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_file_path = {}
        for index, file_path in enumerate(xlsx_files):
            logging.info(f"Analysing: {file_path}")

            clean_xlsx_params_future = executor.submit(
                generate_clean_sheet_params_for_xlsx_file,
                file_path,
                sheet_names,
                only_clean_workbook,
                remove_whitespaces,
            )
            future_to_file_path[clean_xlsx_params_future] = file_path

        for clean_xlsx_params_future in concurrent.futures.as_completed(
            future_to_file_path.keys()
        ):
            file_path = future_to_file_path[clean_xlsx_params_future]
            try:
                clean_xlsx_params = clean_xlsx_params_future.result()
            except Exception as e:
                logging.error(f"Error processing {file_path}: {e}")
            else:
                clean_xlsx_parameters[clean_xlsx_params.file_path] = clean_xlsx_params
                if len(future_to_file_path) % 100 == 0:
                    logging.log(
                        21,
                        f"Remaining xlsx files to analyse: {len(future_to_file_path)}",
                    )
                del future_to_file_path[clean_xlsx_params_future]
                del clean_xlsx_params
                del clean_xlsx_params_future

    return clean_xlsx_parameters


def create_cleaned_xlsx_zip_handle(
    clean_xlsx_params: CleanXlsxParameters,
    output_path: str,
    force_overwrite: bool = False,
) -> XlsxFileToWrite:
    if output_path.lower().endswith(".xlsx"):
        output_zip_path = output_path
    else:
        output_zip_path = os.path.join(
            output_path, os.path.basename(clean_xlsx_params.file_path)
        )

    if os.path.exists(output_zip_path):
        if not force_overwrite:
            raise FileExistsError(
                f"Output file already exists: {output_zip_path}. Use --force-overwrite to replace it."
            )
        os.remove(output_zip_path)

    try:
        zip_handle = zipfile.ZipFile(output_zip_path, "w")
    except PermissionError as e:
        raise PermissionError(
            f"Cannot write to {output_zip_path}. Check permissions or file locks."
        ) from e

    with zipfile.ZipFile(clean_xlsx_params.file_path, "r") as zin:
        for info in clean_xlsx_params.other_zip_files:
            data = zin.read(info.filename)
            zip_handle.writestr(
                info.filename, data, info.compress_type, info.compress_level
            )

    remaining_files_to_write = [
        sheet_zip_file.filename for sheet_zip_file in clean_xlsx_params.sheet_zip_files
    ] + [clean_xlsx_params.workbook_zip_info.filename]

    return XlsxFileToWrite(
        zip_handle,
        remaining_files_to_write,
        failed_sheet_files=[],
        successfully_written_sheet_files=[],
    )


def create_output_directory(output_path: str, force_overwrite: bool = False) -> None:
    if output_path.lower().endswith(".xlsx"):
        parent_dir = os.path.dirname(output_path)
        if parent_dir and not os.path.exists(parent_dir):
            os.makedirs(parent_dir, exist_ok=True)

        if os.path.exists(output_path):
            if not force_overwrite:
                raise FileExistsError(
                    f"Output file already exists: {output_path}. Use --force-overwrite to replace it."
                )

            try:
                os.remove(output_path)
            except PermissionError:
                raise
    else:
        # output_path is a directory
        os.makedirs(output_path, exist_ok=True)


def read_member_tolerant_crc(zipf: zipfile.ZipFile, zinfo: zipfile.ZipInfo) -> bytes:
    """
    Reads a ZIP member and *keeps* the bytes even if zipfile raises
    BadZipFile('Bad CRC-32 for file ...') at the end.
    """
    chunks = []
    zef = None
    try:
        zef = zipf.open(zinfo, "r")
        while True:
            try:
                chunk = zef.read(1024 * 1024)  # 1 MiB
            except zipfile.BadZipFile as e:
                # This is the CRC mismatch that happens at end-of-stream.
                if "Bad CRC-32 for file" in str(e):
                    break  # keep what we already read
                raise
            if not chunk:
                break
            chunks.append(chunk)
    finally:
        if zef is not None:
            try:
                zef.close()
            except zipfile.BadZipFile as e:
                # close() can also surface the CRC error
                if "Bad CRC-32 for file" not in str(e):
                    raise

    return b"".join(chunks)


def write_original_bytes_unchanged(
    clean_xlsx_params: CleanXlsxParameters,
    clean_data_params: CleanSheetParameters | CleanWorkbookParameters,
    xlsx_file_writer: XlsxFileToWrite,
    clean_xlsx_params_to_writer: Dict[str, XlsxFileToWrite],
):
    with zipfile.ZipFile(clean_xlsx_params.file_path, "r") as zin:
        original_bytes = read_member_tolerant_crc(zin, clean_data_params.zip_info)

        xlsx_file_writer.zip_handle.writestr(
            clean_data_params.zip_info.filename,
            original_bytes,
            clean_data_params.zip_info.compress_type,
            clean_data_params.zip_info.compress_level,
        )

        logging.log(
            21,
            f"Wrote Sheet: {clean_data_params.zip_info.filename}, {clean_xlsx_params.file_path} unchanged",
        )


def close_xlsx_file_writer_if_done(
    xlsx_file_writer: XlsxFileToWrite,
    finished_file_clean_data_params: CleanSheetParameters | CleanWorkbookParameters,
    clean_xlsx_params,
    clean_xlsx_params_to_writer: Dict[str, XlsxFileToWrite],
):

    logging.log(
        21,
        f"Remaining: {len(xlsx_file_writer.remaining_files_to_write)} in {clean_xlsx_params.file_path}",
    )
    if not xlsx_file_writer.remaining_files_to_write:
        writer = clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path)
        writer.zip_handle.close()

        logging.log(
            21,
            f"Wrote File {clean_xlsx_params.file_path}",
        )


def process_xlsx_files_multi_threaded(
    clean_xlsx_parameters_list: Dict[str, CleanXlsxParameters],
    output_path: str,
    max_workers: int = 1,
    chunk_size: int = 1024,
    max_in_flight_per_worker: int = 8,
    writer_workers: int = 4,
    max_pending_writes: int = 200,
    disable_sheetdata_splitting: bool = False,
    force_overwrite: bool = False,
):

    def iter_all_tasks() -> Iterable[CleanWorkbookParameters | CleanSheetParameters]:
        # Produce tasks lazily (does not allocate a huge list).
        for clean_xlsx_params in clean_xlsx_parameters_list.values():
            logging.info(f"Processing: {clean_xlsx_params.file_path}")

            yield clean_xlsx_params.clean_workbook_parameters
            for p in clean_xlsx_params.clean_sheet_parameters:
                yield p

    file_locks = defaultdict(threading.Lock)

    clean_xlsx_params_to_writer: Dict[str, XlsxFileToWrite] = {}

    def retrieve_or_create_new_writer(
        clean_xlsx_params: CleanXlsxParameters,
    ) -> XlsxFileToWrite:
        fp = clean_xlsx_params.file_path
        if fp not in clean_xlsx_params_to_writer:
            clean_xlsx_params_to_writer[fp] = create_cleaned_xlsx_zip_handle(
                clean_xlsx_params, output_path, force_overwrite=force_overwrite
            )
        return clean_xlsx_params_to_writer[fp]

    def finalise_xlsx_if_done(
        *,
        clean_xlsx_params: CleanXlsxParameters,
        xlsx_file_writer: XlsxFileToWrite,
    ) -> None:
        """
        Must be called under the per-file lock.
        Finalises the output XLSX when there are no remaining inner files to write.
        """
        if xlsx_file_writer.remaining_files_to_write:
            return  # not done yet

        # If nothing was successfully written, we replace the partially-written output with a copy of the original
        if (
            not xlsx_file_writer.successfully_written_sheet_files
            and not clean_xlsx_params.only_clean_workbook
        ):
            logging.warning(
                f"All sheets failed for {clean_xlsx_params.file_path}, copying original file"
            )
            xlsx_file_writer.zip_handle.close()

            output_zip_path: str = xlsx_file_writer.zip_handle.filename

            # Remove writer tracking so we don't accidentally reuse it
            clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path, None)

            shutil.copy2(clean_xlsx_params.file_path, output_zip_path)
            logging.log(21, f"Copied original file: {clean_xlsx_params.file_path}")

            # Ensure it's definitely gone
            clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path, None)
            return

        # Otherwise: some sheets were written successfully, but some failed -> write original bytes for failed sheets
        for failed_sheet_file in list(xlsx_file_writer.failed_sheet_files):
            try:
                write_original_bytes_unchanged(
                    clean_xlsx_params,
                    failed_sheet_file,
                    xlsx_file_writer,
                    clean_xlsx_params_to_writer,
                )

            except Exception as e2:
                logging.error(
                    f"Failed to write original bytes for {failed_sheet_file.zip_info.filename} "
                    f"in file {clean_xlsx_params.file_path}: {e2}"
                )
        xlsx_file_writer.failed_sheet_files.clear()

    def write_one(
        clean_data_params: CleanSheetParameters | CleanWorkbookParameters,
        cleaned_data: bytes,
    ):
        clean_xlsx_params = clean_xlsx_parameters_list[clean_data_params.xlsx_file_path]
        lock = file_locks[clean_xlsx_params.file_path]

        # IMPORTANT: everything that touches the zip handle + your writer state is under the lock
        with lock:
            xlsx_file_writer = retrieve_or_create_new_writer(clean_xlsx_params)

            xlsx_file_writer.zip_handle.writestr(
                clean_data_params.zip_info.filename,
                cleaned_data,
                clean_data_params.zip_info.compress_type,
                clean_data_params.zip_info.compress_level,
            )

            fn = clean_data_params.zip_info.filename
            if fn in xlsx_file_writer.remaining_files_to_write:
                xlsx_file_writer.remaining_files_to_write.remove(fn)
            if isinstance(clean_data_params, CleanSheetParameters):
                xlsx_file_writer.successfully_written_sheet_files.append(
                    clean_data_params.zip_info.filename
                )

            close_xlsx_file_writer_if_done(
                xlsx_file_writer,
                clean_data_params,
                clean_xlsx_params,
                clean_xlsx_params_to_writer,
            )

            finalise_xlsx_if_done(
                clean_xlsx_params=clean_xlsx_params,
                xlsx_file_writer=xlsx_file_writer,
            )

        # return something small so the future doesn't keep big objects
        return len(cleaned_data)

    max_in_flight = max(1, max_workers * max_in_flight_per_worker)

    with concurrent.futures.ProcessPoolExecutor(
        max_workers=max_workers
    ) as executor, concurrent.futures.ThreadPoolExecutor(
        max_workers=writer_workers
    ) as write_executor:

        pending_clean: dict[
            concurrent.futures.Future, CleanWorkbookParameters | CleanSheetParameters
        ] = {}
        pending_writes: set[concurrent.futures.Future] = set()

        task_iter = iter_all_tasks()

        def submit_next_clean_task() -> bool:
            try:
                p = next(task_iter)
            except StopIteration:
                return False
            fut = executor.submit(
                clean_xml_data, p, chunk_size, disable_sheetdata_splitting
            )
            pending_clean[fut] = p
            return True

        # Prime the pipeline (bounded)
        for _ in range(max_in_flight):
            if not submit_next_clean_task():
                break

        while pending_clean:
            done, _ = concurrent.futures.wait(
                pending_clean.keys(),
                return_when=concurrent.futures.FIRST_COMPLETED,
            )

            for fut in done:
                clean_data_params = pending_clean.pop(fut)

                try:
                    cleaned_data = fut.result()

                    # Submit write to thread pool (parallel across files)
                    wf = write_executor.submit(
                        write_one, clean_data_params, cleaned_data
                    )
                    pending_writes.add(wf)

                    # Drop references ASAP
                    del cleaned_data

                except Exception as e:
                    clean_xlsx_params = clean_xlsx_parameters_list[
                        clean_data_params.xlsx_file_path
                    ]
                    logging.error(
                        f"Error processing {clean_data_params.zip_info.filename} "
                        f"in file {clean_xlsx_params.file_path}: {e}"
                    )
                    lock = file_locks[clean_xlsx_params.file_path]
                    with lock:
                        xlsx_file_writer = retrieve_or_create_new_writer(
                            clean_xlsx_params
                        )
                        if (
                            clean_data_params.zip_info.filename
                            in xlsx_file_writer.remaining_files_to_write
                        ):
                            xlsx_file_writer.remaining_files_to_write.remove(
                                clean_data_params.zip_info.filename
                            )

                        if isinstance(clean_data_params, CleanSheetParameters):
                            xlsx_file_writer.failed_sheet_files.append(
                                clean_data_params
                            )

                        finalise_xlsx_if_done(
                            clean_xlsx_params=clean_xlsx_params,
                            xlsx_file_writer=xlsx_file_writer,
                        )

                # Refill clean pipeline (bounded)
                submit_next_clean_task()

                # Backpressure for writer queue (VERY important!)
                if len(pending_writes) >= max_pending_writes:
                    wdone, pending_writes = concurrent.futures.wait(
                        pending_writes,
                        return_when=concurrent.futures.FIRST_COMPLETED,
                    )
                    # call result() so exceptions surface and futures release internals
                    for w in wdone:
                        _ = w.result()

            # also periodically drain completed writer futures
            if pending_writes:
                wdone = {w for w in pending_writes if w.done()}
                for w in wdone:
                    pending_writes.remove(w)
                    _ = w.result()

        # wait for all writes to finish
        for w in concurrent.futures.as_completed(pending_writes):
            _ = w.result()


def generate_clean_sheet_params_for_xlsx_files_in_directory_single_threaded(
    xlsx_files: List[str],
    sheet_names: List[str] | None = None,
    only_clean_workbook: bool = False,
    remove_whitespaces: bool = False,
) -> Dict[str, CleanXlsxParameters]:
    clean_xlsx_parameters: Dict[str, CleanXlsxParameters] = {}

    for file_path in xlsx_files:
        try:
            logging.info(f"Analysing: {file_path}")
            clean_xlsx_params = generate_clean_sheet_params_for_xlsx_file(
                file_path, sheet_names, only_clean_workbook, remove_whitespaces
            )
            clean_xlsx_parameters[clean_xlsx_params.file_path] = clean_xlsx_params
        except Exception as e:
            logging.error(f"Error processing {file_path}: {e}")

    return clean_xlsx_parameters


def process_xlsx_files_single_threaded(
    clean_xlsx_parameters_list: Dict[str, CleanXlsxParameters],
    output_path: str,
    chunk_size: int,
    disable_sheetdata_splitting: bool = False,
    force_overwrite: bool = False,
) -> None:
    """
    Single-threaded version of the XLSX processing function.
    """
    clean_xlsx_params_to_writer: Dict[str, XlsxFileToWrite] = {}

    def retrieve_or_create_new_writer(
        clean_xlsx_params: CleanXlsxParameters,
    ) -> XlsxFileToWrite:
        fp = clean_xlsx_params.file_path
        if fp not in clean_xlsx_params_to_writer:
            clean_xlsx_params_to_writer[fp] = create_cleaned_xlsx_zip_handle(
                clean_xlsx_params, output_path, force_overwrite=force_overwrite
            )
        return clean_xlsx_params_to_writer[fp]

    def finalise_xlsx_if_done(
        *, clean_xlsx_params: CleanXlsxParameters, xlsx_file_writer: XlsxFileToWrite
    ) -> None:
        if xlsx_file_writer.remaining_files_to_write:
            return

        if (
            not xlsx_file_writer.successfully_written_sheet_files
            and not clean_xlsx_params.only_clean_workbook
        ):
            logging.warning(
                f"All sheets failed for {clean_xlsx_params.file_path}, copying original file"
            )
            xlsx_file_writer.zip_handle.close()
            output_zip_path: str = xlsx_file_writer.zip_handle.filename

            # Remove writer tracking so we don't accidentally reuse it
            clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path, None)

            shutil.copy2(clean_xlsx_params.file_path, output_zip_path)
            logging.log(21, f"Copied original file: {clean_xlsx_params.file_path}")

            clean_xlsx_params_to_writer.pop(clean_xlsx_params.file_path, None)
            return

        for failed_params in list(xlsx_file_writer.failed_sheet_files):
            try:
                write_original_bytes_unchanged(
                    clean_xlsx_params,
                    failed_params,  # expects object with .zip_info.filename (same as your multi-process)
                    xlsx_file_writer,
                    clean_xlsx_params_to_writer,
                )
            except Exception as e2:
                logging.error(
                    f"Failed to write original bytes for {failed_params.zip_info.filename} "
                    f"in file {clean_xlsx_params.file_path}: {e2}"
                )
        xlsx_file_writer.failed_sheet_files.clear()

    def mark_done_and_maybe_finalise(
        clean_xlsx_params: CleanXlsxParameters,
        writer: XlsxFileToWrite,
        clean_data_params: CleanSheetParameters | CleanWorkbookParameters,
        *,
        sheet_succeeded: bool,
    ) -> None:
        fn = clean_data_params.zip_info.filename

        if fn in writer.remaining_files_to_write:
            writer.remaining_files_to_write.remove(fn)

        if isinstance(clean_data_params, CleanSheetParameters):
            if sheet_succeeded:
                writer.successfully_written_sheet_files.append(fn)
            else:
                writer.failed_sheet_files.append(clean_data_params)

        # If you have this helper in your codebase (as in multi-process), keep using it
        close_xlsx_file_writer_if_done(
            writer,
            clean_data_params,
            clean_xlsx_params,
            clean_xlsx_params_to_writer,
        )

        finalise_xlsx_if_done(
            clean_xlsx_params=clean_xlsx_params, xlsx_file_writer=writer
        )

    total_tasks = sum(
        1 + len(params.clean_sheet_parameters)
        for params in clean_xlsx_parameters_list.values()
    )
    tasks_completed = 0

    for clean_xlsx_params in clean_xlsx_parameters_list.values():
        logging.info(f"Processing: {clean_xlsx_params.file_path}")

        workbook_params = clean_xlsx_params.clean_workbook_parameters
        writer = retrieve_or_create_new_writer(clean_xlsx_params)

        try:
            cleaned_data = clean_xml_data(
                workbook_params,
                chunk_size,
                disable_sheetdata_splitting,
            )
        except Exception as e:
            logging.error(
                f"Error processing {workbook_params.zip_info.filename} "
                f"in file {clean_xlsx_params.file_path}: {e}"
            )
            mark_done_and_maybe_finalise(
                clean_xlsx_params, writer, workbook_params, sheet_succeeded=False
            )
        else:
            writer.zip_handle.writestr(
                workbook_params.zip_info.filename,
                cleaned_data,
                workbook_params.zip_info.compress_type,
                workbook_params.zip_info.compress_level,
            )
            del cleaned_data
            mark_done_and_maybe_finalise(
                clean_xlsx_params, writer, workbook_params, sheet_succeeded=True
            )

        tasks_completed += 1
        if (total_tasks - tasks_completed) % 1000 == 0:
            logging.log(
                21,
                f"Remaining Xml Documents (sheet.xml or workbook.xml): {total_tasks - tasks_completed}",
            )

        for clean_data_params in clean_xlsx_params.clean_sheet_parameters:
            writer = retrieve_or_create_new_writer(clean_xlsx_params)

            try:
                cleaned_data = clean_xml_data(
                    clean_data_params,
                    chunk_size,
                    disable_sheetdata_splitting,
                )
            except Exception as e:
                logging.error(
                    f"Error processing {clean_data_params.zip_info.filename} "
                    f"in file {clean_xlsx_params.file_path}: {e}"
                )
                mark_done_and_maybe_finalise(
                    clean_xlsx_params, writer, clean_data_params, sheet_succeeded=False
                )
            else:
                writer.zip_handle.writestr(
                    clean_data_params.zip_info.filename,
                    cleaned_data,
                    clean_data_params.zip_info.compress_type,
                    clean_data_params.zip_info.compress_level,
                )
                del cleaned_data
                mark_done_and_maybe_finalise(
                    clean_xlsx_params, writer, clean_data_params, sheet_succeeded=True
                )

            tasks_completed += 1
            if tasks_completed % 1000 == 0:
                logging.log(
                    21,
                    f"Remaining Xml Documents (sheet.xml or workbook.xml): {total_tasks - tasks_completed}",
                )


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


def main(
    path: str,
    output_path: str,
    recursive: bool = False,
    disable_mp: bool = False,
    chunk_size: int = 1024,
    sheet_names: List[str] | None = None,
    only_clean_workbook: bool = False,
    remove_whitespaces: bool = False,
    disable_sheetdata_splitting: bool = False,
    force_overwrite: bool = False,
) -> None:
    start_time = time.time()

    if is_url(path):
        path = download_file_to_cache(path)

    create_output_directory(output_path, force_overwrite=force_overwrite)

    xlsx_files = gather_xlsx_files_in_directory(path, recursive=recursive)

    if not force_overwrite:
        xlsx_files = [
            f
            for f in xlsx_files
            if not os.path.exists(os.path.join(output_path, os.path.basename(f)))
        ]

    if disable_mp:
        clean_xlsx_parameters_list = (
            generate_clean_sheet_params_for_xlsx_files_in_directory_single_threaded(
                xlsx_files, sheet_names, only_clean_workbook, remove_whitespaces
            )
        )
        process_xlsx_files_single_threaded(
            clean_xlsx_parameters_list,
            output_path,
            chunk_size,
            disable_sheetdata_splitting,
            force_overwrite=force_overwrite,
        )
    else:
        max_workers = os.cpu_count() or 1

        clean_xlsx_parameters_list = (
            generate_clean_sheet_params_for_xlsx_files_multi_threaded(
                xlsx_files,
                max_workers,
                sheet_names,
                only_clean_workbook,
                remove_whitespaces,
            )
            if len(xlsx_files) > 1
            else generate_clean_sheet_params_for_xlsx_files_in_directory_single_threaded(
                xlsx_files, sheet_names, only_clean_workbook, remove_whitespaces
            )
        )

        process_xlsx_files_multi_threaded(
            clean_xlsx_parameters_list,
            output_path,
            max_workers=max_workers,
            chunk_size=chunk_size,
            disable_sheetdata_splitting=disable_sheetdata_splitting,
            force_overwrite=force_overwrite,
        )

    end_time = time.time()
    elapsed_time = end_time - start_time
    logging.info(f"Script completed in {elapsed_time:.2f} seconds.")


if __name__ == "__main__":
    IMPORTANT_LEVEL = 21
    logging.addLevelName(IMPORTANT_LEVEL, "IMPORTANT")

    parser = argparse.ArgumentParser(description="Clean xlsx files.")
    parser.add_argument(
        "path", help="Path to xlsx file or directory containing xlsx files."
    )
    parser.add_argument("output", help="Path to output directory or file.")
    parser.add_argument(
        "--recursive", action="store_true", help="Recursively process directories."
    )
    parser.add_argument(
        "--disable-multiprocessing",
        action="store_true",
        help="Run in single-threaded mode (no multiprocessing).",
    )
    parser.add_argument(
        "--logging-level",
        type=str,
        default="INFO",
        help="Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=16777216,
        help="Set the chunk size for processing XML data.",
    )
    parser.add_argument(
        "--sheets",
        nargs="+",
        default=None,
        help="List of sheet names to process (e.g. sheet1.xml sheet2.xml)",
    )
    parser.add_argument(
        "--only-clean-workbook",
        action="store_true",
        help="Only clean the workbook.xml file and not the sheet.xml files.",
    )

    parser.add_argument(
        "--remove-whitespaces",
        action="store_true",
        help=(
            "Remove whitespace-only cell contents from worksheet files. "
            "WARNING: This may change formula behaviour, as a whitespace is treated as "
            "cell content (unlike an empty cell). It can also alter the visual layout, "
            "since whitespaces are sometimes used to prevent text overflow into adjacent cells "
            "instead of using cell alignment (Horizontal: Fill)."
        ),
    )

    parser.add_argument(
        "--disable-sheetdata-splitting",
        action="store_true",
        help="Disable splitting of sheet data range with cells and range with visible rows.",
    )

    parser.add_argument(
        "--force-overwrite",
        action="store_true",
        help="Force overwrite existing files.",
    )

    args = parser.parse_args()
    level_arg = args.logging_level.upper()

    # Try numeric first (e.g. "21")
    if level_arg.isdigit():
        numeric_level = int(level_arg)
    else:
        # Resolve via logging's internal name registry
        numeric_level = logging.getLevelName(level_arg)
        if isinstance(numeric_level, str):
            raise ValueError(f"Invalid logging level: {args.logging_level}")
    logging.basicConfig(level=numeric_level)

    main(
        args.path,
        args.output,
        args.recursive,
        args.disable_multiprocessing,
        args.chunk_size,
        args.sheets,
        args.only_clean_workbook,
        args.remove_whitespaces,
        args.disable_sheetdata_splitting,
        args.force_overwrite,
    )
