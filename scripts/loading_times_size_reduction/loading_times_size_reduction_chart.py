#!/usr/bin/env python3

from importlib.metadata import files
import os
import sys
import subprocess
import urllib.request
import logging
import psutil
import argparse
import time
import json
import shutil
import textwrap
import tempfile
from openpyxl import load_workbook
import pandas as pd
import csv
import certifi
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.transforms as mtransforms
from matplotlib.gridspec import GridSpec, GridSpecFromSubplotSpec
from collections import OrderedDict

plt.switch_backend("agg")

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

chart_output_path: str = os.path.join(
    "scripts",
    "loading_times_size_reduction",
    "charts",
    "time_and_filesize_comparison_by_file.pdf",
)


def measure_one_main(args):
    """
    measure-one mode entry point:
    --library "openpyxl(default)" etc.
    --file "some.xlsx"
    We run the library-specific open/save code, then print JSON to stdout.
    """
    library_name = args.library
    file_path = args.file

    # Dictionary of library name -> measurement function
    lib_funcs = {
        "openpyxl(default)": benchmark_openpyxl_default,
        "pandas": benchmark_pandas,
        "Microsoft Excel": benchmark_excel_com,
        "R openxlsx": benchmark_r_openxlsx,
        "R readxl+writexl": benchmark_r_readxl_writexl,
    }

    if library_name not in lib_funcs:
        logging.error(f"Library '{library_name}' not recognized.")
        print(json.dumps({"error": "Unknown library"}))
        sys.exit(1)

    open_t, save_t = lib_funcs[library_name](file_path)

    result = {
        "library": library_name,
        "file": os.path.basename(file_path),
        "Original Open Time (s)": open_t,
        "Original Save Time (s)": save_t,
    }
    print(json.dumps(result))


def _bench_template(name, opener, saver):  # ### NEW helper
    try:
        t0 = time.perf_counter()
        opener()
        t_open = time.perf_counter() - t0

        t0 = time.perf_counter()
        saver()
        t_save = time.perf_counter() - t0

        return t_open, t_save
    except Exception as e:
        logging.error(f"{name} error: {e}")
        return None, None


def benchmark_openpyxl_default(file_path):
    """
    Measure openpyxl.load_workbook + wb.save() with standard settings.
    """
    wb = None

    def _open():
        nonlocal wb
        wb = load_workbook(file_path)

    def _save():
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            p = tmp.name
        wb.save(p)
        os.remove(p)

    return _bench_template("openpyxl", _open, _save)


def benchmark_pandas(file_path):
    """
    Benchmark: Pandas reads ALL sheets with read_excel, then writes them back.
    """
    dfs = {}

    def _open():
        nonlocal dfs
        dfs = pd.read_excel(file_path, sheet_name=None)

    def _save():
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            p = tmp.name
        with pd.ExcelWriter(p, engine="openpyxl") as writer:
            for sheet, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        os.remove(p)

    return _bench_template("pandas", _open, _save)


# ---------------------------------------------------------------------------


def benchmark_excel_com(file_path):
    """
    Benchmark: via win32com (only on Windows).
    First runs excel_shrink.py --only-clean-workbook on the original,
    then opens the cleaned workbook with Excel COM and saves it.
    """

    name = "Microsoft Excel"
    try:
        import win32com.client
    except ImportError:
        logging.warning("win32com not available - skipping Excel COM benchmark.")
        return None, None, None

    with tempfile.TemporaryDirectory() as shrink_dir:

        excel, wb = None, None

        def _open():
            nonlocal excel, wb

            logging.info(f"Running excel_shrink --only-clean-workbook on {file_path}")
            subprocess.run(
                [
                    sys.executable,
                    EXCEL_SHRINK_SCRIPT,
                    file_path,
                    shrink_dir,
                    "--only-clean-workbook",
                ],
                check=True,
            )

            cleaned_path = os.path.join(shrink_dir, os.path.basename(file_path))
            if not os.path.exists(cleaned_path):
                logging.error(f"Cleaned workbook not found at {cleaned_path}")
                return None, None, None

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False

            excel.Application.DisplayAlerts = False

            abs_file = os.path.abspath(cleaned_path)

            file_exists = os.path.exists(abs_file)
            if not file_exists:
                logging.error(f"File does not exist: {abs_file}")
                return None, None, None

            try:
                wb = excel.Workbooks.Open(abs_file)
            except Exception as e:
                logging.error(f"Error opening workbook: {e}")
                return None, None, None

            if wb is None:
                logging.error(f"Failed to open workbook: {abs_file}")
                return None, None, None

        def _save():
            nonlocal excel, wb
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp_path = tmp.name
            wb.SaveAs(os.path.abspath(tmp_path))
            wb.Close()
            excel.Quit()
            os.remove(tmp_path)

    return _bench_template("Excel COM", _open, _save)


def _run_r_script(r_code: str, args: list):
    if shutil.which("Rscript") is None:
        logging.warning("Rscript not found – skipping R benchmark.")
        return None, None

    with tempfile.NamedTemporaryFile(
        delete=False, suffix=".R", mode="w", encoding="utf-8"
    ) as rf:
        rf.write(textwrap.dedent(r_code))
        r_path = rf.name

    try:
        completed = subprocess.run(
            ["Rscript", r_path, *args], capture_output=True, text=True, check=True
        )

        out = completed.stdout.strip()
        open_s, save_s = map(float, out.split(","))

        return open_s, save_s

    except subprocess.CalledProcessError as cpe:
        logging.error(
            f"Rscript failed (exit code {cpe.returncode}); " f"stderr: {cpe.stderr!r}"
        )
        raise cpe

    except Exception as e:
        logging.error(f"R benchmark error: {e}")
        raise e

    finally:
        try:
            os.remove(r_path)
        except Exception as e:
            raise e


def benchmark_r_openxlsx(file_path):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        out_path = tmp.name
    r_code = """
    args <- commandArgs(trailingOnly=TRUE)
    infile  <- args[1]
    outfile <- args[2]

    library(openxlsx)

    t_start_open <- proc.time()
    wb <- tryCatch({
        loadWorkbook(infile)
    }, error = function(e) {
        cat("ERROR_OPEN:", e$message, sep=""); quit(status=1)
    })
    t_finish_open <- proc.time() - t_start_open

    open_time <- t_finish_open[["elapsed"]][[1]]

    t_start_save <- proc.time()
    tryCatch({
        saveWorkbook(wb, outfile, overwrite = TRUE)
    }, error = function(e) {
        cat("ERROR_SAVE:", e$message, sep=""); quit(status=1)
    })
    t_finish_save <- proc.time() - t_start_save
    save_time <- t_finish_save[["elapsed"]][[1]]

    cat(open_time, save_time, sep = ",")
    """
    open_t, save_t = _run_r_script(r_code, [file_path, out_path])
    if open_t is None:
        os.remove(out_path)
        return None, None
    os.remove(out_path)
    return open_t, save_t


def benchmark_r_readxl_writexl(file_path):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        out_path = tmp.name
    r_script = f"""
    args <- commandArgs(trailingOnly = TRUE)
    file_in <- args[1]
    file_out <- args[2]

    library(readxl)
    library(writexl)

    sheets <- excel_sheets(file_in)

    t1 <- proc.time()
    data_list <- lapply(sheets, function(sht) read_excel(file_in, sheet = sht))
    t2 <- proc.time() - t1
    open_time <- t2["elapsed"][[1]]

    names(data_list) <- sheets

    t3 <- proc.time()
    write_xlsx(data_list, path=file_out)
    t4 <- proc.time() - t3
    save_time <- t4["elapsed"][[1]]

    cat(open_time, save_time, sep=",")
    """
    open_t, save_t = _run_r_script(r_script, [file_path, out_path])
    if open_t is None:
        os.remove(out_path)
        return None, None
    os.remove(out_path)
    return open_t, save_t


def run_and_measure(cmd, *, capture_output=False, text=True, poll_interval=0.1):
    if capture_output:
        proc = psutil.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=text
        )
    else:
        proc = psutil.Popen(cmd)

    peak_private = 0
    t0 = time.perf_counter()

    def _safe_children(p):
        try:
            return p.children(recursive=True)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return []

    def _safe_private_bytes(p):
        """Return private bytes for a psutil process, or 0 if not available."""
        try:
            m = p.memory_full_info()
            return int(getattr(m, "private", 0) or 0)
        except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
            return 0

    try:
        # Poll until process terminates
        while True:
            # update peak (parent + children)
            total_private = _safe_private_bytes(proc)
            for child in _safe_children(proc):
                total_private += _safe_private_bytes(child)

            if total_private > peak_private:
                peak_private = total_private

            # has the process ended?
            try:
                rc = proc.poll()
            except (psutil.NoSuchProcess, OSError):
                # process handle invalid now; treat as ended
                rc = 0
                break

            if rc is not None:
                break

            time.sleep(poll_interval)

        # Make sure we reap the process and collect output if requested
        stdout = stderr = None
        if capture_output:
            try:
                stdout, stderr = proc.communicate(timeout=5)
            except Exception:
                # If communicate fails (rare), fall back
                stdout, stderr = None, None
        else:
            try:
                proc.wait(timeout=5)
            except Exception:
                pass

        # One last peak update after exit (best-effort)
        total_private = _safe_private_bytes(proc)
        for child in _safe_children(proc):
            total_private += _safe_private_bytes(child)
        if total_private > peak_private:
            peak_private = total_private

        duration = time.perf_counter() - t0

        # Returncode
        try:
            returncode = proc.returncode
            if returncode is None:
                returncode = proc.wait(timeout=1)
        except (psutil.NoSuchProcess, OSError):
            # already gone
            returncode = 0

        return stdout, stderr, duration, peak_private, returncode

    except Exception:
        # If *we* crash while monitoring, do not crash again on kill().
        try:
            if proc.is_running():
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
            pass
        raise


EXCEL_SHRINK_SCRIPT = os.path.join("scripts", "excel_shrink", "excel_shrink.py")

# If your old analyser is elsewhere, adjust this path.
OLD_ANALYSER_SCRIPT = os.path.join(
    "scripts", "excel_shrink_analyzer", "excel_shrink_analyzer.py"
)

EXCEL_SHRINK_SINGLE_CORE_ARGS = ["--disable-multiprocessing"]
EXCEL_SHRINK_NO_SPLIT_ARGS = ["--disable-sheetdata-splitting"]


def run_excel_shrink_variant(
    original_file: str,
    output_dir: str,
    *,
    variant_name: str,
    extra_args: list[str] | None = None,
    chunk_size: int = 2097152,
):
    """
    Runs excel_shrink.py with optional extra args.
    Returns (duration_s, output_file_path, peak_mem_bytes, returncode).
    """
    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, os.path.basename(original_file))

    cmd = [
        sys.executable,
        EXCEL_SHRINK_SCRIPT,
        original_file,
        output_dir,
        "--chunk-size",
        str(chunk_size),
        "--force-overwrite",
    ]
    if extra_args:
        # put extra args at the end (safe for most argparse setups)
        cmd.extend(extra_args)

    logging.info("Running Excel Shrink [%s] on %s", variant_name, original_file)
    _, _, duration, peak_mem_bytes, returncode = run_and_measure(
        cmd, capture_output=False
    )

    if returncode == 0:
        logging.info(
            "Excel Shrink [%s] completed in %.2fs; peak private %.2f MiB",
            variant_name,
            duration,
            peak_mem_bytes / (1024**2),
        )
    else:
        logging.error(
            "Excel Shrink [%s] failed (exit %s) for %s",
            variant_name,
            returncode,
            original_file,
        )

    return duration, output_file, peak_mem_bytes, returncode


def run_old_excel_shrink_analyser(
    original_file: str,
    output_dir: str,
    *,
    log_csv_path: str | None = None,
):
    """
    Runs the old excel_shrink_analyser.py on one XLSX and measures time + peak memory.
    Produces an output XLSX (same basename in output_dir).
    Returns (duration_s, output_file_path, peak_mem_bytes, returncode, stdout, stderr).
    """
    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, os.path.basename(original_file))

    cmd = [
        sys.executable,
        OLD_ANALYSER_SCRIPT,
        original_file,
        output_file,
    ]

    if log_csv_path:
        cmd.extend(["--log", log_csv_path])

    logging.info("Running OLD analyser on %s", original_file)

    stdout, stderr, duration, peak_mem_bytes, returncode = run_and_measure(
        cmd, capture_output=True, text=True
    )

    if returncode == 0:
        logging.info(
            "OLD analyser completed in %.2fs; peak private %.2f MiB",
            duration,
            peak_mem_bytes / (1024**2),
        )
    else:
        logging.error(
            "OLD analyser failed (exit %s) for %s; stderr=%r",
            returncode,
            original_file,
            stderr,
        )

    return duration, output_file, peak_mem_bytes, returncode, stdout, stderr


def run_excel_shrink(original_file, output_dir):
    # Default run = no extra args
    duration, output_file, peak_mem_bytes, returncode = run_excel_shrink_variant(
        original_file,
        output_dir,
        variant_name="default",
        extra_args=None,
        chunk_size=2097152,
    )

    if returncode != 0:
        return duration, output_file, peak_mem_bytes

    return duration, output_file, peak_mem_bytes


def controller_main(args):
    """
    In 'controller' mode, we loop over all files and libraries, call measure-one
    as a separate subprocess for each scenario, then run shrink, measure shrunk file,
    and finally write CSV and generate the chart.
    """

    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.abspath(os.path.join(script_dir, args.input_folder))
    shrunk_folder = os.path.abspath(os.path.join(script_dir, args.shrunk_folder))
    second_shrink_shrunk_folder = os.path.abspath(
        os.path.join(script_dir, args.second_shrinkage_shrunk_folder)
    )
    single_core_shrunk_folder = os.path.abspath(
        os.path.join(script_dir, args.shrunk_folder + "_single_core")
    )
    no_split_shrunk_folder = os.path.abspath(
        os.path.join(script_dir, args.shrunk_folder + "_no_split")
    )
    # analyser_shrunk_folder = os.path.abspath(
    #    os.path.join(script_dir, args.shrunk_folder + "_old_analyser")
    # )

    csv_out = os.path.abspath(os.path.join(script_dir, args.csv_out))

    files_to_download = [
        {
            "url": "https://www.destatis.de/DE/Themen/Staat/Oeffentliche-Finanzen/Ausgaben-Einnahmen/Publikationen/Downloads-Ausgaben-und-Einnahmen/statistischer-bericht-rechnungsergebnis-kernhaushalt-gemeinden-2140331217005.xlsx?__blob=publicationFile&v=4",
            "filename": "statistischer-bericht-kernhaushalt-gemeinden.xlsx",
        },
        {
            "url": "https://pasteur.epa.gov/uploads/10.23719/1503098/QT_RR%20data_dobutamine%20challenge%20test_Hazari_Dec%202015.xlsx",
            "filename": "QT_RR data_dobutamine challenge test_Hazari_Dec 2015.xlsx",
        },
    ]

    output_folder = "scripts/loading_times_size_reduction/input"

    os.environ["SSL_CERT_FILE"] = certifi.where()

    os.makedirs(output_folder, exist_ok=True)

    for file in files_to_download:
        output_path = os.path.join(output_folder, file["filename"])
        if os.path.exists(output_path):
            print(f"✅ File already exists: {output_path}")
        else:
            print(f"⬇️ Downloading {file['filename']}...")
            try:
                urllib.request.urlretrieve(file["url"], output_path)
                print(f"✅ Downloaded: {output_path}")
            except Exception as e:
                print(f"❌ Failed to download {file['filename']}: {e}")

    logging.info("Starting measurements...")

    for p in [
        shrunk_folder,
        single_core_shrunk_folder,
        no_split_shrunk_folder,
        # analyser_shrunk_folder,
    ]:
        if not os.path.exists(p):
            os.makedirs(p)
            logging.info("Created folder: %s", p)

    all_files = [
        f
        for f in os.listdir(input_folder)
        if f.lower().endswith(".xlsx") and os.path.isfile(os.path.join(input_folder, f))
    ]
    logging.info(f"Found {len(all_files)} Excel files in {input_folder}.")

    library_names = [
        "openpyxl(default)",
        "pandas",
        "Microsoft Excel",
        "R openxlsx",
        "R readxl+writexl",
    ]

    rows = []

    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".xlsx")]
    for fname in excel_files:
        logging.info(f"Processing file: {fname}")

        original_path = os.path.join(input_folder, fname)
        orig_size = os.path.getsize(original_path)

        logging.info(f"Running shrink on: {fname}")
        shrink_time, shrunk_path, shrink_peak_mem = run_excel_shrink(
            original_path, shrunk_folder
        )
        logging.info(
            f"Shrink results for {fname}: Time={shrink_time}, Shrunk Path={shrunk_path}, Peak Memory={shrink_peak_mem}"
        )

        # Excel Shrink single-core
        sc_time, sc_path, sc_peak_mem, sc_rc = run_excel_shrink_variant(
            original_path,
            single_core_shrunk_folder,
            variant_name="single-core",
            extra_args=EXCEL_SHRINK_SINGLE_CORE_ARGS,
        )
        sc_size = (
            os.path.getsize(sc_path)
            if (sc_rc == 0 and os.path.exists(sc_path))
            else None
        )

        rows.append(
            {
                "File": fname,
                "Library": "Excel Shrink (single-core)",
                "Original Open Time (s)": None,
                "Original Save Time (s)": None,
                "Original Peak Memory": None,
                "Shrink Time (s)": sc_time,
                "Shrinked Open Time (s)": None,
                "Shrinked Save Time (s)": None,
                "Shrinked Peak Memory": sc_peak_mem,
                "Original Size (bytes)": orig_size,
                "Shrinked Size (bytes)": sc_size,
            }
        )

        # Excel Shrink with "no split sheetData" mode
        ns_time, ns_path, ns_peak_mem, ns_rc = run_excel_shrink_variant(
            original_path,
            no_split_shrunk_folder,
            variant_name="no-split",
            extra_args=EXCEL_SHRINK_NO_SPLIT_ARGS,
        )
        ns_size = (
            os.path.getsize(ns_path)
            if (ns_rc == 0 and os.path.exists(ns_path))
            else None
        )

        rows.append(
            {
                "File": fname,
                "Library": "Excel Shrink (no-split-sheetdata)",
                "Original Open Time (s)": None,
                "Original Save Time (s)": None,
                "Original Peak Memory": None,
                "Shrink Time (s)": ns_time,
                "Shrinked Open Time (s)": None,
                "Shrinked Save Time (s)": None,
                "Shrinked Peak Memory": ns_peak_mem,
                "Original Size (bytes)": orig_size,
                "Shrinked Size (bytes)": ns_size,
            }
        )

        # 3) OLD analyser run
        # Optional: per-file analyser log
        analyser_log = None
        # analyser_log = os.path.join(analyser_shrunk_folder, f"{os.path.splitext(fname)[0]}_analyser_log.csv")

        # an_time, an_path, an_peak_mem, an_rc, an_stdout, an_stderr = (
        #    run_old_excel_shrink_analyser(
        #        original_path,
        #        analyser_shrunk_folder,
        #        log_csv_path=analyser_log,
        #    )
        # )
        #
        # an_size = (
        #    os.path.getsize(an_path)
        #    if (an_rc == 0 and os.path.exists(an_path))
        #    else None
        # )
        #
        # rows.append(
        #    {
        #        "File": fname,
        #        "Library": "Excel Shrink Analyser (old)",
        #        "Original Open Time (s)": None,
        #        "Original Save Time (s)": None,
        #        "Original Peak Memory": None,
        #        "Shrink Time (s)": an_time,  # runtime of analyser
        #        "Shrinked Open Time (s)": None,
        #        "Shrinked Save Time (s)": None,
        #        "Shrinked Peak Memory": an_peak_mem,  # peak private mem during analyser
        #        "Original Size (bytes)": orig_size,
        #        "Shrinked Size (bytes)": an_size,  # size of analyser-produced XLSX
        #    }
        # )

        benchmark_results = {}
        for lib in library_names:
            logging.info(f"Measuring library: {lib} on file: {fname}")
            benchmark_results[lib] = call_measure_one(lib, original_path)
            logging.info(
                f"Results for {lib} on {fname}: Open Time={benchmark_results[lib][0]}, Save Time={benchmark_results[lib][1]}, Peak Memory={benchmark_results[lib][2]}"
            )

        shrunk_results = {}
        if shrunk_path and os.path.exists(shrunk_path):
            shrunk_size = os.path.getsize(shrunk_path)
            for lib in library_names:
                logging.info(f"Measuring library: {lib} on file: {shrunk_path}")
                shrunk_results[lib] = call_measure_one(lib, shrunk_path)
                logging.info(
                    f"Results for {lib} on shrunk {fname}: Open Time={shrunk_results[lib][0]}, Save Time={shrunk_results[lib][1]}, Peak Memory={shrunk_results[lib][2]}"
                )
        else:
            for lib in library_names:
                shrunk_results[lib] = (None, None, None)

        for lib in library_names:
            orig_open, orig_save, orig_mem = benchmark_results[lib]
            shr_open, shr_save, shr_mem = shrunk_results[lib]
            row = {
                "File": fname,
                "Library": lib,
                "Original Open Time (s)": orig_open,
                "Original Save Time (s)": orig_save,
                "Original Peak Memory": orig_mem,
                "Shrink Time (s)": shrink_time,
                "Shrinked Open Time (s)": shr_open,
                "Shrinked Save Time (s)": shr_save,
                "Shrinked Peak Memory": shr_mem,
                "Original Size (bytes)": orig_size,
                "Shrinked Size (bytes)": shrunk_size,
            }
            rows.append(row)

        if not os.path.exists(second_shrink_shrunk_folder):
            os.makedirs(second_shrink_shrunk_folder)
            logging.info(f"Created second shrink folder: {second_shrink_shrunk_folder}")

        if shrunk_path and os.path.exists(shrunk_path):
            logging.info(f"Running second shrink on: {shrunk_path}")
            shrunk_shrink_time, shrunk_shrink_path, shrunk_shrink_peak_mem = (
                run_excel_shrink(shrunk_path, second_shrink_shrunk_folder)
            )
            logging.info(
                f"Second shrink results for {shrunk_path}: Time={shrunk_shrink_time}, Shrunk Path={shrunk_shrink_path}, Peak Memory={shrunk_shrink_peak_mem}"
            )
            if shrunk_shrink_path and os.path.exists(shrunk_shrink_path):
                shrink_shrunk_size = os.path.getsize(shrunk_shrink_path)
            else:
                logging.error(
                    f"Shrunk file {shrunk_shrink_path} not found after second shrink."
                )
                shrink_shrunk_size = 0

        row = {
            "File": fname,
            "Library": "Excel Shrink",
            "Original Open Time (s)": 0,
            "Original Save Time (s)": shrink_time,
            "Original Peak Memory": shrink_peak_mem,
            "Shrink Time (s)": shrink_time,
            "Shrinked Open Time (s)": 0,
            "Shrinked Save Time (s)": shrunk_shrink_time,
            "Shrinked Peak Memory": shrunk_shrink_peak_mem,
            "Original Size (bytes)": shrunk_size,
            "Shrinked Size (bytes)": shrink_shrunk_size,
        }
        rows.append(row)

    fieldnames = list(rows[0].keys()) if rows else []
    with open(csv_out, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
    logging.info(f"Measurements saved to {csv_out}")

    generate_chart(csv_out)


def call_measure_one(library_name, file_path):
    """
    Calls this script in 'measure-one' mode for the given library and file.
    Returns (open_time, save_time) or (None, None) on error.
    """

    cmd = [
        sys.executable,
        os.path.abspath(__file__),
        "measure-one",
        "--library",
        library_name,
        "--file",
        file_path,
    ]

    stdout, stderr, duration, peak_mem_bytes, returncode = run_and_measure(
        cmd, capture_output=True, text=True
    )

    logging.info(
        f"{library_name} completed in {duration:.2f}s; Peak memory usage {peak_mem_bytes / (1024**2):.2f} MB; file: {file_path}"
    )

    if returncode == 0:
        try:
            data = json.loads(stdout.strip())
            return (
                data.get("Original Open Time (s)"),
                data.get("Original Save Time (s)"),
                peak_mem_bytes,
            )
        except json.JSONDecodeError:
            logging.error(f"JSON parse error from measure-one output: {stdout!r}")
            return (None, None, None)
    else:
        logging.error(
            f"measure-one ({library_name}) failed with exit code {returncode}; stderr: {stderr!r}"
        )
        return (None, None, None)


def generate_chart(csv_file):
    """
    Generate a chart from the CSV file.
    """

    logging.info(f"Generating chart from {csv_file}")
    if not os.path.exists(csv_file):
        logging.warning(f"CSV file {csv_file} not found; skipping chart.")
        return

    try:
        df = pd.read_csv(csv_file)
        _generate_complex_chart(df)
    except Exception as e:
        logging.error(f"Could not generate chart: {e}")


def _generate_complex_chart(df):
    logging.info("Using extended chart routine (with memory at bottom)...")

    if df.empty:
        logging.warning("DataFrame is empty; nothing to plot.")
        return

    needed_cols = {
        "File",
        "Library",
        "Original Open Time (s)",
        "Original Save Time (s)",
        "Shrink Time (s)",
        "Shrinked Open Time (s)",
        "Shrinked Save Time (s)",
        "Original Size (bytes)",
        "Shrinked Size (bytes)",
        "Original Peak Memory",
        "Shrinked Peak Memory",
    }
    missing = needed_cols - set(df.columns)
    if missing:
        logging.warning(f"Missing columns: {missing}. Cannot plot.")
        return

    # --- layout constants (same as before) ---
    PX_PER_CHAR = 6.0
    INSIDE_PAD_PX = 0.0
    RIGHT_LABEL_PAD_PX = 12.0
    RIGHT_ARROW_SHRINK = 1
    FONT_SIZE = 7
    VERTICAL_LABEL_OFFSET_PX = -0.6

    def can_fit_inside(ax, width_data, text):
        x_min, x_max = ax.get_xlim()
        data_per_px = (x_max - x_min) / max(ax.bbox.width, 1.0)
        needed_data = ((len(text) * PX_PER_CHAR) + INSIDE_PAD_PX) * data_per_px
        return width_data >= needed_data

    def right_label(ax, x_end, y, parts):
        x_min, x_max = ax.get_xlim()
        data_per_px = (x_max - x_min) / max(ax.bbox.width, 1.0)
        x_text = x_end + RIGHT_LABEL_PAD_PX * data_per_px
        text = ", ".join(parts)
        ax.annotate(
            text,
            xy=(x_end, y + 0.001),
            xytext=(x_text, y + 0.008),
            va="center",
            ha="left",
            fontsize=FONT_SIZE,
            arrowprops=dict(arrowstyle="->", shrinkA=RIGHT_ARROW_SHRINK, lw=0.8),
        )

    def text_vshift(ax, points: float):
        return mtransforms.ScaledTranslation(
            0, points / 72.0, ax.figure.dpi_scale_trans
        )

    library_info = [
        ("openpyxl(default)", "darkorange", "orangered"),
        ("pandas", "steelblue", "royalblue"),
        ("Microsoft Excel", "darkgreen", "limegreen"),
        ("R openxlsx", "red", "firebrick"),
        ("R readxl+writexl", "chocolate", "sienna"),
        ("Excel Shrink", "mediumpurple", "purple"),
    ]
    n_libs = len(library_info)

    extra_run_info = [
        ("Excel Shrink (single-core)", "slateblue"),
        ("Excel Shrink (no-split-sheetdata)", "teal"),
        # ("Excel Shrink Analyser (old)", "dimgray"),
    ]
    n_extra = len(extra_run_info)

    files = sorted(df["File"].unique())
    file_labels = {f: f"File {i+1}" for i, f in enumerate(files)}

    n_files = len(files)
    if n_files == 0:
        logging.info("No files found in data; nothing to plot.")
        return

    temp = df.set_index(["File", "Library"], drop=False)

    def get_times(f, lib, col):
        if (f, lib) in temp.index and col in temp.columns:
            val = temp.loc[(f, lib), col]
            if isinstance(val, pd.Series):
                val = val.iloc[0]
            try:
                return float(val) if pd.notna(val) else 0.0
            except Exception:
                return 0.0
        return 0.0

    def shortfile(fn):
        return fn if len(fn) <= 15 else fn[:5] + "..." + fn[-5:]

    # ==== Figure layout ====
    fig = plt.figure(figsize=(14, 7 + 1.8 * n_files))

    # Outer grid: keep a small gap between top block and memory block
    gs = GridSpec(nrows=2, ncols=1, height_ratios=[7, 1.8 * n_files], hspace=0.20)

    # Top (time + size)
    gs_top = GridSpecFromSubplotSpec(
        nrows=2, ncols=1, subplot_spec=gs[0], height_ratios=[4, 1], hspace=0.35
    )
    ax_time = fig.add_subplot(gs_top[0, 0])
    ax_size = fig.add_subplot(gs_top[1, 0])

    # Bottom (memory panes)
    gs_bottom = GridSpecFromSubplotSpec(
        nrows=n_files, ncols=1, subplot_spec=gs[1], hspace=0.08  # e.g., 0.05–0.10
    )
    mem_axes = [fig.add_subplot(gs_bottom[i, 0]) for i in range(n_files)]

    mem_axes[0].set_title("Memory Consumption")
    plt.subplots_adjust(left=0.18)  # tweak to taste (0.16–0.22)

    # ==== TIME section (top) ====
    total_bars_per_file = (n_libs - 1) + n_libs + n_extra
    bar_height = 0.09
    offsets_all = np.linspace(-0.60, 0.60, total_bars_per_file)
    offsets_orig = offsets_all[: (n_libs - 1)]
    offsets_shr = offsets_all[(n_libs - 1) : (n_libs - 1 + n_libs)]
    offsets_extra = offsets_all[(n_libs - 1 + n_libs) :]

    y_positions_top = np.arange(n_files) * (bar_height * (total_bars_per_file + 2))
    for i, f in enumerate(files):
        base = y_positions_top[i]
        shrink_t = get_times(f, "Excel Shrink", "Shrink Time (s)")

        libraries_no_shrink = [
            (n, co, cs) for (n, co, cs) in library_info if n != "Excel Shrink"
        ]
        for lib_index, (lib_name, c_open, c_save) in enumerate(libraries_no_shrink):
            pos = base + offsets_orig[lib_index]
            open_t = get_times(f, lib_name, "Original Open Time (s)")
            save_t = get_times(f, lib_name, "Original Save Time (s)")

            ax_time.barh(
                pos,
                open_t,
                height=bar_height,
                color=c_open,
                label=f"{lib_name} open" if i == 0 else "",
            )
            ax_time.barh(
                pos,
                save_t,
                height=bar_height,
                left=open_t,
                color=c_save,
                label=f"{lib_name} save" if i == 0 else "",
            )

            label_open = f"{open_t:.1f}"
            label_save = f"{save_t:.1f}"
            ok_open = open_t > 0 and can_fit_inside(ax_time, open_t, label_open)
            ok_save = save_t > 0 and can_fit_inside(ax_time, save_t, label_save)

            if ok_open:
                ax_time.text(
                    open_t / 2.0,
                    pos,
                    label_open,
                    va="center",
                    ha="center",
                    fontsize=FONT_SIZE,
                    color="white",
                    transform=ax_time.transData
                    + text_vshift(ax_time, VERTICAL_LABEL_OFFSET_PX),
                )
            if ok_save:
                ax_time.text(
                    open_t + (save_t / 2.0),
                    pos,
                    label_save,
                    va="center",
                    ha="center",
                    fontsize=FONT_SIZE,
                    color="white",
                    transform=ax_time.transData
                    + text_vshift(ax_time, VERTICAL_LABEL_OFFSET_PX),
                )
            if not (ok_open and ok_save):
                x_end = open_t + save_t
                parts = []
                if open_t > 0:
                    parts.append(label_open)
                if save_t > 0:
                    parts.append(label_save)
                if parts:
                    right_label(ax_time, x_end, pos, parts)

        for lib_index, (lib_name, c_open, c_save) in enumerate(library_info):
            pos = base + offsets_shr[lib_index]
            shr_open = get_times(f, lib_name, "Shrinked Open Time (s)")
            shr_save = get_times(f, lib_name, "Shrinked Save Time (s)")

            ax_time.barh(
                pos,
                shrink_t,
                height=bar_height,
                color="mediumorchid",
                label="Shrink time" if (i == 0 and lib_index == 0) else "",
            )
            ax_time.barh(
                pos,
                shr_open,
                height=bar_height,
                left=shrink_t,
                color=c_open,
                label=f"{lib_name} shrunk open" if i == 0 else "",
            )
            ax_time.barh(
                pos,
                shr_save,
                height=bar_height,
                left=shrink_t + shr_open,
                color=c_save,
                label=f"{lib_name} shrunk save" if i == 0 else "",
            )

            lab0, lab1, lab2 = f"{shrink_t:.1f}", f"{shr_open:.1f}", f"{shr_save:.1f}"
            ok0 = shrink_t > 0 and can_fit_inside(ax_time, shrink_t, lab0)
            ok1 = shr_open > 0 and can_fit_inside(ax_time, shr_open, lab1)
            ok2 = shr_save > 0 and can_fit_inside(ax_time, shr_save, lab2)

            if ok0:
                ax_time.text(
                    shrink_t / 2.0,
                    pos,
                    lab0,
                    va="center",
                    ha="center",
                    fontsize=FONT_SIZE,
                    color="white",
                    transform=ax_time.transData
                    + text_vshift(ax_time, VERTICAL_LABEL_OFFSET_PX),
                )
            if ok1:
                ax_time.text(
                    shrink_t + (shr_open / 2.0),
                    pos,
                    lab1,
                    va="center",
                    ha="center",
                    fontsize=FONT_SIZE,
                    color="white",
                    transform=ax_time.transData
                    + text_vshift(ax_time, VERTICAL_LABEL_OFFSET_PX),
                )
            if ok2:
                ax_time.text(
                    shrink_t + shr_open + (shr_save / 2.0),
                    pos,
                    lab2,
                    va="center",
                    ha="center",
                    fontsize=FONT_SIZE,
                    color="white",
                    transform=ax_time.transData
                    + text_vshift(ax_time, VERTICAL_LABEL_OFFSET_PX),
                )
            if not (ok0 and ok1 and ok2):
                x_end = shrink_t + shr_open + shr_save
                parts = []
                if shrink_t > 0:
                    parts.append(lab0)
                if shr_open > 0:
                    parts.append(lab1)
                if shr_save > 0:
                    parts.append(lab2)
                if parts:
                    right_label(ax_time, x_end, pos, parts)

        # --- Extra run-only bars (single segment each) ---
        for ex_i, (run_name, run_color) in enumerate(extra_run_info):
            pos = base + offsets_extra[ex_i]
            run_t = get_times(f, run_name, "Shrink Time (s)")

            ax_time.barh(
                pos,
                run_t,
                height=bar_height,
                color=run_color,
                label=run_name if (i == 0) else "",
            )

            lab = f"{run_t:.1f}"
            ok = run_t > 0 and can_fit_inside(ax_time, run_t, lab)
            if ok:
                ax_time.text(
                    run_t / 2.0,
                    pos,
                    lab,
                    va="center",
                    ha="center",
                    fontsize=FONT_SIZE,
                    color="white",
                    transform=ax_time.transData
                    + text_vshift(ax_time, VERTICAL_LABEL_OFFSET_PX),
                )
            elif run_t > 0:
                right_label(ax_time, run_t, pos, [lab])

    ax_time.margins(x=0.02)
    for ax in mem_axes:
        ax.margins(y=0.08)

    ax_time.set_yticks(y_positions_top)
    ax_time.set_yticklabels([file_labels[f] for f in files])
    ax_time.invert_yaxis()
    ax_time.set_xlabel("Time (seconds)")
    ax_time.set_title("Time Comparison per Excel File")
    # Collect all handles/labels that were added while drawing
    handles, labels = ax_time.get_legend_handles_labels()

    # Drop empties and keep first occurrence of each label
    pairs = [(h, l) for h, l in zip(handles, labels) if l]
    # Build an OrderedDict keyed by label; first occurrence wins
    seen = OrderedDict()
    for h, l in pairs:
        # rename "Excel Shrink shrunk save" to "Second Excel Shrink run" for legend brevity
        if "Excel Shrink" in l and "shrunk" in l:
            l = l.replace("Excel Shrink shrunk save", "Second Excel Shrink run")
        # drop if contains "shrunk" to reduce legend size
        if "shrunk" in l:
            continue
        if l not in seen:
            seen[l] = h

    # Create the legend with unique entries
    ax_time.legend(list(seen.values()), list(seen.keys()), ncol=3, fontsize=8)
    # ==== SIZE section (middle) ====

    def _bar_label(ax, x, y, text, fontsize):
        ax.text(
            x,
            y,
            text,
            va="center",
            ha="center",
            fontsize=fontsize,
            color="white",
        )

    group_size = df.groupby("File")[
        ["Original Size (bytes)", "Shrinked Size (bytes)"]
    ].max()
    orig_size_mb_list, shrunk_size_mb_list = [], []
    orig_size_b_list, shrunk_size_b_list = [], []
    for f in files:
        if f in group_size.index:
            row = group_size.loc[f]
            orig_size = float(row.get("Original Size (bytes)", 0) or 0)  # bytes
            shrunk_size = float(row.get("Shrinked Size (bytes)", 0) or 0)  # bytes
        else:
            orig_size = 0.0
            shrunk_size = 0.0

        orig_size_b_list.append(orig_size)
        shrunk_size_b_list.append(shrunk_size)
        orig_size_mb_list.append(orig_size / (1024 * 1024))
        shrunk_size_mb_list.append(shrunk_size / (1024 * 1024))

    y_positions_bottom = np.arange(n_files) * 0.1
    file_bar_height = 0.04
    for i, f in enumerate(files):
        base = y_positions_bottom[i]
        # draw bars
        b1 = ax_size.barh(
            base - file_bar_height * 0.55,
            orig_size_mb_list[i],
            height=file_bar_height,
            color="darkorange",
            label="Original Size (MB)" if i == 0 else "",
        )
        b2 = ax_size.barh(
            base + file_bar_height * 0.55,
            shrunk_size_mb_list[i],
            height=file_bar_height,
            color="steelblue",
            label="Shrinked Size (MB)" if i == 0 else "",
        )

        # add value labels (MB)
        _bar_label(
            ax_size,
            orig_size_mb_list[i] * 0.97,
            base - file_bar_height * 0.55,
            f"{int(round(orig_size_mb_list[i]))} MB",
            FONT_SIZE,
        )
        # _bar_label(
        #    ax_size,
        #    shrunk_size_mb_list[i] - 1.5,
        #    base + file_bar_height * 0.55,
        #    f"{int(round(shrunk_size_mb_list[i]))} MB",
        #    FONT_SIZE,
        # )

    for i, f in enumerate(files):
        base = y_positions_bottom[i]
        orig_b = orig_size_b_list[i]
        shr_b = shrunk_size_b_list[i]
        shr_mb = shrunk_size_mb_list[i]

        if orig_b > 0:
            red_pct = 100.0 * (1.0 - (shr_b / orig_b))
            label = f"{int(round(shrunk_size_mb_list[i]))} MB (−{red_pct:.1f}%)"

            x_min, x_max = ax_size.get_xlim()
            data_per_px = (x_max - x_min) / max(ax_size.bbox.width, 1.0)
            x_text = shr_mb + 10 * data_per_px
            x_text = max(
                shr_mb + 10 * data_per_px, ax_size.get_xlim()[0] + 5 * data_per_px
            )
            ax_size.annotate(
                label,
                xy=(shr_mb + 0.15, base + file_bar_height * 0.55),
                xytext=(x_text, base + file_bar_height * 0.55),
                va="center",
                ha="left",
                fontsize=FONT_SIZE,
                arrowprops=dict(arrowstyle="->", lw=0.8, shrinkA=0, shrinkB=0),
            )

    ax_size.set_yticks(y_positions_bottom)
    ax_size.set_yticklabels([file_labels[f] for f in files])
    ax_size.invert_yaxis()
    ax_size.set_xlabel("File Size (MB)")
    ax_size.set_title("File Size Comparison per Excel File")
    ax_size.legend(ncol=2, fontsize=8)

    # ==== MEMORY section (bottom) ====
    # Prepare numbers (MiB)
    df["Original Peak Memory"] = pd.to_numeric(
        df["Original Peak Memory"], errors="coerce"
    ).fillna(0.0)
    df["Shrinked Peak Memory"] = pd.to_numeric(
        df["Shrinked Peak Memory"], errors="coerce"
    ).fillna(0.0)

    lib_order = [n for (n, _, _) in library_info] + [n for (n, _) in extra_run_info]
    present = set(df["Library"].dropna().unique())

    # keep order but remove duplicates
    libraries = list(OrderedDict.fromkeys([l for l in lib_order if l in present]))
    extra_run_names = {n for (n, _) in extra_run_info}

    plt.rcParams.update({"xtick.labelsize": 8})

    for ax, file_name in zip(mem_axes, files):
        subset = df[df["File"] == file_name]

        # one row per library; peak memory should be the max across repeats
        subset_by_lib = subset.groupby("Library", as_index=True)[
            ["Original Peak Memory", "Shrinked Peak Memory"]
        ].max()

        x = np.arange(len(libraries))
        bar_width = 0.25

        original_mem = []
        shrunk_mem = []
        run_mem = []

        for lib in libraries:
            row = subset_by_lib.loc[lib] if lib in subset_by_lib.index else None

            if lib in extra_run_names:
                # extra runs: only show the "run" bar (stored in Shrinked Peak Memory)
                original_mem.append(0.0)
                shrunk_mem.append(0.0)
                run_mem.append(
                    (float(row["Shrinked Peak Memory"]) / (1024**2))
                    if row is not None
                    else 0.0
                )
            else:
                original_mem.append(
                    (float(row["Original Peak Memory"]) / (1024**2))
                    if row is not None
                    else 0.0
                )
                shrunk_mem.append(
                    (float(row["Shrinked Peak Memory"]) / (1024**2))
                    if row is not None
                    else 0.0
                )
                run_mem.append(0.0)

        # sanity check: prevents shape mismatch forever
        assert len(x) == len(original_mem) == len(shrunk_mem) == len(run_mem)

        # Make run bars centered for extra runs
        run_mem_center = [
            run_mem[i] if libraries[i] in extra_run_names else 0.0
            for i in range(len(libraries))
        ]
        run_mem_right = [
            run_mem[i] if libraries[i] not in extra_run_names else 0.0
            for i in range(len(libraries))
        ]

        bars_orig = ax.bar(
            x - bar_width,
            original_mem,
            width=bar_width,
            label="Original Peak Memory (MiB)",
            color="steelblue",
            edgecolor="black",
        )

        bars_shr = ax.bar(
            x,
            shrunk_mem,
            width=bar_width,
            label="Shrinked Peak Memory (MiB)",
            color="darkorange",
            edgecolor="black",
        )

        # centered run bars for the extra runs
        bars_run_center = ax.bar(
            x,
            run_mem_center,
            width=bar_width,
            label="Run Peak Memory (MiB)",
            color="mediumpurple",
            edgecolor="black",
        )

        # (optional) keep the old right-offset run bars for non-extra (likely all zero)
        bars_run_right = ax.bar(
            x + bar_width,
            run_mem_right,
            width=bar_width,
            color="mediumpurple",
            edgecolor="black",
        )

        # --- Add numeric labels to memory bars ---
        def _data_per_px_y(ax):
            y_min, y_max = ax.get_ylim()
            return (y_max - y_min) / max(ax.bbox.height, 1.0)

        def _can_fit_inside_vertical(ax, bar_height_data, font_size_pts=7, fudge=1.1):
            text_px = (font_size_pts / 72.0) * ax.figure.dpi * fudge
            needed_data = text_px * _data_per_px_y(ax)
            return bar_height_data >= needed_data

        PAD_PX = 6  # gap above bar for outside labels

        for bar in (
            list(bars_orig)
            + list(bars_shr)
            + list(bars_run_center)
            + list(bars_run_right)
        ):
            h = bar.get_height()
            if h <= 0:
                continue

            label = f"{int(round(h))}"
            x_center = bar.get_x() + bar.get_width() / 2.0

            if _can_fit_inside_vertical(ax, h, font_size_pts=FONT_SIZE):
                ax.text(
                    x_center,
                    h / 2.0,
                    label,
                    ha="center",
                    va="center",
                    fontsize=FONT_SIZE,
                    color="white",
                )
            else:
                y_pad = PAD_PX * _data_per_px_y(ax)
                ax.annotate(
                    label,
                    xy=(x_center, h),
                    xytext=(x_center, h + y_pad),
                    ha="center",
                    va="bottom",
                    fontsize=FONT_SIZE,
                    arrowprops=dict(arrowstyle="->", lw=0.8, shrinkA=0, shrinkB=0),
                )

        ax.set_ylabel(f"Peak Memory (MiB)\n{file_labels[file_name]}", fontsize=9)
        ax.ticklabel_format(style="plain", axis="y")

        display_labels = [
            lib.replace(
                "Excel Shrink (single-core)", "Excel Shrink\n(single-core)"
            ).replace(
                "Excel Shrink (no-split-sheetdata)",
                "Excel Shrink\n(no-split-sheetdata)",
            )
            for lib in libraries
        ]
        ax.set_xticks(x)
        ax.set_xticklabels(display_labels, ha="center", fontsize=8)

        ax.legend(fontsize=8)
        ax.grid(axis="x", linestyle="--", alpha=0.5)

        # if ax is not mem_axes[-1]:
        #    ax.set_xticklabels([])
        #    ax.set_xlabel("")

    plt.tight_layout()
    # create directory if not exists
    os.makedirs(os.path.dirname(chart_output_path), exist_ok=True)

    plt.savefig(chart_output_path, dpi=300, bbox_inches="tight")
    logging.info(f"Chart generated and saved to {chart_output_path}")
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(
        description="Run measurements on Excel files and generate comparison chart."
    )

    subparsers = parser.add_subparsers(
        dest="mode", required=True, help="Subcommand: controller or measure-one"
    )

    # add run tests / controller mode
    ctrl_parser = subparsers.add_parser(
        "controller", help="Run measurements on all files and libraries."
    )

    ctrl_parser.add_argument(
        "--run-tests",
        "--run_tests",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Whether to run the measurements (default: True).",
    )

    meas_parser = subparsers.add_parser(
        "measure-one", help="Measure one library on one file."
    )
    meas_parser.add_argument(
        "--library", required=True, help="Library name, e.g. 'openpyxl(default)'"
    )
    meas_parser.add_argument("--file", required=True, help="Path to the .xlsx file")

    args, unknown = parser.parse_known_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(script_dir, "input")
    shrunk_folder = os.path.join(script_dir, "shrunk_files")
    second_shrink_shrunk_folder = os.path.join(
        script_dir, "second_shrinkage_shrunk_files"
    )
    csv_out = os.path.join(script_dir, "excel_benchmarks.csv")

    if not args.mode:
        args.mode = "controller"

    if args.mode == "measure-one":
        measure_one_main(args)
    else:
        args.input_folder = input_folder
        args.shrunk_folder = shrunk_folder
        args.second_shrinkage_shrunk_folder = second_shrink_shrunk_folder
        args.csv_out = csv_out

        if not args.run_tests and os.path.exists(args.csv_out):
            logging.info(
                f"CSV file {args.csv_out} already exists. Skipping measurements."
            )
            generate_chart(args.csv_out)
            sys.exit(0)
        controller_main(args)


if __name__ == "__main__":
    main()
