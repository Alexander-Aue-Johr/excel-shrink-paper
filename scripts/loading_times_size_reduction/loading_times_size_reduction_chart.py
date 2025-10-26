#!/usr/bin/env python3

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


plt.switch_backend("agg")

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
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

            shrink_script = os.path.join("..", "excel-shrink", "excel_shrink.py")
            logging.info(f"Running excel_shrink --only-clean-workbook on {file_path}")
            subprocess.run(
                [
                    sys.executable,
                    shrink_script,
                    "--only-clean-workbook",
                    file_path,
                    shrink_dir,
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
        return None, None

    except Exception as e:
        logging.error(f"R benchmark error: {e}")
        return None, None

    finally:
        try:
            os.remove(r_path)
            raise
        except OSError:
            raise


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


import psutil, subprocess, time


def run_and_measure(cmd, *, capture_output=False, text=True, poll_interval=0.1):
    if capture_output:
        proc = psutil.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=text
        )
    else:
        proc = psutil.Popen(cmd)

    peak_private = 0
    t0 = time.perf_counter()

    try:
        while proc.is_running():
            try:
                m = proc.memory_full_info()
                private_parent = getattr(m, "private", 0)  # Bytes, die committed sind

                total_private = private_parent
                for child in proc.children(recursive=True):
                    try:
                        cm = child.memory_full_info()
                        total_private += getattr(cm, "private", 0)
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass

                if total_private > peak_private:
                    peak_private = total_private

            except (psutil.NoSuchProcess, psutil.AccessDenied):
                break

            time.sleep(poll_interval)

        proc.wait()

        try:
            m = proc.memory_full_info()
            total_private = getattr(m, "private", 0)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            total_private = 0
        for child in proc.children(recursive=True):
            try:
                cm = child.memory_full_info()
                total_private += getattr(cm, "private", 0)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        if total_private > peak_private:
            peak_private = total_private

    except Exception:
        proc.kill()
        raise

    duration = time.perf_counter() - t0

    stdout = stderr = None
    if capture_output:
        stdout, stderr = proc.communicate()

    return stdout, stderr, duration, peak_private, proc.returncode


def run_excel_shrink(original_file, output_dir):
    logging.info(f"Running excel_shrink on {original_file}")
    output_file = os.path.join(output_dir, os.path.basename(original_file))

    shrink_script = os.path.join("..", "excel-shrink", "excel_shrink.py")

    cmd = [
        sys.executable,
        shrink_script,
        original_file,
        output_dir,
        "--chunk-size",
        "2097152",
    ]

    _, _, duration, peak_mem_bytes, returncode = run_and_measure(
        cmd, capture_output=False
    )

    if returncode == 0:
        logging.info(f"excel_shrink completed in {duration:.2f}s for {original_file}")
        logging.info(f"Peak memory usage: {peak_mem_bytes / (1024**2):.2f} MB")
    else:
        logging.error(
            f"excel_shrink failed (exit code {returncode}) for {original_file}"
        )

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

    csv_out = os.path.abspath(os.path.join(script_dir, args.csv_out))

    files_to_download = [
        {
            "url": "https://www.destatis.de/DE/Themen/Staat/Oeffentliche-Finanzen/Ausgaben-Einnahmen/Publikationen/Downloads-Ausgaben-und-Einnahmen/statistischer-bericht-rechnungsergebnis-kernhaushalt-gemeinden-2140331217005.xlsx?__blob=publicationFile&v=4",
            "filename": "statistischer-bericht-kernhaushalt-gemeinden.xlsx",
        },
        {
            "url": "https://www.destatis.de/DE/Themen/Gesellschaft-Umwelt/Verkehrsunfaelle/Publikationen/Downloads-Verkehrsunfaelle/verkehrsunfaelle-zeitreihen-xlsx-5462403.xlsx?__blob=publicationFile&v=19",
            "filename": "verkehrsunfaelle-zeitreihen-xlsx-5462403.xlsx",
        },
        {
            "url": "https://www.destatis.de/DE/Themen/Gesellschaft-Umwelt/Bevoelkerung/Wanderungen/Publikationen/Downloads-Wanderungen/wanderungen-2010120217005.xlsx?__blob=publicationFile&v=3",
            "filename": "wanderungen-2010120217005.xlsx",
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

    if not os.path.exists(shrunk_folder):
        os.makedirs(shrunk_folder)
        logging.info(f"Created shrunk folder: {shrunk_folder}")

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

        benchmark_results = {}
        for lib in library_names:
            logging.info(f"Measuring library: {lib} on file: {fname}")
            benchmark_results[lib] = call_measure_one(lib, original_path)
            logging.info(
                f"Results for {lib} on {fname}: Open Time={benchmark_results[lib][0]}, Save Time={benchmark_results[lib][1]}, Peak Memory={benchmark_results[lib][2]}"
            )

        logging.info(f"Running shrink on: {fname}")
        shrink_time, shrunk_path, shrink_peak_mem = run_excel_shrink(
            original_path, shrunk_folder
        )
        logging.info(
            f"Shrink results for {fname}: Time={shrink_time}, Shrunk Path={shrunk_path}, Peak Memory={shrink_peak_mem}"
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
        df = pd.read_csv(csv_file)
        _generate_memory_chart(df)
    except Exception as e:
        logging.error(f"Could not generate chart: {e}")


def _generate_complex_chart(df):

    logging.info("Using extended chart routine...")

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
    }
    missing = needed_cols - set(df.columns)
    if missing:
        logging.warning(f"Missing columns: {missing}. Cannot plot.")
        return

    library_info = [
        ("openpyxl(default)", "darkorange", "orangered"),
        ("pandas", "steelblue", "royalblue"),
        ("Microsoft Excel", "darkgreen", "limegreen"),
        ("R openxlsx", "red", "firebrick"),
        ("R readxl+writexl", "chocolate", "sienna"),
        ("Excel Shrink", "gray", "dimgray"),
    ]
    n_libs = len(library_info)

    files = sorted(df["File"].unique())
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
            return float(val) if pd.notna(val) else 0.0
        return 0.0

    total_bars_per_file = 2 * n_libs
    offsets = np.linspace(-0.7, 0.7, total_bars_per_file)
    bar_height = 0.1

    fig, (ax_time, ax_size) = plt.subplots(
        nrows=2, figsize=(14, 7), gridspec_kw={"height_ratios": [4, 1]}
    )

    y_positions_top = np.arange(n_files) * (bar_height * (total_bars_per_file + 2))

    for i, f in enumerate(files):
        base = y_positions_top[i]

        subdf = df[df["File"] == f]
        shrink_t = get_times(f, "Excel Shrink", "Shrink Time (s)")

        for lib_index, (lib_name, c_open, c_save) in enumerate(library_info):
            pos = base + offsets[lib_index]
            open_t = get_times(f, lib_name, "Original Open Time (s)")
            save_t = get_times(f, lib_name, "Original Save Time (s)")

            if i == 0:
                ax_time.barh(
                    pos,
                    open_t,
                    height=bar_height,
                    color=c_open,
                    label=f"{lib_name} open",
                )
                ax_time.barh(
                    pos,
                    save_t,
                    height=bar_height,
                    left=open_t,
                    color=c_save,
                    label=f"{lib_name} save",
                )
            else:
                ax_time.barh(pos, open_t, height=bar_height, color=c_open)
                ax_time.barh(pos, save_t, height=bar_height, left=open_t, color=c_save)

        for lib_index, (lib_name, c_open, c_save) in enumerate(library_info):
            pos = base + offsets[n_libs + lib_index]
            shrunk_open = get_times(f, lib_name, "Shrinked Open Time (s)")
            shrunk_save = get_times(f, lib_name, "Shrinked Save Time (s)")

            if i == 0:
                ax_time.barh(
                    pos,
                    shrink_t,
                    height=bar_height,
                    color="purple",
                    label="Shrink time" if lib_index == 0 else "",
                )
                ax_time.barh(
                    pos,
                    shrunk_open,
                    height=bar_height,
                    left=shrink_t,
                    color=c_open,
                    label=f"{lib_name} shrunk open",
                )
                ax_time.barh(
                    pos,
                    shrunk_save,
                    height=bar_height,
                    left=shrink_t + shrunk_open,
                    color=c_save,
                    label=f"{lib_name} shrunk save",
                )
            else:
                ax_time.barh(pos, shrink_t, height=bar_height, color="purple")
                ax_time.barh(
                    pos, shrunk_open, height=bar_height, left=shrink_t, color=c_open
                )
                ax_time.barh(
                    pos,
                    shrunk_save,
                    height=bar_height,
                    left=shrink_t + shrunk_open,
                    color=c_save,
                )

    ax_time.set_yticks(y_positions_top)

    def shortfile(fn):
        return fn if len(fn) <= 15 else fn[:5] + "..." + fn[-5:]

    ax_time.set_yticklabels([shortfile(f) for f in files])
    ax_time.invert_yaxis()
    ax_time.set_xlabel("Time (seconds)")
    ax_time.set_title("Time Comparison per Excel File")
    ax_time.legend(ncol=2, fontsize=8)

    group_size = df.groupby("File")[
        ["Original Size (bytes)", "Shrinked Size (bytes)"]
    ].mean()
    orig_size_mb_list = []
    shrunk_size_mb_list = []
    for f in files:
        if f in group_size.index:
            row = group_size.loc[f]
            orig_size = row["Original Size (bytes)"] or 0
            shrunk_size = row["Shrinked Size (bytes)"] or 0
        else:
            orig_size = 0
            shrunk_size = 0
        orig_size_mb_list.append(orig_size / (1024 * 1024))
        shrunk_size_mb_list.append(shrunk_size / (1024 * 1024))

    y_positions_bottom = np.arange(n_files) * 0.6
    file_bar_height = 0.25
    for i, f in enumerate(files):
        base = y_positions_bottom[i]
        ax_size.barh(
            base - file_bar_height * 0.3,
            orig_size_mb_list[i],
            height=file_bar_height,
            color="darkorange",
            label="Original Size (MB)" if i == 0 else "",
        )
        ax_size.barh(
            base + file_bar_height * 0.3,
            shrunk_size_mb_list[i],
            height=file_bar_height,
            color="steelblue",
            label="Shrinked Size (MB)" if i == 0 else "",
        )

    ax_size.set_yticks(y_positions_bottom)
    ax_size.set_yticklabels([shortfile(f) for f in files])
    ax_size.invert_yaxis()
    ax_size.set_xlabel("File Size (MB)")
    ax_size.set_title("File Size Comparison per Excel File")
    ax_size.legend(ncol=2, fontsize=8)

    plt.tight_layout()
    outname = "time_and_filesize_comparison_by_file.pdf"
    plt.savefig(outname, dpi=300, bbox_inches="tight")
    logging.info(f"Chart generated and saved to {outname}")
    plt.close(fig)


def _generate_memory_chart(df):
    required_cols = {"File", "Library", "Original Peak Memory", "Shrinked Peak Memory"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(
            f"Die folgenden notwendigen Spalten fehlen in der CSV: {missing}"
        )

    df["Original Peak Memory"] = pd.to_numeric(
        df["Original Peak Memory"], errors="coerce"
    ).fillna(0.0)
    df["Shrinked Peak Memory"] = pd.to_numeric(
        df["Shrinked Peak Memory"], errors="coerce"
    ).fillna(0.0)

    files = sorted(df["File"].unique())
    libraries = sorted(df["Library"].unique())

    n_files = len(files)
    if n_files == 0:
        raise ValueError("Keine Dateien in der CSV gefunden.")

    fig, axes = plt.subplots(nrows=n_files, figsize=(12, 4 * n_files), sharey=True)
    if n_files == 1:
        axes = [axes]

    plt.rcParams.update({"xtick.labelsize": 8})

    for ax, file_name in zip(axes, files):
        subset = df[df["File"] == file_name]

        n_libs = len(libraries)
        x = np.arange(n_libs)
        bar_width = 0.35

        original_mem = [
            (
                float(subset[subset["Library"] == lib]["Original Peak Memory"].iloc[0])
                / (1024**2)
                if lib in subset["Library"].values
                else 0.0
            )
            for lib in libraries
        ]
        shrunk_mem = [
            (
                float(subset[subset["Library"] == lib]["Shrinked Peak Memory"].iloc[0])
                / (1024**2)
                if lib in subset["Library"].values
                else 0.0
            )
            for lib in libraries
        ]

        ax.bar(
            x - bar_width / 2,
            original_mem,
            width=bar_width,
            label="Original Peak Memory (MiB)",
            color="skyblue",
            edgecolor="black",
        )

        ax.bar(
            x + bar_width / 2,
            shrunk_mem,
            width=bar_width,
            label="Shrinked Peak Memory (MiB)",
            color="orange",
            edgecolor="black",
        )

        ax.set_title(f"Memory Consumption:\n'{file_name}'", fontsize=10, pad=10)
        ax.set_xlabel("Library", fontsize=9)
        ax.set_ylabel("Peak Memory", fontsize=9)
        ax.ticklabel_format(style="plain", axis="y")
        ax.set_xticks(x)
        ax.set_xticklabels(libraries, rotation=45, ha="right", fontsize=8)
        ax.legend(fontsize=8)
        ax.grid(axis="x", linestyle="--", alpha=0.5)

    plt.tight_layout()
    output_path = "memory_consumption.pdf"
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
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
