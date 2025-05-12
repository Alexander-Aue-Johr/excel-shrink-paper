#!/usr/bin/env python3

import os
import sys
import subprocess
import importlib
import urllib.request
import logging

# ============================================================================
# 🛠 Helper: Ensure required packages
# ============================================================================

required_packages = ["pandas", "openpyxl", "matplotlib", "numpy"]

def install_package(package):
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--user", package
        ])
    except subprocess.CalledProcessError:
        print(f"⚠️ Normal --user install for '{package}' failed. Trying with --break-system-packages...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--user", "--break-system-packages", package
        ])

# Try to import each package, install if missing
for package in required_packages:
    try:
        importlib.import_module(package)
    except ImportError:
        print(f"🔹 Package '{package}' not found. Attempting installation...")
        install_package(package)
        
import os
import sys
import argparse
import subprocess
import time
import logging
import json

# Only import heavy libraries in measure-one mode to avoid overhead in controller mode
# We will import them conditionally in the child process.

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ========================================================================
#                       measure_one() mode
# ========================================================================
#
# This mode measures a single library on a single file and returns JSON.
# The controller calls this mode for each combination to avoid memory bloat.
#

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

    measure_func = lib_funcs[library_name]

    open_t, save_t = measure_func(file_path)
    # If error or None, None returned
    if open_t is None or save_t is None:
        result = {
            "library": library_name,
            "file": os.path.basename(file_path),
            "Original Open Time (s)": None,
            "Original Save Time (s)": None,
        }
    else:
        result = {
            "library": library_name,
            "file": os.path.basename(file_path),
            "Original Open Time (s)": open_t,
            "Original Save Time (s)": save_t,
        }

    print(json.dumps(result))


# ========================================================================
#                  The actual library benchmark functions
# ========================================================================

def benchmark_openpyxl_default(file_path):
    """
    Measure openpyxl.load_workbook + wb.save() with standard settings.
    """
    from openpyxl import load_workbook
    import time
    import tempfile
    import os

    name = "openpyxl (default)"
    try:
        start = time.perf_counter()
        wb = load_workbook(filename=file_path)
        open_time = time.perf_counter() - start

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp_name = tmp.name

        start = time.perf_counter()
        wb.save(tmp_name)
        save_time = time.perf_counter() - start

        os.remove(tmp_name)
        return open_time, save_time
    except Exception as e:
        logging.error(f"{name} error: {e}")
        return None, None

def benchmark_pandas(file_path):
    """
    Benchmark: Pandas reads ALL sheets with read_excel, then writes them back.
    """
    import time
    import pandas as pd
    import tempfile
    import os

    name = "pandas"
    try:
        start = time.perf_counter()
        dataframes = pd.read_excel(file_path, sheet_name=None)
        open_time = time.perf_counter() - start

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp_name = tmp.name

        start = time.perf_counter()
        with pd.ExcelWriter(tmp_name, engine='openpyxl') as writer:
            for sheet_name, df in dataframes.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        save_time = time.perf_counter() - start

        os.remove(tmp_name)
        return open_time, save_time
    except Exception as e:
        logging.error(f"{name} error: {e}")
        return None, None

def benchmark_excel_com(file_path):
    """
    Benchmark: via win32com (only on Windows). 
    First runs excel_shrink.py --only-clean-workbook on the original,
    then opens the cleaned workbook with Excel COM and saves it.
    """
    import time
    import os
    import tempfile
    import subprocess
    import logging

    name = "Microsoft Excel"
    try:
        import win32com.client
    except ImportError:
        logging.warning("win32com not available; skipping Excel COM benchmark.")
        return None, None

    # 1) Prepare a temp dir and run excel_shrink with only-clean-workbook
    with tempfile.TemporaryDirectory() as shrink_dir:
        shrink_script = os.path.join("..", "excel-shrink", "excel_shrink.py")
        logging.info(f"Running excel_shrink --only-clean-workbook on {file_path}")
        subprocess.run([
            sys.executable,
            shrink_script,
            "--only-clean-workbook",
            file_path,
            shrink_dir
        ], check=True)

        cleaned_path = os.path.join(shrink_dir, os.path.basename(file_path))
        if not os.path.exists(cleaned_path):
            logging.error(f"Cleaned workbook not found at {cleaned_path}")
            return None, None

        # 2) Now open the cleaned workbook via COM
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.Application.DisplayAlerts = False

            abs_file = os.path.abspath(cleaned_path)
            start = time.perf_counter()
            wb = excel.Workbooks.Open(abs_file)
            open_time = time.perf_counter() - start

            # save to a temporary file to measure save time
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp_name = tmp.name

            start = time.perf_counter()
            wb.SaveAs(os.path.abspath(tmp_name))
            wb.Close()
            save_time = time.perf_counter() - start

            excel.Quit()
            os.remove(tmp_name)
            return open_time, save_time

        except Exception as e:
            logging.error(f"{name} error after cleaning: {e}")
            return None, None


def benchmark_r_openxlsx(file_path):
    """
    Uses R openxlsx to open & save. We generate an R script, run it,
    parse times from stdout.
    """
    import time
    import tempfile
    import os
    import subprocess
    import shutil

    name = "R openxlsx"

    if shutil.which("Rscript") is None:
        logging.warning("Rscript not found; skipping R openxlsx benchmark.")
        return None, None

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_name = tmp.name

    r_script = f'''
    args <- commandArgs(trailingOnly = TRUE)
    file_in <- args[1]
    file_out <- args[2]

    library(openxlsx)

    t1 <- proc.time()
    wb <- loadWorkbook(file_in)
    t2 <- proc.time() - t1
    open_time <- t2["elapsed"][[1]]

    t3 <- proc.time()
    saveWorkbook(wb, file_out, overwrite = TRUE)
    t4 <- proc.time() - t3
    save_time <- t4["elapsed"][[1]]

    cat(open_time, save_time, sep=",")
    '''

    with tempfile.NamedTemporaryFile(delete=False, suffix=".R") as rfile:
        rfile_name = rfile.name
        rfile.write(r_script.encode("utf-8"))

    try:
        result = subprocess.run(
            ["Rscript", rfile_name, file_path, tmp_name],
            capture_output=True, text=True, check=True
        )
        output_str = result.stdout.strip()
        open_time_str, save_time_str = output_str.split(",")
        open_time = float(open_time_str)
        save_time = float(save_time_str)

        os.remove(tmp_name)
        return open_time, save_time
    except subprocess.CalledProcessError as e:
        logging.error(f"{name} subprocess error: {e}")
        return None, None
    except Exception as e:
        logging.error(f"{name} error: {e}")
        return None, None
    finally:
        if os.path.exists(rfile_name):
            os.remove(rfile_name)

def benchmark_r_readxl_writexl(file_path):
    """
    Uses R readxl + writexl to open all sheets, then write them.
    """
    import time
    import tempfile
    import os
    import subprocess
    import shutil

    name = "R readxl+writexl"

    if shutil.which("Rscript") is None:
        logging.warning("Rscript not found; skipping R readxl+writexl benchmark.")
        return None, None

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_name = tmp.name

    r_script = f'''
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
    '''

    with tempfile.NamedTemporaryFile(delete=False, suffix=".R") as rfile:
        rfile_name = rfile.name
        rfile.write(r_script.encode("utf-8"))

    try:
        result = subprocess.run(
            ["Rscript", rfile_name, file_path, tmp_name],
            capture_output=True, text=True, check=True
        )
        output_str = result.stdout.strip()
        open_time_str, save_time_str = output_str.split(",")
        open_time = float(open_time_str)
        save_time = float(save_time_str)

        os.remove(tmp_name)
        return open_time, save_time
    except subprocess.CalledProcessError as e:
        logging.error(f"{name} subprocess error: {e}")
        return None, None
    except Exception as e:
        logging.error(f"{name} error: {e}")
        return None, None
    finally:
        if os.path.exists(rfile_name):
            os.remove(rfile_name)


# ========================================================================
#                            EXCEL-SHRINK
# ========================================================================

def run_excel_shrink(original_file, output_dir):
    """
    Calls external excel_shrink.py in ../excel-shrink/ with subprocess.
    Returns (shrink_time, shrunk_path).
    """
    logging.info(f"Running excel_shrink on {original_file}")
    output_file = os.path.join(output_dir, os.path.basename(original_file))
    start = time.perf_counter()

    # Adjust the path to excel_shrink.py as needed:
    shrink_script = os.path.join("..", "excel-shrink", "excel_shrink.py")

    subprocess.run([sys.executable, shrink_script, original_file, output_dir],
                   check=True)

    shrink_time = time.perf_counter() - start
    logging.info(f"excel_shrink completed in {shrink_time:.2f}s for {original_file}")

    return shrink_time, output_file


# ========================================================================
#                    Controller Mode (run all)
# ========================================================================

def controller_main(args):
    """
    In 'controller' mode, we loop over all files and libraries, call measure-one
    as a separate subprocess for each scenario, then run shrink, measure shrunk file,
    and finally write CSV and generate the chart.
    """
    import logging
    import csv
    import os
    import pandas as pd

    # 📍 Resolve paths relative to script location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.abspath(os.path.join(script_dir, args.input_folder))
    shrunk_folder = os.path.abspath(os.path.join(script_dir, args.shrunk_folder))
    csv_out = os.path.abspath(os.path.join(script_dir, args.csv_out))

    # ============================================================================
    # 📥 Download Excel files from Destatis
    # ============================================================================

    files_to_download = [
        {
            "url": "https://www.destatis.de/DE/Themen/Staat/Oeffentliche-Finanzen/Ausgaben-Einnahmen/Publikationen/Downloads-Ausgaben-und-Einnahmen/statistischer-bericht-rechnungsergebnis-kernhaushalt-gemeinden-2140331217005.xlsx?__blob=publicationFile&v=4",
            "filename": "statistischer-bericht-kernhaushalt-gemeinden.xlsx"
        },
        {
            "url": "https://www.destatis.de/DE/Themen/Gesellschaft-Umwelt/Verkehrsunfaelle/Publikationen/Downloads-Verkehrsunfaelle/verkehrsunfaelle-zeitreihen-xlsx-5462403.xlsx?__blob=publicationFile&v=19",
            "filename": "verkehrsunfaelle-zeitreihen-xlsx-5462403.xlsx"
        },
        {
            "url": "https://www.destatis.de/DE/Themen/Gesellschaft-Umwelt/Bevoelkerung/Wanderungen/Publikationen/Downloads-Wanderungen/wanderungen-2010120217005.xlsx?__blob=publicationFile&v=3",
            "filename": "wanderungen-2010120217005.xlsx"
        }
    ]

    # Directory to store downloaded files
    output_folder = "scripts/loading_times_size_reduction/input"

    # ============================================================================
    # 📥 Download Logic
    # ============================================================================

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

    # 1) Prepare shrink folder
    if not os.path.exists(shrunk_folder):
        os.makedirs(shrunk_folder)
        logging.info(f"Created shrunk folder: {shrunk_folder}")

    # 2) Collect Excel files
    all_files = [f for f in os.listdir(input_folder)
                 if f.lower().endswith('.xlsx') and os.path.isfile(os.path.join(input_folder, f))]
    logging.info(f"Found {len(all_files)} Excel files in {input_folder}.")

    # 3) Define all library names we want to measure (data_only removed)
    library_names = [
        "openpyxl(default)",
        "pandas",
        "Microsoft Excel",
        "R openxlsx",
        "R readxl+writexl",
    ]

    # We'll collect rows of data for final CSV
    rows = []

    for f in all_files:
        original_path = os.path.join(input_folder, f)
        logging.info(f"Processing file: {f}")
        try:
            original_size = os.path.getsize(original_path)
        except:
            original_size = None

        # 4a) For each library, measure the original file in a subprocess
        benchmark_results = {}
        for lib_name in library_names:
            open_t, save_t = call_measure_one(lib_name, original_path)
            benchmark_results[lib_name] = (open_t, save_t)

        # 4b) Shrink the file once
        try:
            shrink_time, shrunk_path = run_excel_shrink(original_path, shrunk_folder)
        except Exception as e:
            shrink_time = None
            shrunk_path = None
            logging.error(f"excel_shrink error for {f}: {e}")

        if shrunk_path and os.path.exists(shrunk_path):
            shrunk_size = os.path.getsize(shrunk_path)
        else:
            shrunk_size = None
            logging.warning(f"Shrunk file not found for {f}")

        # 4c) For each library, measure the shrunk file (both open and save times)
        shrunk_results = {}
        if shrunk_path and os.path.exists(shrunk_path):
            for lib_name in library_names:
                open_t2, save_t2 = call_measure_one(lib_name, shrunk_path)
                shrunk_results[lib_name] = (open_t2, save_t2)
        else:
            for lib_name in library_names:
                shrunk_results[lib_name] = (None, None)

        # 4d) Collect final rows. We store one row per library.
        for lib_name in library_names:
            (orig_open, orig_save) = benchmark_results[lib_name]
            (shrunk_open, shrunk_save) = shrunk_results[lib_name]

            # Only store complete measurements
            if (orig_open is not None and orig_save is not None and
                shrunk_open is not None and shrunk_save is not None and
                shrink_time is not None):
                row = {
                    "File": f,
                    "Library": lib_name,
                    "Original Open Time (s)": orig_open,
                    "Original Save Time (s)": orig_save,
                    "Shrink Time (s)": shrink_time,
                    "Shrinked Open Time (s)": shrunk_open,
                    "Shrinked Save Time (s)": shrunk_save,
                    "Original Size (bytes)": original_size,
                    "Shrinked Size (bytes)": shrunk_size
                }
                rows.append(row)

        logging.info(f"Finished processing file: {f}")

    # 5) Write CSV
    fieldnames = [
        "File", "Library",
        "Original Open Time (s)", "Original Save Time (s)",
        "Shrink Time (s)", "Shrinked Open Time (s)", "Shrinked Save Time (s)",
        "Original Size (bytes)", "Shrinked Size (bytes)"
    ]
    with open(csv_out, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)

    logging.info(f"Measurements saved to {csv_out}")

    # 6) Generate the chart
    generate_chart(csv_out)


def call_measure_one(library_name, file_path):
    """
    Calls this script in 'measure-one' mode for the given library and file.
    Returns (open_time, save_time) or (None, None) on error.
    """
    import json

    cmd = [
        sys.executable,
        os.path.abspath(__file__),
        "measure-one",
        "--library", library_name,
        "--file", file_path
    ]
    logging.info(f"Subprocess: {cmd}")
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    except subprocess.CalledProcessError as e:
        logging.error(f"Error in measure-one for {library_name} with {file_path}: {e}")
        return (None, None)

    try:
        data = json.loads(result.stdout.strip())
        if "Original Open Time (s)" in data and "Original Save Time (s)" in data:
            return (data["Original Open Time (s)"], data["Original Save Time (s)"])
        else:
            return (None, None)
    except json.JSONDecodeError:
        logging.error(f"Failed to parse JSON from measure-one: {result.stdout}")
        return (None, None)


# ========================================================================
#                           Chart Generation
# ========================================================================

def generate_chart(csv_file):
    """
    Generate a chart from the CSV file.
    """
    import pandas as pd
    import matplotlib.pyplot as plt
    import logging

    logging.info(f"Generating chart from {csv_file}")
    if not os.path.exists(csv_file):
        logging.warning(f"CSV file {csv_file} not found; skipping chart.")
        return

    df = pd.read_csv(csv_file)

    if df.empty:
        logging.warning("CSV file is empty; no chart to generate.")
        return

    try:
        # Use the extended chart routine
        _generate_complex_chart(df)
    except Exception as e:
        logging.error(f"Could not generate chart: {e}")


def _generate_complex_chart(df):
    """
    Extended chart function that plots two groups of bars per file:
      - One group shows original measurements (stacked: open then save)
      - The other group shows shrunk file measurements (stacked:
          shrink time + shrunk open + shrunk save)
    Each file will display 2 bars per library.
    """
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import logging

    logging.info("Using extended chart routine...")

    if df.empty:
        logging.warning("DataFrame is empty; no chart to plot.")
        return

    needed_cols = {
        "File", "Library",
        "Original Open Time (s)", "Original Save Time (s)",
        "Shrink Time (s)", "Shrinked Open Time (s)", "Shrinked Save Time (s)",
        "Original Size (bytes)", "Shrinked Size (bytes)"
    }
    missing = needed_cols - set(df.columns)
    if missing:
        logging.warning(f"Missing columns: {missing}. Cannot plot.")
        return

    # Define libraries to show (matching CSV library names)
    library_info = [
        ("openpyxl(default)", "darkorange", "orangered", "peru"),
        ("pandas", "steelblue", "royalblue", "dodgerblue"),
        ("Microsoft Excel", "darkgreen", "green", "limegreen"),
        ("R openxlsx", "red", "firebrick", "lightsalmon"),
        ("R readxl+writexl", "chocolate", "sienna", "tan"),
    ]
    n_libs = len(library_info)

    # Collect the distinct files in sorted order
    files = sorted(df["File"].unique())
    n_files = len(files)
    if n_files == 0:
        logging.info("No files found in data; nothing to plot.")
        return

    # Create a helper lookup for each (file, library) and column
    temp = df.set_index(["File", "Library"], drop=False)

    def get_times(f, lib, col):
        if (f, lib) in temp.index and col in temp.columns:
            val = temp.loc[(f, lib), col]
            if isinstance(val, pd.Series):
                val = val.iloc[0]
            return float(val)
        return 0.0

    # Calculate offsets: Two groups per file (original and shrunk) per library
    total_bars_per_file = 2 * n_libs  # first n_libs: original; next n_libs: shrunk measurements
    offsets = np.linspace(-0.7, 0.7, total_bars_per_file)
    bar_height = 0.1

    # Create figure with two subplots: time and file size
    fig, (ax_time, ax_size) = plt.subplots(
        nrows=2, figsize=(14, 7), gridspec_kw={'height_ratios': [4, 1]}
    )

    y_positions_top = np.arange(n_files) * (bar_height * 20)

    for i, f in enumerate(files):
        base = y_positions_top[i]

        # For each file, obtain the shrink time (same for all libraries)
        subdf = df[df["File"] == f]
        shrink_t = subdf["Shrink Time (s)"].mean() if not subdf.empty else 0.0

        # 1) Original measurements: one bar per library (stacked: open then save)
        for lib_index, (lib_name, c_open, c_save, c_shrunk) in enumerate(library_info):
            offset_idx = lib_index  # indices 0 .. n_libs-1
            pos = base + offsets[offset_idx]
            open_t = get_times(f, lib_name, "Original Open Time (s)")
            save_t = get_times(f, lib_name, "Original Save Time (s)")

            if i == 0:
                ax_time.barh(pos, open_t, height=bar_height, color=c_open,
                             label=f"{lib_name} original open")
                ax_time.barh(pos, save_t, height=bar_height, left=open_t, color=c_save,
                             label=f"{lib_name} original save")
            else:
                ax_time.barh(pos, open_t, height=bar_height, color=c_open)
                ax_time.barh(pos, save_t, height=bar_height, left=open_t, color=c_save)

        # 2) Shrunk measurements: one bar per library (stacked: shrink time + shrunk open + shrunk save)
        for lib_index, (lib_name, c_open, c_save, c_shrunk) in enumerate(library_info):
            offset_idx = n_libs + lib_index  # indices n_libs .. 2*n_libs-1
            pos = base + offsets[offset_idx]
            shrunk_open = get_times(f, lib_name, "Shrinked Open Time (s)")
            shrunk_save = get_times(f, lib_name, "Shrinked Save Time (s)")
            # Stack three segments: shrink time (global), then shrunk open, then shrunk save.
            if i == 0:
                ax_time.barh(pos, shrink_t, height=bar_height, color="purple",
                             label="Shrink time")
                ax_time.barh(pos, shrunk_open, height=bar_height, left=shrink_t, color=c_shrunk,
                             label=f"{lib_name} shrunk open")
                ax_time.barh(pos, shrunk_save, height=bar_height, left=shrink_t + shrunk_open, color=c_save,
                             label=f"{lib_name} shrunk save")
            else:
                ax_time.barh(pos, shrink_t, height=bar_height, color="purple")
                ax_time.barh(pos, shrunk_open, height=bar_height, left=shrink_t, color=c_shrunk)
                ax_time.barh(pos, shrunk_save, height=bar_height, left=shrink_t + shrunk_open, color=c_save)

    ax_time.set_yticks(y_positions_top)
    def shortfile(fn):
        return fn if len(fn) <= 15 else fn[:5] + "..." + fn[-5:]
    ax_time.set_yticklabels([shortfile(f) for f in files])
    ax_time.invert_yaxis()
    ax_time.set_xlabel('Time (seconds)')
    ax_time.set_title('Time Comparison per Excel File')
    ax_time.legend(ncol=2, fontsize=8)

    # Bottom chart: File Size Comparison
    group_size = df.groupby('File')[["Original Size (bytes)", "Shrinked Size (bytes)"]].mean()
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
    for i in range(n_files):
        base = y_positions_bottom[i]
        ax_size.barh(base - file_bar_height*0.3, orig_size_mb_list[i], height=file_bar_height,
                     color="darkorange", label="Original Size (MB)" if i == 0 else "")
        ax_size.barh(base + file_bar_height*0.3, shrunk_size_mb_list[i], height=file_bar_height,
                     color="steelblue", label="Shrinked Size (MB)" if i == 0 else "")

    ax_size.set_yticks(y_positions_bottom)
    ax_size.set_yticklabels([shortfile(f) for f in files])
    ax_size.invert_yaxis()
    ax_size.set_xlabel('File Size (MB)')
    ax_size.set_title('File Size Comparison per Excel File')
    ax_size.legend(ncol=2, fontsize=8)

    plt.tight_layout()
    outname = 'time_and_filesize_comparison_by_file.pdf'
    plt.savefig(outname, dpi=300, bbox_inches='tight')
    logging.info(f"Chart generated and saved to {outname}")
    plt.show()


# ========================================================================
#                            MAIN DISPATCH
# ========================================================================

def main():
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description="Run measurements on Excel files and generate comparison chart."
    )

    subparsers = parser.add_subparsers(dest="mode", help="Subcommand: controller or measure-one")

    # measure-one subcommand bleibt unverändert
    ctrl_parser = subparsers.add_parser("controller", help="Controller mode: full benchmark")
    meas_parser = subparsers.add_parser("measure-one", help="Measure one library on one file.")
    meas_parser.add_argument("--library", required=True, help="Library name, e.g. 'openpyxl(default)'")
    meas_parser.add_argument("--file", required=True, help="Path to the .xlsx file")

    args, unknown = parser.parse_known_args()

    # Standard-Speicherort relativ zum Skriptpfad (nur für controller)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(script_dir, "input")
    shrunk_folder = os.path.join(script_dir, "shrunk_files")
    csv_out = os.path.join(script_dir, "excel_benchmarks.csv")

    if not args.mode:
        args.mode = "controller"

    if args.mode == "measure-one":
        measure_one_main(args)
    else:
        class Args:
            pass
        args = Args()
        args.input_folder = input_folder
        args.shrunk_folder = shrunk_folder
        args.csv_out = csv_out
        args.run_tests = True

        if not args.run_tests and os.path.exists(args.csv_out):
            logging.info(f"CSV file {args.csv_out} already exists. Skipping measurements.")
            generate_chart(args.csv_out)
            sys.exit(0)
        controller_main(args)


if __name__ == "__main__":
    main()
