
# Excel Benchmarking & Shrinking – Supplementary Material

This repository contains the supplementary material for the anonymous submission:

**“Efficient Measurement of Excel File Handling: A Cross-Library Benchmark with File Shrinking”**  
(submission ID: [REDACTED])

The repository is structured as follows:

- `sigconf-excel-shrink.tex`: The LaTeX source for the paper (required for reproducibility, but not central to the evaluation)
- `scripts/loading_times_size_reduction/`: Core implementation for measurement, shrinking, and chart generation
- Example input files (optional, to be added on request)

We recommend reviewers to begin with the script:

```
scripts/loading_times_size_reduction/loading_times_size_reduction_chart.py
```

which includes all benchmark logic and produces the main result figures in the paper.

---

## 🚀 Quick Start (No Setup Required)

This script is designed to run out-of-the-box on any machine with Python 3.8+ and basic `pip` access.

```bash
python loading_times_size_reduction_chart.py --input-folder ./input --shrunk-folder ./shrunk
```

All required Python dependencies are automatically installed on first run if missing.  
No virtual environment or external package manager is required.

---

## 🧪 What It Does

For each input `.xlsx` file:

- Runs open/save benchmarks using:
  - `openpyxl`
  - `pandas`
  - Microsoft Excel COM (Windows only)
  - `R openxlsx` (if R is installed)
  - `R readxl + writexl` (if R is installed)
- Performs lossless optimization (shrinking)
- Reruns all benchmarks on the optimized version
- Outputs:
  - `excel_benchmarks.csv`
  - `time_and_filesize_comparison_by_file.pdf`

---

## 📦 Dependencies

All dependencies are handled internally by the script via `pip`.

If preferred, they can be installed manually:

- `pandas`, `openpyxl`, `matplotlib`, `numpy`
- `pywin32` (Windows only)

R-related benchmarks require `Rscript` to be available in `PATH`, and the following R packages:

- `openxlsx`
- `readxl`
- `writexl`

Check if these packages are installed:

```bash
Rscript -e "pkgs <- c('openxlsx','readxl','writexl'); sapply(pkgs, function(p) as.character(packageVersion(p)))"
```

Install them if missing:

```bash
Rscript -e "install.packages(c('openxlsx','readxl','writexl'), repos='https://cloud.r-project.org')"
```

If R or any required packages are not found, the script will automatically skip R-related benchmarks.

---

## 🧘 Optional: Virtual Environments (`venv`)

To isolate the environment, standard Python `venv` can be used:

```bash
# Create and activate (Linux/macOS)
python3 -m venv .venv
source .venv/bin/activate

# or on Windows
.venv\Scripts\activate
```

However, the script runs without `venv` by default.

---

## 🧠 Design Principles

This tool is designed to be:

- ✅ Self-contained
- ✅ Robust in unknown environments
- ✅ Compatible with non-technical users and automated test systems

Optional components (such as R or Excel COM) are autodetected and skipped if unavailable.

---

## 📎 Citation / Context

This script is part of an anonymous submission to a peer-reviewed venue.  
Please do not attribute authorship during the review process.

---

## License

To be specified upon acceptance.
```