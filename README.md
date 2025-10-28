# Excel Benchmarking & Shrinking – Supplementary Material

This repository accompanies the submission:

**“Excel Shrink: High-Performance Bytestream
Parsing and Cleaning of Bloated Excel Spreadsheets”**  

It provides all scripts and configurations necessary to reproduce the benchmark results and figures presented in the paper.

---

## 📂 Repository Structure

| Path | Description |
|------|--------------|
| `excel-shrink-paper.tex` | LaTeX source for the paper (for transparency, not required for evaluation). |
| `scripts/loading_times_size_reduction/` | Core measurement logic for benchmarking, shrinking, and chart generation. |
| `scripts/scrape-destatis-xlsx-files/` | Crawls and downloads public `.xlsx` files from the German Statistical Office (DESTATIS). |
| `scripts/excel_shrink/` | Stand-alone implementation of the *Excel Shrink* optimizer that removes redundant XML markup. |
| `scripts/excel_shrink_analyzer/` | Variant used for analysis of shrink ratios and XML component frequencies. |
| `scripts/analyze_destatis_xlsx_files/` | Aggregates results and generates global distributions of file-size reduction. |

All relevant VS Code launch configurations for reproducibility are included in  
`.vscode/launch.json` and `.vscode/tasks.json`.

---

## 🚀 Quick Start (Automatic Setup)

No manual installation is required.  
Each script automatically bootstraps its environment and installs missing dependencies.

Example – to reproduce the main benchmark figures:

```bash
python scripts/loading_times_size_reduction/loading_times_size_reduction_chart.py \
       --input-folder ./input --shrunk-folder ./shrunk
```

All dependencies will be installed on first run.

---

## 🧩 Alternative (VS Code Integration)

You can also execute all workflows directly from **VS Code** via the pre-configured launch entries:

| Launch Configuration                                         | Purpose                                                                 |
| ------------------------------------------------------------ | ----------------------------------------------------------------------- |
| **Compare Excel File Sizes**                                 | Compares original and shrunk workbook sizes.                            |
| **Scrape and Download Destatis XLSX Files**                  | Performs web scraping and download of source workbooks.                 |
| **Download Destatis XLSX Files**                             | Downloads previously discovered URLs (skip new crawl).                  |
| **Shrink Destatis XLSX Files (excel-shrink)**                | Runs the core shrinker on all downloaded files.                         |
| **Analyze Destatis XLSX Files (excel-shrink-analyzer)**      | Evaluates reduction ratios and XML statistics.                          |
| **Generate Distribution of Size Reduction Chart**            | Produces the histogram of reduction percentages.                        |
| **Generate Chart Loading Times Size Reduction**              | Runs full benchmark and creates the main time/size figures.             |
| **Generate Chart Loading Times Size Reduction (only chart)** | Regenerates plots from existing CSV data without re-running benchmarks. |

All launch entries automatically call the `py:bootstrap` task, which:

1. creates a virtual environment (`.venv`) if missing, and
2. installs required Python packages via `pip install -r requirements.txt`.

---

## 🧪 What the Benchmark Does

For each `.xlsx` file, the benchmark:

1. Measures **open/save times** and **memory usage** using:

   * `openpyxl`
   * `pandas`
   * `Microsoft Excel` (COM interface, Windows only)
   * `R openxlsx`
   * `R readxl + writexl`
2. Performs **lossless XML shrinking** using *Excel Shrink*.
3. Re-measures performance on the cleaned files.
4. Generates:

   * `excel_benchmarks.csv`
   * `time_and_filesize_comparison_by_file.pdf`

---

## 📦 Dependencies

Handled automatically by the bootstrap step.

Manual installation (optional):

```bash
pip install pandas openpyxl matplotlib numpy
```

Windows-only:

```bash
pip install pywin32
```

R-based benchmarks require:

```r
install.packages(c('openxlsx','readxl','writexl'), repos='https://cloud.r-project.org')
```

If R or any of these packages are missing, R-related benchmarks are skipped gracefully.

---

## 🧘 Optional: Manual Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate  # Linux/macOS
# or
.venv\Scripts\activate     # Windows
```

Then run any script or VS Code launch target.

---

---

## 🏛️ Data Acknowledgement

We gratefully acknowledge the **Statistisches Bundesamt (DESTATIS)** for making the analyzed `.xlsx` datasets publicly available and permitting the publication of aggregated results.

---

## 📜 License

To be specified upon acceptance.
