# Enron spreadsheet corpus

This experiment reconstructs spreadsheet inputs from the 159 ZIP archives in
the [EDRM Enron v2 XML collection](https://archive.org/details/edrm.enron.email.data.set.v2.xml).
The existing SheetJS subset was useful for initial tests, but its manifest also
lists files absent from that snapshot. We therefore inventory the original
archives and retain a source mapping for every candidate, duplicate, and exception.
This is a defined source collection, not a claim to have every Enron document.

## Run

Windows, Python 3.9 or later, and desktop Excel are required. The new workflow
uses the Python standard library and PowerShell; it does not install the older
experiments' plotting dependencies. Keep Windows signed in while Excel runs.

From the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start_enron.ps1
python scripts/enron.py status
python scripts/enron.py stop
```

The launcher runs in the background with a frozen copy of its code. After all
downloads pass verification, it inventories candidates and checks saved study
results before starting Excel for files without a match. Running the launcher again resumes checkpoints.
It cannot run twice against the same data directory. A reboot stops the process;
start it again to resume. Failed conversion attempts are retained and skipped on
resume; use `-RetryReview` to retry them after investigating the cause.

Data defaults to `F:\corpora\enron-v2`; pass `-DataRoot` to use another directory.
The source ZIPs alone need about 79 GB (decimal), plus extracted files and outputs.

| Location in the data directory | Contents |
|---|---|
| `sources.json`, `raw/` | Frozen archive list and verified downloads |
| `inventory.jsonl`, `extracted/` | Member inventory and candidate originals, with source folders retained |
| `xlsx/`, `conversions.jsonl` | Accepted XLSX files and all source-to-output mappings |
| `reuse/`, `variants/` | Saved study bundles and imported shrink variants |
| `review/files/`, `review/issues/` | Copies of problem files and their current diagnoses |
| `_work/`, `runs/`, `launches/` | Attempt evidence, run records, code snapshots, and logs |
| `status.json`, `summary.json` | Current progress and completed-run counts |

## Queued preparations

The saved queue in `F:\corpora\preparation-queue` waits for Enron 1/4,
imports the verified 4/4 transfer, retries failed Enron preparations, then
prepares EUSES. EUSES downloads independently while Enron runs. Saved valid
preparations are reused; imported measurements keep their original provenance.
EUSES has its own study output and Observatory port, 8772.

`status.json` identifies the current step; `checkpoints.json` and `jobs/` retain
completed steps, launch receipts and logs. Resume the same queue with:

```powershell
python scripts/run_preparation_queue.py F:\corpora\preparation-queue\config.json
```

The queue checks the runner's final preparation checkpoint and released output
lock. A nonzero exit caused by recorded skips does not discard completed work.
An interrupted run blocks later steps. Creating `stop.request` in the queue
directory prevents further steps; an active preparation is left running.

## Compare with an older corpus

```powershell
python scripts/enron.py compare F:\corpora\enron.zip
```

The comparison searches the original bytes and matches SHA-256 values, including
renamed files and spreadsheet formats beyond `.xls`. It reuses completed
inventories and checkpoints each remaining source archive. It does not open Excel.
Reports are written to `comparison/` in the data directory:

- `new-only.csv`: detected spreadsheet candidates absent from the old archive.
- `review-only.csv`: additional files whose spreadsheet format remains uncertain.
- `matches.csv`: byte-identical files and their source locations.
- `old-not-found.csv`: old files not found in the searched sources.
- `scan-issues.csv`: unreadable entries and uninspected nested archives.

`summary.json` separates file occurrences from unique contents. The reverse
difference remains provisional while archives are unsearched or have coverage
gaps. File headers guide the search; this is not a CRC audit of unrelated files.

## Conversion and review

Existing preparation is reused by the original source SHA-256, regardless of
filename or source folder. An import checks package integrity and copies the
saved baseline, `default`, `maximum-cleaning`, and `too-aggressive-cleaning`
outputs together. It does not rerun Excel or shrinking. Historical producer
settings, baseline repairs, and validation remain explicit in the manifest;
an import is not recorded as a fresh prompt-free Excel reopen.

Preserve preparations from each machine before the older study cleans them up:

```powershell
python scripts/enron.py export-study F:\repos\excel-shrink-rust\.tmp\regression\study F:\corpora\enron-v2\reuse --watch --stop-file F:\corpora\enron-v2\reuse-export.stop
```

Copy an exported directory from another machine, then list its path in the data
directory's `reuse-sources.json`: `{"caches": ["G:/enron-study-export"]}`.
The launcher also starts this preservation step for the local study directory;
pass `-StudyRoot` for another location. The local `reuse/` directory is always checked. Keep exports available until
preparation finishes. Conflicting preparations stay separate; one complete
bundle supplies a baseline and its matching variants. Damaged imports go to
review instead of silently repeating the work. Only unmatched sources require
new conversion. No benchmark runs as part of preparation.

Content detection supplements filename extensions. Ambiguous formats remain
explicit review entries. Paths are escaped for Windows; original ZIP paths stay
in the inventory. Equal SHA-256 contents share one conversion and retain all
their source occurrences. Originals are never moved or overwritten.

The first run uses one isolated Excel process at a time. Each stage has a
120-second limit and each file a 420-second limit. A timeout closes only that
worker's processes. It does not establish that a file is password protected.

Excel opens files read-only with macros disabled, link updates disabled, and
manual calculation. Legacy workbooks are saved as XLSX without automatic repair.
Existing XLSX files are copied unchanged. A fresh Excel instance then reopens
the output with alerts enabled and explicit empty passwords. Acceptance means
that this recorded COM reopen completed in this Windows/Excel environment; it
does not guarantee prompt-free opening on every installation.

SaveAs can change unsupported features, including VBA and Excel 4 macro sheets.
Disabling link updates also affects chart import for some Lotus formats. These
are normalization limitations, not Excel Shrink results. Sheet protection alone
does not make a workbook unreadable. Manual recovery comes after the automatic
pass; record the original hash, action, and new output before including a repair.

## Release and optional later measurements

Create a local archive only after reviewing the inventory and exceptions:

```powershell
python scripts/enron.py package F:\corpora\releases\enron-xlsx-v1.zip
python scripts/enron.py benchmark F:\repos\excel-shrink-rust\target\release\excel-shrink-cli.exe F:\corpora\enron-benchmarks\run-01
python scripts/enron.py excel-benchmark F:\corpora\enron-benchmarks\excel-01
```

Packaging includes source mappings, exceptions, and checksums. It does not upload
anything or assign a new license to the source material. The benchmark requires
the CLI version with `--batch-report`. It invokes one prebuilt process for the
entire corpus for each profile, with three repetitions and rotating profile order.
Output hashes are computed outside the timed run. `variants.jsonl` records the
inputs and outputs for subsequent visual comparison. The `too-aggressive` profile
is an intentionally unsafe comparison, not an equivalence claim.

Report original-to-XLSX construction time separately from shrink time. A fairer
performance comparison also measures Excel opening and saving the frozen XLSX
inputs. Conversion verification, downloads, manual work, builds, and outer ZIP
compression must not be folded into that comparison. Existing visual-validation
tools still need an adapter for these saved variants before they can reuse them.

Implementation lives in `src/excel_shrink_research/enron/`; tests live in `tests/`.
Existing paper sources and earlier experiments remain at their established paths.

```powershell
$env:PYTHONPATH = 'src'
python -m unittest discover -s tests -p 'test_enron_*.py'
```

### Embedded workbooks in Word documents

Some `new-only.csv` candidates are Word documents containing Excel objects. Their
whole-document hashes cannot establish whether those workbooks are additional.
On Windows, with `pywin32==311` and `olefile==0.47` installed, extract and
compare them without Office:

```powershell
$env:PYTHONPATH = 'src'
python -m excel_shrink_research.enron.embedded_workbooks --candidates F:\corpora\enron-v2\comparison\new-only.csv --raw F:\corpora\enron-v2\raw --old-zip F:\corpora\enron.zip --old-index F:\corpora\enron-v2\comparison\old-files.jsonl --output F:\corpora\enron-v2\embedded-workbooks
```

`comparison.csv` maps each extracted `.xls` to its document hash, object path and
old-corpus matches. `source-documents.json` preserves all original occurrences.
Directory timestamps are not interpreted; invalid legacy dates cannot block
reading. Extraction copies the complete embedded storage and verifies its named streams
and storage classes. It does not convert to XLSX or verify that Excel can open it.

`same_storage` means matching named stream bytes and storage classes, independently
of OLE sector allocation. `same_workbook_stream` matches only the Excel `Book` or
`Workbook` stream; other content may differ. `no_workbook_stream_match` is a byte
comparison, not a claim that the workbook is semantically new. Check `issues.json`
and `status.json` before interpreting missing matches. This scan covers `.doc`
candidates; it does not survey every possible embedding format in the release.

To inspect the remaining candidates by content, use
`python -m excel_shrink_research.enron.container_embeddings --help`.
It reads nested OLE storages, PowerPoint compressed objects, native packages and
MIME parts without Office. `inventory.csv` identifies actual content types;
`comparison.csv` compares extracted Excel streams with the old index and the Word
extraction. PowerPoint results include stored historical revisions, not just
objects on current slides. Read `issues.json` for incomplete coverage.

### Targeted PST source comparison

`python -m excel_shrink_research.enron.pst_probe --help` runs the isolated Saibi
and Sanders probe. It verifies pinned PST and Node downloads, installs a pinned
reader with an npm lockfile, and starts attachment extraction after the supplied
ready receipt exists. It never starts Office. The ready receipt and source
manifest record the scripts and downloads used.

The probe writes `results.csv` for the 131 unresolved originals and `matches.csv`
with attachment names and message identifiers. All attachment payloads are
retained by SHA-256; reader exceptions remain in `reader-issues.jsonl`.
The selected PST release was modified by its publisher and cannot establish
completeness of the historical Nuix source. Matching a Workbook stream alone
also does not establish complete-workbook equivalence.

### Complete pinned PST publication

`python -m excel_shrink_research.enron.pst_download --help` downloads all 171 PSTs
from the reviewed revision. Interrupted files resume with HTTP ranges; every
final file must pass its published size and SHA-256. The two probe PSTs are reused.
`download-complete.json` confirms this publication was downloaded in full.

Before starting, `scripts/process_full_pst.py --prepare-only` freezes the follow-up
scripts and saves `pipeline-ready.json`; supply the same output and probe reuse
roots as the downloader. Processing follows automatically after download. Per-PST
checkpoints retain completed work across restarts. Every readable attachment is
hashed. Candidate spreadsheets, containers, text and exact search targets are
saved separately; other payloads remain in the unchanged source PSTs. The index
records whether a separate payload was retained. Office is not started.

This is a modified publication with documented removals. Neither a full download
nor a negative match establishes completeness of the original Enron or Nuix data.
