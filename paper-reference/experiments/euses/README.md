# EUSES spreadsheet inputs

Sources are pinned to the [original Zenodo release](https://zenodo.org/records/581673),
DOI `10.5281/zenodo.581673`: eleven category ZIPs and a README. Published MD5
checksums and locally computed SHA-256 values verify the downloads.

The release describes 5,607 files, including 4,499 processed files; these counts
differ from the 4,498 workbooks in the original paper. Actual archive counts are
recorded rather than forced to match either description.

```powershell
$env:PYTHONPATH = 'src'
python -m excel_shrink_research.euses download --root F:\corpora\euses
python -m excel_shrink_research.euses assemble --root F:\corpora\euses
```

Both steps resume verified outputs and run without Excel. `raw/` retains the
original sources. `input-manifest.jsonl` maps every original occurrence to the
derived runner ZIP, including `processed`, `bad` and `duplicates`. Equal original
SHA-256 values share one preparation job. Files with unusual suffixes receive
an additional `.xls` suffix only in the runner ZIP; their bytes are unchanged.

`runner-archive.json` records source and output hashes, unique job counts and
raw occurrence counts. Download and assembly progress are separate JSON status
files. The preparation queue starts the normal Rust `--prepare-only` runner
after Enron retries finish, with separate output under `.tmp/regression/euses`.
Conversion failures remain recorded exceptions.

## Reusable Excel conversion

The research converter accepts a directory or a ZIP of workbook inputs:

```powershell
$env:PYTHONPATH = 'src'
python -m excel_shrink_research.convert_workbooks F:\corpora\euses\corpus\euses-original-inputs.zip F:\corpora\euses\xlsx-preparation
```

Use `--members-file selected.json` for a JSON array of exact member names. Repeat
the same command to resume accepted outputs whose source, output and converter
hashes still match. `stop.request` stops between files. Timeouts are per file;
only the worker's private Excel process is terminated. Macros, events and link
updates are disabled. No Office security settings are changed.

A short explicit test password prevents omitted-password prompts. Only an
explicit password-related Excel error is classified as `password_rejected`;
other opening failures remain unresolved. A password to modify may still allow
read-only conversion and does not prove an opening password is required.

A shared-workbook save failure triggers one retry with `SaveAs(..., xlExclusive)`.
The record marks removal of the legacy shared revision history. Originals are
unchanged. Accepted XLSX packages pass a CRC check and reopen in a fresh Excel
instance with alerts enabled, without requesting repair. Sheet names and order
must agree before and after reopening. This does not establish visual equivalence.

`conversions.jsonl` maps inputs and outputs with hashes, normalization steps and
phase results. `records/` and `attempts/` preserve individual failures and logs.
For a release, include these records, the converter package, its PowerShell and
C# helpers, and the recorded Excel version. This workflow is separate from the
Rust regression runner; it does not rewrite that runner's cached results.

The seven EUSES revision-history cases are rerun with this converter under
`F:/corpora/euses/conversion-retry`, rather than imported as untracked temporary files.

## Recovery of 30 File Block failures

The follow-up opened legacy inputs in Protected View and used Windows UI
Automation to invoke Enable Editing in the owned Excel process. Generic SaveAs
errors persisted until the native Save As dialog revealed a separate restriction:
Trust Center also blocked saving modern XLSX workbooks.

At the user's request, the Excel 2007-and-later workbook/template restriction was
cleared through the UI. Excel coupled Open and Save for this row; both were
cleared, while all other File Block rows were verified unchanged. COM saving then
worked. All 30 selected files passed XLSX saving and fresh-process reopening.
This establishes conversion acceptance, not visual equivalence.

The run, frozen code, selected input names, before/after settings and diagnostic
evidence are under `F:/corpora/euses/after-settings-check`. Its
`reproduction/README.md` gives the exact command and explains the distinction
between the UI-assisted source opening and the unattended output check.
The converter itself does not change global Trust Center settings.

## Dialog review of remaining opening failures

The separate recovery review under F:/corpora/euses/dialog-review records native
Open and Repair interactions on copies, exact dialog text, COM errors and a short
interpretation for each of 37 unresolved opening cases. Its method.md explains
selection, consented repairs, timeouts and the distinction between recovery and
verified conversion. Recovered workbooks are kept separate for visual review.
