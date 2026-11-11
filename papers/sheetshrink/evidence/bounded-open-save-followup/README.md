# Bounded opening and saving follow-up

This separate host lane repeats the five workbook/library combinations that
reached the laptop study's 4 GiB job-commit ceiling. Each combination has three
Original trials and three fresh SheetShrink-plus-library trials under both
4 GiB and 24 GiB aggregate process-tree commit limits. All sixty planned trials
have a terminal receipt. The original laptop measurements retain their own
host, worker configuration and outcomes.

`protocol.json` records source identities, software versions, binary identity,
configuration and outcome counts. `trials.jsonl.gz` contains path-free individual
receipts, including failures, phase samples, working sets, process-tree commit,
CPU time, saved-output identities and the separate second cleanup. `summary.json`
keeps completed-endpoint medians and ranges separate from failure termination
times. The deterministic gzip file's SHA-256 is recorded in the protocol.
One worker traceback remains in the private raw receipt. Its public projection
retains the error type and outcome, with an explicit private-traceback flag.

All thirty fresh-cleanup pipelines completed. The Original completed nine
trials and reached a configured memory limit in twenty-one. At 4 GiB, all
fifteen Originals reached the limit. At 24 GiB, the R/readxl, Finance/openpyxl
and QT/RR/openpyxl Originals completed all nine trials. The two Migration/pandas
endpoints reached the limit in all six attempts. A stopped allocation is a
censored endpoint and receives no completed-runtime speedup.

Elapsed time covers process creation through process exit. The SheetShrink lane
includes fresh cleanup and the same library round trip. Individual open/save
and cleanup times remain separate fields. Repeated-cleanup measurements run
after the timed endpoint. They are excluded from its elapsed time. Their byte
identity flags describe recorded ZIP outputs and do not establish general
cleanup idempotence or native visual preservation of a second cleanup.

The Windows host has an AMD Ryzen 7 7800X3D, sixteen logical processors and
102,066,212,864 OS-reported physical bytes. Fresh cleanup uses sixteen Rayon and
sixteen writer workers with one active workbook. The controller checks an
8 GiB system reserve for both physical RAM and commit headroom and contains
each trial tree in its own Windows job before execution. Its timeout is
900 seconds. The sampled minima remained above the reserve in all sixty trials.
Native visual comparisons, the paired memchr experiment and read-only evidence
analysis ran concurrently. Filesystem caches were retained. Report these as
selected-case descriptive measurements under that recorded context.

## Fresh candidate identity controls

After all timed trials, a further fresh output was generated for each input
with the same binary and cleanup options. All uncompressed member names, byte
lengths and SHA-256 values match the corresponding retained native candidate.
`fresh-native-joins.json` preserves the controls and their resource guards.
Migration's native observation uses a documented normalized baseline. Finance
uses a raw selected original. QT/RR's retained selected plan has zero captures.
These coverage qualifications remain when associating a representative output
with the resource study. The timed fresh outputs' ZIP hashes are recorded, while
their complete per-trial member inventories were not retained.

## Reproduction

`measured-code` is a frozen source snapshot. Its original bytes are preserved
for comparison with the start-of-run Python script hashes. It includes the
owned Windows-job helper, library workers, controller, pipeline and auxiliary
scripts. Two pure guard tests passed in this snapshot. A separate 256 MiB
allocation probe terminated at its job memory boundary before the study.

Use the exact original files identified by the protocol and configure a manifest
with only the intended case. Its Original and SheetShrink entries may both
identify the original input when `--sheetshrink-binary` is supplied. The fresh
pipeline creates its own cleaned workbook. Keep the original SHA-256 check.

Prepare each result root's `software/packages` with the recorded Python packages,
`software/R` with the R distribution and `software/R-library` with the recorded
R packages. The Windows owned-job helper is supplied with `--helper-directory`.
The controller is the same interface as the original bounded-open-save artifact.

```text
python measured-code/run_bounded.py --root NEW_RESULT_ROOT --manifest CASE_MANIFEST.json --helper-directory measured-code --libraries LIBRARY --repetitions 3 --limit-mib 4096 --reserve-mib 8192 --timeout 900 --sheetshrink-binary FROZEN_SHEETSHRINK.exe
```

Run the five combinations in protocol order, with separate roots, then repeat
with `--limit-mib 24576`. The combinations are Migration/pandas-to-XLSX,
Migration/pandas-to-CSV, Migration/readxl-with-writexl, Finance/openpyxl and
QT/RR/openpyxl. Their CLI library names are `pandas`, `pandas-csv`, `r-readxl`
and `openpyxl`. The controller retains Original then fresh SheetShrink order
within each repetition. A reserve of zero disables preflight reserve waiting
and should be treated as a different safety configuration.
