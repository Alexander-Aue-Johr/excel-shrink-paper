# Destatis cleanup comparison preparation

## Prepared study

The local study root is `C:/SheetShrinkStudies/baseline-preparation-20261109`.
It contains verified input copies, current SheetShrink outputs, per-file receipts,
per-sheet statistics, software sources, build logs and installers.
The preparation uses the normal SheetShrink cleanup policy with one Rayon
thread, one writer and one active workbook. Attribution statistics are enabled.
The competitor cleanup and independent native comparison stages await the
separate benchmark approval.

The source corpus is available locally in
`C:/SheetShrinkStudies/distributed-six-20261109/susann-laptop/artifacts/destatis`.
All 1,182 historical workbook names matched that directory. The source archive
SHA-256 was verified as
`f0b39b1a30778a80227b01ae64109db964c6f71f83e6ccb3097e2f9e26f4ac45`.
The versioned profile identifies each input by name, size and SHA-256.

The fetched software tip remains `origin/codex/datagov-validation-watch` at
`a7eef1ae8d9a2c2ed03720eb0433ce2a8bf7e216`. A fresh release build of the
`sheetshrink` application produced executable SHA-256
`d7bddcd49c7df9bda92038f106fa2772ab67c10707a41583d4c95c1b01290f47`.
The build receipt records Cargo.lock identity, toolchain and selection-script
identity. Source export and compilation use the study's software directory.

## Selection and cleanup observations

The selection alternates the historical absolute and relative reduction
rankings until 100 distinct files are selected. Historical values are summed
worksheet XML sizes from `analysis_output.csv`, rather than compressed file
sizes. Wanderungen was already included at selection position 5. The rule
would additionally include it if absent from the initial 100.

This is a purposive stress selection. Population-level reduction estimates need
their own sampling design. Historical cleanup rules and source versions differ
from the current implementation, so historical values serve only to select leads.

| Measure | Observed value |
| --- | --- |
| Selected workbooks | 100 |
| CLI processed outcomes | 100 |
| Input copies checked against source SHA-256 | 100 |
| Output ZIP archives passing CRC checks | 100 |
| Original compressed bytes | 495,908,582 |
| Output compressed bytes | 246,499,411 |
| Aggregate compressed reduction | 50.2934% |
| Smaller outputs | 99 |
| Larger outputs | 1 |
| Sum of individual observed wall times | 258.3735 seconds |

These times include process startup and analysis statistics while other setup
work occurred on the same host. They are preparation observations. Matched
performance measurements and native preservation evaluation remain pending.

The one larger output is
`statistischer-bericht-vj-schulden-oeffentl-gesamthaushalt-2140520243215.xlsx`,
which changes from 4,323,548 to 4,371,485 bytes, an increase of 47,937 bytes
(1.1087%). Its outcome remains in the complete ranking.

## Provisional Destatis panel

Six candidates provide a compact initial Destatis panel. Reserving six inputs
each for Enron and Data.gov would produce an 18-input balanced panel, within one
NXPowerLite trial batch. The other two populations still require input selection.
The actual panel size can grow within the documented trial period.

| Workbook | Original bytes | SheetShrink bytes | Reduction |
| --- | ---: | ---: | ---: |
| `statistischer-bericht-rechnungsergebnis-kernhaushalt-gemeinden-2140331217005.xlsx` | 89,367,137 | 4,602,649 | 94.8497% |
| `verkehrsunfaelle-zeitreihen-xlsx-5462403.xlsx` | 37,207,883 | 2,608,628 | 92.9890% |
| `wanderungen-2010120217005.xlsx` | 16,926,818 | 1,579,059 | 90.6713% |
| `statistischer-bericht-stromsteuerstatistik-2140970237005.xlsx` | 1,432,587 | 145,225 | 89.8627% |
| `kostenstruktur-aerzte-2020161199005.xlsx` | 15,237,103 | 3,318,914 | 78.2182% |
| `statistischer-bericht-zeitverwendung-priv-haushalte-5639102229005.xlsx` | 3,840,894 | 715,062 | 81.3829% |

`benchmark-preparation-20261109/provisional-panel.json` binds these provisional
choices to source and output hashes before comparator processing. The two full
CSV rankings retain all 100 outcomes. The shortlist reduces repeated publication
families and includes the required hidden-zero case.

## Wanderungen hidden-zero case

The identified sheet is `Tabelle 3.4 `, including a trailing space in its name,
resolved to `xl/worksheets/sheet31.xml`. The original contains 361,086 cell
records. Of those, 343,833 are numeric zero constants in hidden columns L:XFD.
They span rows 11–30 and 32, with 16,373 such constants in each of 21 rows.

The current output contains 787 cell records and zero numeric zero constants
on that sheet. Its worksheet XML changes from 11,725,919 to 32,423 bytes.
The CLI's `UnreferencedHiddenContentCells` statistic independently records
343,833 removals. The hidden column interval remains stored.
`benchmark-preparation-20261109/wanderungen-zero-tail.json` retains the direct
source/output XML counts. These are package observations. Native visibility
and preservation outcomes still require the approved comparison stage.

The inspection uses ZIP access and streaming XML parsing. The original-workbook
Pandas workload remains prohibited. The prepared workflow imports no Pandas.

## Installed tools and availability

| Tool | Verified local state | Benchmark state |
| --- | --- | --- |
| NXPowerLite Desktop | Version 11.0.1, valid Neuxpower-signed MSI, installed under `C:/Program Files/NXPowerLite`. Trial dialog shows 14 days remaining. | Empty file list, default Balanced profile and create-copy mode observed. |
| XLStyles Tool UWP | Microsoft Store package `6764XLGeekCoder.XLStylesTool`, version 3.0.1.0, publisher XLGeekCoder. | Launch succeeds, no file selected and Process File disabled. |
| Excel Check Performance | Excel x64 `16.0.20527.20150`; Review tab exposes `CheckPerformance` / Leistung prüfen. | Command availability checked on a new blank workbook. |
| Inquire | Absent from installed Office components and the inspected COM Add-ins list. Installed Office product is `O365HomePremRetail`. | Requires a supported Enterprise/Professional Office installation. |

The NXPowerLite installer SHA-256 is
`eecfcb16282465bd6c4c38094e1514f0c95541643c73ecdc4a1d6828c4e4c322`.
The standalone installation omits Explorer, Outlook, Word, PowerPoint and Excel
integration and a desktop shortcut. The final elevated installation returned 0.
Earlier non-elevated attempts returned 1603 and remain in the local logs.
[NXPowerLite deployment guide](https://support.neuxpower.com/hc/en-us/articles/115002535913-NXPowerLite-Deployment-Guide-Windows-only)

The current FAQ specifies a 14-day trial with at most 20 files processed at the
same time. It describes a batch limit. It does not state a lifetime allowance of
20 files. The observed UI confirms the available trial period.
[NXPowerLite FAQ](https://support.neuxpower.com/hc/en-us/articles/360026595053-NXPowerLite-Desktop-FAQ)

Microsoft limits Inquire to Enterprise and equivalent supported Office editions.
Its absence here is a concrete installation/edition constraint. An appropriate
installation on another benchmark host can supply this comparison lane.
[Microsoft Inquire availability and activation](https://support.microsoft.com/en-us/excel/turn-on-the-inquire-add-in)

## Next benchmark freeze

- [x] Locate and verify Destatis source identities.
- [x] Select and clean 100 current-version working copies.
- [x] Retain full rankings, per-file outcomes and the Wanderungen zero evidence.
- [x] Install and launch NXPowerLite and XLStylesTool.
- [x] Verify Check Performance command availability and Inquire's local blocker.
- [ ] Select the Enron and Data.gov portions and freeze the final panel.
- [ ] Provide an Inquire-capable host and record its Excel version.
- [ ] Approve the comparator run and native validation scope.
- [ ] Record exact comparator profiles before processing. Separate a documented
  default profile from any profile restricted to cell cleanup. Image resampling,
  hidden-content removal and object flattening settings need explicit accounting.
- [ ] Run competitors against fresh source copies and retain source/output hashes,
  size changes, phase timings, failures and unchanged outcomes.
- [ ] Evaluate all prepared outputs under a matched native observation protocol,
  with save/rebuild controls and independent reference areas.

Competitor behavior on the hidden-zero case and possible visual regressions
remain hypotheses for that stage. The current preparation contributes no
competitive ranking or visual-preservation verdict.

## Reproduction

Use `scripts/benchmark_preparation/prepare_destatis_subset.py` with a fresh study
output root, the pinned Destatis profile, historical CSV and built executable.
It verifies inputs, stages copies and runs the production entry point. Its
exclusive JSONL journal preserves completed receipts if a later input fails.
The source and output ranking hashes identify the current preparation even
when historical filenames occur in other corpus versions.

The executed helper is retained locally as
`software/prepare_destatis_subset.executed.py`. The tracked helper additionally
rejects an existing attempt journal before staging, preventing manifest changes
on reuse of a completed study root. The build receipt distinguishes the executed
helper hash from the current reproduction helper hash.
