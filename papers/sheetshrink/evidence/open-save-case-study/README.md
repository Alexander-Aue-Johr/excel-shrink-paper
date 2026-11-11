# Opening and memory case-study tables

The transformation table and the continued resource table are generated with
`python scripts/open_save_benchmark/prepare_case_study_tables.py`.

`sizes.json` records SHA-256-verified local manifest inputs and prepared outputs.
ZIP directory metadata supplies the sum of all uncompressed member sizes.
Duplicate member names were rejected. The manifest is the retained
open-save-memory-20261110 manifest. Original hashes match the bounded trial
receipts. The Enron original is reserved as a fourth case. Comparator outputs
for that additional case remain pending in this table.

Fresh SheetShrink ZIP sizes come from the input_bytes field of the separate
second-pass receipt, which identifies the first cleanup output before library
processing. All first-output sizes agree within each case. Their archive hashes
vary across repetitions. Uncompressed sizes of previously prepared SheetShrink
outputs are retained in the inventory, while the main table marks the fresh
output's uncompressed size and visual join pending until equivalence is verified.

`visual-summary.json` identifies twelve exact-output native attempts, including
two completed Municipal Finance reports with limitations. Original and output
hashes match the same variants in the size manifest. The two diagnostic extracts
retain per-region pixel counts, sheet names, capture ranges and geometry flags,
with the full source-report SHA-256. D sums changed_pixels from comparison 0's
finding diagnostics once. Its unit is pixel observations across selected captures.
The original/candidate states are paired native rebuilds under capture policy
v12. N remains pending because a qualified denominator over all compared
surfaces has not been extracted. Differences remain observations requiring
attribution. Memory-limited and failed attempts remain explicit.

`resource-summary.json` aggregates all 288 existing trial receipts from
`../bounded-open-save/trials.jsonl.gz`, including all six variants and sixteen
workbook/library endpoints. Each endpoint has three attempts. Timing columns
use completed attempts. Memory maxima include failed attempts. RSS columns use
sampled aggregate process-tree working sets by phase. Commit uses the Windows
job peak counter. The original and prepared-output total is wall_seconds.
Fresh SheetShrink uses its directly measured pipeline_seconds. Historic cleanup
times from other workflows are excluded from these end-to-end columns.

CPU medians, trial statuses and elapsed ranges are retained in the JSON.
The underlying receipts preserve individual phase values and all repetitions.
An M@time value means time until a memory-limit stop. P marks pending evidence.
The further Enron lane remains pending. The larger-memory-host follow-up is
recorded separately in `../bounded-open-save-followup` and its manuscript table.

Color is redundant with numbers and textual status. Pale blue accompanies
positive size reduction or a faster matched SheetShrink total. Pale orange
marks differences, stops or slower matched totals. Black text is retained.
The tables use standard LaTeX tabular and xcolor already loaded by TikZ.
