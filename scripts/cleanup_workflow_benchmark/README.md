# Paired cleanup workflow measurements

`measure_sheetshrink.py` executes the ordinary SheetShrink CLI on every original
in a frozen selection. It uses the versioned `process_metrics.py` helper from
the SheetShrink repository. The executable and helper should be copied into the
experiment before execution. Record their SHA-256 values and the source revision.
Python requires `psutil`, with NumPy and Matplotlib for plotting.

```powershell
python measure_sheetshrink.py --selection selection.json --binary bin/sheetshrink.exe --metrics-helper environment/process_metrics.py --output experiment --repetitions 9
```

The selection is a JSON list with `corpus`, `rank`, `file`, `input_path` and
`sha256`. Every original is hash-verified before execution. One retained warm-up
round is excluded from the nine measured rounds, with reproducible shuffled order.
Each trial launches one fresh CLI process with default cleanup, one active
workbook, one Rayon worker and one writer. Wall time covers process creation
through exit. Hashing is outside that interval. Filesystem caches are warm.
Output packages occupy one directory per input and are overwritten by later
trials. Per-trial receipts preserve output hashes, elapsed time, user/kernel CPU,
observed OS peak working set, sampled private memory, timeout status and nearby
native processes. Sampling uses the frozen helper's 10 ms loop.

The runner resumes completed trials only under an identical protocol. A failed
trial remains a failure. Stop and investigate incomplete repetitions before
using a workbook's timing summary. Original workbooks are never opened through
Pandas, OpenPyXL or Excel by these scripts. In particular, the original
Wanderungen Pandas opening/export experiment is prohibited.

## Join and plot retained observations

`export_pairs.py` checks source SHA-256 and matching host identity. Its comparator
input is the native receipt index with an explicit `host_id`. A successful retry
cannot replace a failed primary observation. Protected-part audit exclusions
remain outside the timing plot and inside outcome accounting.

```powershell
python export_pairs.py --sheetshrink experiment/SheetShrink-Top49-Medians.json --trials experiment/trials.jsonl --comparator NXPowerLite-Top49-paired-receipts.json --output paired-cleanup-workflows.csv
python plot_workflow_comparison.py --pairs paired-cleanup-workflows.csv --output paired-cleanup-workflows.pdf
```

The chart reads CSV files only. It plots raw seconds and workbook-level ratios
of primary comparator time to median SheetShrink time, retaining no-change cases.
Points and interquartile ranges describe the selected workbooks. They do not
estimate repeated-trial uncertainty for a single-observation comparator.
NXPowerLite's wall interval includes GUI orchestration, native confirmation and
initial output verification. SheetShrink's interval is CLI process lifetime.
These boundaries and different cleanup policies must accompany interpretation.
Visual validation is a separate eligibility result.

`import_xlstyles.py` reads the earlier conversion database in read-only mode.
It requires the first conversion attempt to contain both native confirmation
and a per-attempt resource receipt. Hash-matched outputs pass a protected-part
audit before entering the figure. The retained Top49 lane contributes 122 files.
The other 25 files remain excluded from timing because their complete primary
receipts are unavailable. Legacy resources stored only on the mutable job row
are insufficient for this import.

```powershell
python import_xlstyles.py --database destatis.sqlite --selection selection.json --host-id LAPTOP-H86KBIH3 --output XLStylesTool-Top49-paired-receipts.json
python export_pairs.py --sheetshrink experiment/SheetShrink-Top49-Medians.json --trials experiment/trials.jsonl --comparator XLStylesTool-Top49-paired-receipts.json --output top49-xlstyles-pairs.csv
python plot_workflow_comparison.py --pairs paired-cleanup-workflows.csv top49-xlstyles-pairs.csv --common-inputs --output paired-cleanup-workflows.pdf
```

Tool names appear directly in the legend and beside every corpus group. Colors
and marker shapes identify NXPowerLite, XLStylesTool and SheetShrink. The joint
figure uses `--common-inputs`, retaining 110 shared originals across the three
tools, with 21 Destatis, 41 Enron and 48 Data.gov workbooks. The larger individual
receipt sets remain available. The same repeated SheetShrink baseline is paired
with both external tools for each shared input.

For Check Performance, Inquire and XLStylesTool on other hosts, generate a paired
CSV with the same columns and match a SheetShrink baseline measured on that
host. Pass multiple CSV paths after `--pairs`. Each plotted group retains the
host, tool and corpus. Preserve required input opening, saving and other phases
in each lane's protocol. Add machine metadata, raw observations and exclusions
with the new lane. Do not combine unmatched hosts or fabricate missing data.

The two-panel design follows the presentation of absolute cost alongside ratios
and explicitly defined distribution summaries in Hoefler and Belli, SC 2015,
Figure 7, DOI `10.1145/2807591.2807644`. Position on a common axis, distinct marker
shapes, restrained color and logarithmic axes expose the full range without
compressing small observations into zero-width bars. The historical
`plot_retained_wanderungen.py` remains the separate opening-plus-saving figure.
Cleanup workflow measurements do not supply new opening-time observations.

The optional `--receipts` argument adds selection-outcome bars from complete
compressed per-tool receipt inventories. Counts distinguish plotted pairs,
other eligible inputs, failed conversions, protected-payload exclusions and
incomplete timing receipts. The current manuscript command adds
`--receipts papers/sheetshrink/evidence/top49-nxpowerlite-receipts.json.gz papers/sheetshrink/evidence/top49-xlstyles-receipts.json.gz`.

For the HP Inquire lane, run `export_inquire.py` first, then add
`papers/sheetshrink/evidence/top49-inquire-pairs.csv` to `--pairs` and
`papers/sheetshrink/evidence/top49-inquire-receipts.json.gz` to `--receipts`.
The Inquire lane uses the frozen shared 110-input cohort and retains its own
host, repetition counts, workflow endpoints and historical storage conditions.
Partial attempts have a separate status category. Storage markers distinguish
NVMe, USB and incompletely recorded historical roots.

For Check Performance, run `export_check_performance.py`. Add
`top49-check-performance-pairs.csv` and `top49-check-performance-receipts.json.gz`
to the respective plot inputs. Use `--cohort-pairs` with the original NXPowerLite
and XLStylesTool pair CSVs instead of `--common-inputs` to freeze the original
110-input cohort and retain each tool's available successful subset.

The manuscript uses `--single-column` for a corpus-grouped two-panel figure.
It retains the complete-pair counts in row labels and historical Inquire
storage markers. Selection outcome totals remain in the manuscript setup and
receipt inventories. Omit the flag to reproduce the wider outcome-panel view.
