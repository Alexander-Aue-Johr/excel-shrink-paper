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
