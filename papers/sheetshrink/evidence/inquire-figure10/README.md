# Inquire workflow comparison

The frozen 110-input cohort matches the original SHA-256 intersection of the
NXPowerLite and XLStylesTool Figure 10 inputs. Primary outcomes are 90 complete,
six partial and fourteen failed. Diagnostic retries are excluded from ratios.

`comparison.csv` joins historical Inquire workflow observations to repeated
same-host Sheet Shrink measurements. `trials.csv.gz` retains the 916 Sheet Shrink
trials, including 110 warm-ups. The exporter recomputes every median and checks
trial counts, original cohort identities, successful trial states and repeated
content identities before plotting. Raw phase boundaries and storage attribution
are retained in `inquire-primary-summary.csv`. Full databases, process samples,
outputs and diagnostic retries are in the archived return identified by
`provenance.json`.

Historical Inquire work directories used local NVMe for 41 complete pairs,
USB SSD for 48, and incompletely recorded storage for one. Sheet Shrink used
local NVMe for all trials. Ratios characterize recorded workflows with different
endpoints and storage. Historical hardware/version attribution partly relies on
execution-session provenance and a later inventory. These observations do not
isolate algorithm execution or establish visual equality.

Regenerate the Inquire plotting inputs with

```powershell
python scripts/cleanup_workflow_benchmark/export_inquire.py
```

Add `top49-inquire-pairs.csv` to `--pairs` and
`top49-inquire-receipts.json.gz` to `--receipts` in the Figure 10 plot command.
