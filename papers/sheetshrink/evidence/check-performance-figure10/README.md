# Check Performance primary workflows

The 110-input cohort matches the original SHA-256 intersection of the
NXPowerLite and XLStylesTool Figure 10 inputs. Primary outcomes are 102 complete,
one partial and seven failed. Complete workflows include 86 confirmed Optimize
All actions and sixteen native no-optimization-needed outcomes. Diagnostic
retries are excluded from ratios.

`phase-comparison.csv` retains combined control-save/optimization/final-save
phase times, native outcomes, output identities, owned Excel CPU and sampled
working-set/private-byte peaks. `preflight.json`, `machine.json` and
`sheetshrink-baseline-audit.json` identify software, host and reused baselines.
`verification.json` records the final original/output hash and SQLite checks.
Full attempt receipts and the SQLite ledger remain in the local
`C:/Users/saue1/CP-Figure10-20261111` return package.

Workflow timing covers native-wrapper launch through exit, including control
saves and artifact creation. Staging and independent hashing are outside it.
Optimization and associated saves are recorded together. Memory is sampled
Excel-only usage, with a target interval of 250 ms. The median of successful
attempts' mean intervals is approximately 251 ms.

Nine historical same-host Sheet Shrink trials per original provide medians.
All 990 trials list a concurrent Check Performance runner process. Process lists
provide no activity measurement for that runner. These ratios characterize
recorded workflows under uncertain background contention and differing endpoints.
They do not isolate algorithm execution or establish visual equality.

Regenerate the plotting inputs with

```powershell
python scripts/cleanup_workflow_benchmark/export_check_performance.py
```

For the four-tool plot, use `--cohort-pairs` with the original NXPowerLite and
XLStylesTool pair CSVs. This freezes their 110-input intersection while preserving
the available successful Check Performance and Inquire subsets. Using a new
four-tool intersection would discard documented failures from other tool groups.
