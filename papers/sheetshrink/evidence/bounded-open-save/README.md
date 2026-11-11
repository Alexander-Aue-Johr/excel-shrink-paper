# Bounded opening and saving results

The six-core laptop ran the independent series from November 10 to November 11, 2026. It has a Ryzen 5 4500U, six logical processors, 8 GiB installed memory, and Windows 11. Each process tree had a 4 GiB commit limit and a 900-second timeout. No free-RAM reserve was configured. `trials.jsonl.gz` contains the 288 trial receipts with local command paths removed. `summary.json` and both figures summarize these same trials. The launcher and plot source are in `scripts/open_save_benchmark` at the repository root.

Each tool has 48 trials across 16 workbook and library combinations. The three workbooks were selected as stress cases. The five XLSX endpoints are openpyxl, pandas with openpyxl export, Excel COM, R openxlsx, and R readxl with writexl. The Migration workbook also has a pandas CSV export endpoint. Each combination has three repetitions.

| Input variant | Completed | Commit limit |
| --- | ---: | ---: |
| Original | 33 | 15 |
| SheetShrink | 48 | 0 |
| NXPowerLite | 39 | 9 |
| XLStylesTool | 33 | 15 |
| Check Performance | 39 | 9 |
| Inquire | 36 | 12 |

SheetShrink revision `765c480` ran on each original with six Rayon and six writer workers before the library opened its fresh output. Its pipeline clock runs from CLI launch through the library process exit. Direct-original elapsed time includes process startup and exit. The phase receipts retain the measured open and save times and the Excel shutdown time. Other tools use prepared workbooks, with historical cleanup measurements labeled by their provenance where available. Those entries support downstream library comparisons.

Across the 11 pairs that completed all three repetitions in both conditions, SheetShrink's median full pipeline was faster in ten. Excel on the municipal finance workbook took 42.46 seconds directly and 69.76 seconds with cleanup. Migration openpyxl took 111.30 and 9.67 seconds respectively. Its maximum job commit over the three repetitions fell from approximately 2,881 MiB to 577 MiB. The job commit peak covers the entire process tree. Resident memory samples and phase-specific private byte samples are in the receipts.

All 48 additional second SheetShrink passes completed. Fifteen QT/RR outputs matched their first output byte for byte. The 18 Migration and 15 municipal finance second outputs had changed archive hashes. Further package inspection is required to attribute the changes. Second-pass timings and memory are recorded separately from the main pipeline.
`publication-projection.json` records the metadata projection that replaces the private producer-host name in 63 cleanup-provenance fields. All other trial fields and measurements are preserved.
