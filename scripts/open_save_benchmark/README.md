# Bounded opening and saving study

The selected cases are the hidden-zero-tail Wanderungen workbook, the large Destatis municipal finance workbook, and the data-heavy Data.gov QT/RR workbook. Their original identities and every supplied output used in a trial are checked with SHA-256 before any trial starts.

Each SheetShrink trial starts from the verified original and runs the frozen SheetShrink CLI immediately before the selected library. Previously prepared SheetShrink outputs are excluded from input verification and processing for this mode. The CLI uses the host's logical-processor count for both Rayon workers and writer workers. LAPTOP-H86KBIH3 exposes six logical processors. One workbook is active in each serial trial.

The first cleanup and library execution share the same bounded Windows job. The receipt records directly measured pipeline time, cleanup time, native SheetShrink memory peaks, phase-level process-tree memory samples, and per-process sampled memory peaks. Pipeline time includes interpreter startup and application shutdown. Excel's separate open, save and close/quit durations remain available. Timing excludes output hashing and temporary-file removal.

After the library trial, the freshly cleaned workbook is processed by SheetShrink a second time in a separate bounded job. Its duration, CPU, memory, exit status, output size and byte-identity comparison are retained under `repeat_sheetshrink`. This second run is excluded from the first pipeline's measured duration and resource counters. Byte identity establishes an unchanged archive for that trial. A changed archive requires inspection before claiming a semantic or visual change.

The current independent launcher writes to `standalone-pipeline-4096mib`. Earlier prepared-output series retain their existing directories. The new SheetShrink measurements use six workers and cannot be substituted by the earlier single-worker cleanup timings.

Excel's SaveAs duration and its subsequent workbook-close and application-quit duration are recorded separately. Application shutdown can take approximately one minute on this host even for a one-cell workbook. The total process lifetime remains in the receipt. The standalone series uses this timing definition and is kept separate from preliminary trials that included closing in the save timer.

The normal entry point is `launch.ps1 -StudyRoot <study directory>`. It runs three repetitions of openpyxl, pandas, Excel COM, R openxlsx, and R readxl with writexl. Each trial opens every worksheet and exports an XLSX workbook. Pandas and readxl exports preserve tabular data according to the legacy workload. Native Excel and openpyxl use their respective workbook save operations. These endpoints have different format-preservation semantics.

Each process tree runs in a Windows job with a 4096 MiB aggregate committed-memory limit and a 900-second timeout. The available-RAM reserve and its preflight check are disabled. Fresh Excel servers are identified and assigned to the job before opening the input. A separate allocation probe verified memory-limit notification and termination. Library allocation failures, job-limit events, system-reserve stops, timeouts, and successful completion are recorded separately. The peak counter may include an allocation attempt around the limit and is retained as reported by Windows. The requested size of a denied allocation is unavailable unless exposed by the library. The receipt leaves that field null.

The process-tree counters cover the library and its owned descendants. The benchmark controller is recorded separately. Launcher environment snapshots record background process memory and available physical memory. Background applications influence available memory, CPU contention, and filesystem caching. A run with ChatGPT closed therefore uses its own `standalone-pipeline-4096mib` directory. Preliminary 512 MiB and 384 MiB results remain separate. The 4096 MiB configuration has no free-physical-memory reserve.

The standalone launcher can run after the chat application closes. Completed trial identities are resumed from durable JSONL receipts. A file lock prevents simultaneous controllers from using the same result directory. An interrupted trial is retried after stale checkpoint and Excel-handshake files are removed. Temporary exports are hashed after successful completion and removed to limit disk use. All trial checkpoints, logs, timings, and memory receipts remain available.

Cleanup costs are included only where a measured, compatible local receipt exists. Gotham-produced cleanup measurements retain their host provenance separately. Missing cleanup timings remain unavailable. A completed open/save operation does not establish visual equivalence. Red failure markers on the comparison chart show termination times and retain their failure status in the machine-readable summary.

The original Wanderungen workload is included under this bounded protocol. The earlier cleanup-only harness has a narrower scope and does not open workbooks.

An Excel diagnostics log can remain locked after its owning trial has finished. Trial receipts are saved before temporary-file removal. Cleanup retries sharing violations briefly and records remaining paths in `cleanup_pending`, allowing the next trial to proceed. On restart, receipts absent from the central JSONL log are recovered by trial identity. This preserves completed measurements after a cleanup interruption. The sharing-violation path was checked with an actual Windows file handle that denied delete sharing.

For another Windows host, invoke `run_bounded.py` with its Python interpreter, a local manifest, the SheetShrink executable, and this script directory as `--helper-directory`. The result root expects psutil and pywin32 under `software/packages`, R under `software/R`, and the R packages under `software/R-library`. Install pandas and openpyxl in the selected Python environment. The plotting script uses Matplotlib. The bundled Windows job helper is the copy used for this study.

## Completion checklist

- [x] Verify all four requested Gotham outputs against their transfer manifest and original hashes.
- [x] Exercise the job memory limit with an allocation probe.
- [x] Exercise all five library entry points on a tiny workbook.
- [x] Prepare resumable independent launch, environment snapshots, input verification and exclusive result-directory locking.
- [x] Complete the independent three-repetition XLSX series for three cases and six variants.
- [x] Complete the bounded Wanderungen pandas CSV series.
- [x] Inspect every failure classification and generated chart.
- [x] Integrate supported timings, resource measurements and hardware context into the manuscript and handover.
- [x] Build and inspect the updated manuscript PDF.

The local study directory and SSH handover retain the current execution state. The broader manuscript audit remains an independent unfinished checklist.
