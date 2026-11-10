# FUSE maximum-cleaning search

Source: [FUSE, Zenodo record 581678](https://zenodo.org/records/581678).
The download is checked against the published size and MD5; SHA-256 is recorded locally.
Original archive contents remain unchanged.

The workflow processes native XLSX first. Legacy files are converted with Excel COM
after the existing EUSES preparation finishes, with per-file timeouts and failure records.
Only Maximum Cleaning runs. Default and Too Aggressive are excluded.

The search compares uncompressed worksheet XML sizes. A candidate must retain at most
0.001% of its baseline XML and contain a value or formula. The workflow stops after
the first batch containing such a sheet and saves both workbooks for visual review.
A size match is not evidence of visual equivalence or exhaustive corpus processing.

Local run: `F:/corpora/fuse`. Download status is in `status.json`; processing status
is in `progress.json`. `workflow.json` pins the executable, helper code and settings.
`results.jsonl` records processed files and failures. Sources and completed outputs
are reused when the frozen runner is restarted. Create `stop.request` to stop
between work items; remove it before restarting `code/run_fuse.py --output F:/corpora/fuse`.
