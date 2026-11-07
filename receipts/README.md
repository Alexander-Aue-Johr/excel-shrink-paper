# Corpus preparation

Local, resumable preparation of Enron, EUSES, FUSE, Destatis and Data.gov inputs.
Original files remain unchanged. All temporary files and outputs use F:.

The queue recursively inspects archives and embedded files in Office documents,
RTF, PDF and MIME messages. It identifies Excel files by content, including
attachments with incorrect or missing suffixes. Workbook XML stays inside its
XLSX package; it is not treated as an independent attachment.

PST files are read again with every binary attachment retained. FUSE joins when
the user's 7-Zip extraction ends. Existing TAR outputs are checked against TAR
member bytes before reuse. Missing or different members are extracted separately.

Conversions use isolated Excel processes with macros and link updates disabled,
per-file timeouts, the shared-workbook save retry, and a fresh reopening check.
For File Block failures, the worker attempts to open the individual copy in
Protected View and enable editing. Global Trust Center settings are unchanged.
The local pilot successfully converted, reopened and Maximum-cleaned two FUSE
BIFF8 workbooks after this fallback; EUSES exact-source reuse also passed.
An existing conversion is reused only for an identical source SHA-256, with its
previous preparation evidence recorded. FUSE outputs receive Maximum cleaning.
Needle candidates are recorded without stopping the queue.

`status.json` gives current counts. `reports/` contains source, membership,
workbook, Enron-match and open-case CSV files, refreshed every 20 minutes.
Whole-file matches and Workbook-stream-only matches are distinguished.
The old Enron release cannot count as evidence of finding a missing file in
the new Enron sources.

Password protection, damaged files, unsupported containers and reader errors
remain explicit open cases. Completion of the queue does not imply every file
was recovered. Review those cases before publishing a completeness claim.

Run `python code/supervisor.py --root F:/corpora/preparation-pipeline-20261031`
using the Python executable recorded in the launch receipt. To pause, create
`stop.request`; workers finish their current bounded operation. Remove it to
resume. The workflow uses checkpoints and an SQLite ledger.

Code and dependency receipts freeze this run. This directory is independent of
the Excel Shrink and paper repositories; publication location remains undecided.

Aggregate counts refresh every two minutes; process health refreshes every ten
seconds. See statistics_updated_utc in status.json for the count timestamp.
The supervisor-only I/O change and measurements are recorded in maintenance/.

Confirmed Excel save errors can now trigger an all-cells clipboard rescue.
Reconstructed files are saved under reconstructed/ with review status. They do
not enter the regular corpus or Maximum queue before review. Method and original
errors are recorded; a successful reopen does not establish visual identity.
