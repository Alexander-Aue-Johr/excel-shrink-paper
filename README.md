# Corpus preparation snapshot

This branch preserves the extraction and Excel conversion workflow used in the
Enron corpus study on 7 November 2026. It is an independent Git root so the
sources can be transferred to a separate repository. Read this file before
running the captured scripts on another machine.

## What is preserved

| Directory | Contents |
| --- | --- |
| `code/` | Exact source copies from the running, resumable preparation pipeline |
| `paper-reference/` | Corpus files introduced by the 20 local paper commits, including current converter edits and tests |
| `receipts/` | Captured configuration, Python versions, binary receipts and available PST dependency lockfiles |
| `provenance/` | Source hashes, original commit patches and pending corpus changes |

`code/` is the reference for the active extraction workflow. The older
`paper-reference/src/excel_shrink_research/enron/ingest.py` contains the
Archive.org downloader. The two implementations have different entry points
and checkpoint formats. The older inventory is insufficient as a replacement
for the current recursive inspector.

The snapshot contains source and configuration. Workbooks, mail payloads,
databases, installed dependencies and application binaries remain local.
Absolute paths in `receipts/` describe the original machine and need adaptation.

## Source selection

The XML source used for this study is
<https://archive.org/details/edrm.enron.email.data.set.v2.xml>.
It consists of 159 ZIP archives containing metadata, messages and native files.
An Archive.org page with the title Enron can identify a different collection.
The CMU/CALO mail corpus omits attachments.

The existing downloader fixes the item identifier and expected archive count in
`paper-reference/src/excel_shrink_research/enron/ingest.py`.
`freeze_sources()` saves the metadata response and its SHA-256, plus archive
names, URLs, sizes and available MD5/SHA-1 values. `download_sources()` performs
the download and integrity checks. The current `code/fetch_sources.py` handles
other configured GitHub sources; it is a separate downloader.

## Processing sequence

1. Freeze the source manifest and download archives to a selected data root.
   Keep partial downloads separate from verified files.
2. Seed each root archive into its corpus ledger. Each object has a SHA-256;
   source membership and parent/member relationships are recorded separately.
3. Claim pending objects and classify their contents. Filename suffixes help
   schedule work and identify candidates, but do not establish workbook validity.
4. Extract supported child payloads and enqueue them for inspection. Repeat
   until the recursive queue is exhausted or explicit review cases remain.
5. Reuse a prepared workbook only with exact source-hash evidence. Otherwise
   copy the source into an isolated attempt directory and run Excel conversion.
6. Validate and reopen produced XLSX files, record source/output hashes and
   export manifests. Keep reconstructed workbooks in a separate review category.
7. Run shrinking, Check Performance or content comparison on these prepared
   local artifacts as separate benchmark stages.

The current pipeline also retains non-spreadsheet payloads to support recursive
discovery and provenance. A rewrite that stores only final XLS/XLSX files must
still retain container traversal evidence and unresolved extraction cases.

## Code map

| Responsibility | Entry points |
| --- | --- |
| Corpus scheduling and lifecycle | `code/supervisor.py`, `partitioned.py`, `seeding.py` |
| Object queue, storage and exports | `code/ledger.py` |
| Container detection and extraction | `code/containers.py` |
| Inspector isolation and bounded process reuse | `code/inspection_process.py`, `workers.py` |
| Exact-content graph reuse | `code/inspection_reuse.py` |
| PST attachment reading | `code/read_pst_attachments.cjs` |
| Conversion scheduling | `code/conversion_pool.py`, `workers.py` |
| Excel process ownership and conversion | `code/excel_shrink_research/enron/excel.py`, `excel_worker.ps1`, `OwnedExcel.cs` |

Container handlers cover ZIP/OOXML, TAR, gzip, bzip2, xz, OLE documents,
PowerPoint embedded objects, MIME messages, RTF, PDF attachments and supported
7-Zip formats. The reader inspects OOXML embedded payloads without treating
worksheet XML as a separate workbook. Unsupported, corrupt, encrypted or
depth-limited containers remain review cases. Coverage of every possible
container format has not been established.

The XML container handler currently uses ElementTree. MIME parsing uses Python's
email parser. This extraction snapshot does not implement the proposed
byte-oriented workbook similarity algorithm.

## Artifacts and resume behavior

Each corpus has its own SQLite database. The study keeps these on SSD D and
all other files on F. Both locations should be configurable in a rewrite.
SQLite uses WAL, bounded transactions and a busy timeout. File extraction and
Excel operations run outside writer transactions.

The main tables describe objects, root sources, containment edges, corpus
memberships, feed positions, inspection reuse, conversion outcomes and issues.
Object inspection states include `pending`, `working`, `complete` and `review`.
Inspection workers reclaim interrupted work within their assigned SHA range.
Conversion has its own status; an inspected object can still need conversion.

Per-corpus outputs include the following directories.

| Path | Purpose |
| --- | --- |
| `extracted/` | Extracted payloads referenced by ledger object paths |
| `xlsx/<hash-prefix>/<source-sha>.xlsx` | Accepted or exact-source reused preparation |
| `reconstructed/` | Cell-copy rescue outputs awaiting review |
| `records/` | Conversion method, hashes and outcome evidence |
| `work/`, `logs/` | Attempt files and diagnostic evidence |
| `reports/` | CSV exports of workbooks, sources, containment, matches and open issues |

Use the database's object path and containment edges to locate originals.
Multiple original names can refer to one stored object. A Workbook-stream
match is recorded separately from an exact whole-file match.

The global `stop.request` stops workers at their checkpoints. Per-corpus
`pause.request` pauses inspection, and `conversion.pause.request` controls the
conversion pool. Consult the captured configuration for actual stop-file paths.
Do not copy a live SQLite database without its WAL using a plain file copy.
Use SQLite backup or stop its writers before migrating it.

## Excel conversion

Conversion runs on Windows with installed Microsoft Excel. It uses owned,
isolated Excel processes, disables macros and link updates, applies deadlines,
and records failures rather than assuming every candidate is convertible.
The captured worker includes shared-workbook save recovery, Protected View
handling and cell-copy reconstruction. Reconstruction and successful reopening
do not establish full visual or semantic fidelity.

`code/workers.py::convert_claimed()` currently passes 90 and 360 seconds to
`convert_file()`. Review the callee's contract before changing these values.
Copying an XLSX unchanged and reopening it is different from opening it and
saving it again. Record these operations separately in a benchmark.

## Reimplementation with one URL

The intended interface accepts an Archive.org item URL plus a local output
directory. Machine paths and Excel settings belong in local configuration.
That complete single-URL interface is not yet implemented in this snapshot.

Implement it by composing the following existing responsibilities.

1. Resolve and validate the item identifier. Freeze its metadata and selected
   source files before downloading. Save final URLs, byte counts and hashes.
2. Download verified files and initialize a fresh per-corpus ledger. Fresh
   initialization needs an explicit command; some captured constructors assume
   an already migrated database and reject a missing database.
3. Seed the archive objects, then run the recursive inspector independently of
   Excel conversion. Generate a manifest even when conversion is disabled.
4. Convert only pending candidates, preserving originals and recording method,
   Excel version, settings, duration, validation outcome and exception details.
5. Expose status, pause, resume and manifest export as explicit operations.

A stable exported manifest should connect the source archive URL and hash,
container/member chain, extracted object path and hash, detected format,
conversion method, output path and hash, and unresolved exception. Preserve
all source occurrences even when byte-identical payloads share one object.

Hash equality is evidence of identical bytes. Similar values or Workbook-stream
hashes do not justify deleting variants. Content deduplication remains a
separate research step.

## Performance and verification

The current implementation streams archive members and uses disjoint SHA ranges
for inspection workers. It serializes large archive access with a shared lock
to limit competing HDD reads. Parser processes can be reused with bounded
lifetimes and memory. Parallel Excel conversion has its own scheduler and memory
reserve. Re-benchmark these choices on the target SSD before increasing workers.

Use a small mixed fixture set first. Verify nested ZIPs, Office embeddings,
misleading extensions, duplicate occurrences, corrupt containers, interruptions
and restart. Compare object hashes, containment edges and open cases against
this implementation before replacing it. Keep extraction throughput and Excel
conversion throughput as separate measurements.

The preserved ledger regression tests can run from this directory with
`python -m unittest discover -s code -p test_ledger_transactions.py` after
installing the recorded dependencies. Set TEMP/TMP to a data-drive scratch
directory when the system drive is constrained. The original multi-corpus
supervisor depends on existing configuration, journals and runtime directories;
it is not a fresh-machine bootstrap command.

## History and transfer

The paper repository's local corpus work spans `0392afd` through `6388dd4`,
20 commits above `1822b55`. The remote `icde2026` branch was verified at
`1822b55cfb94f5fadeaca11c8b57f5584d13df0f` during capture.
`provenance/paper-commits.patch` preserves those commits for later migration.
The pending patch records the converter and EUSES documentation edits separately.
Unrelated paper edits remain in the original checkout.

This branch has no parent commit. It can be fetched into a new repository
without merging paper history. Its deletion follows ordinary Git reference
handling; another retained branch, tag or clone can still keep the snapshot.
