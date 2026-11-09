# Manuscript evidence audit

## Scope and reproducibility

This audit, dated 2026-11-09, maps manuscript claims to retained evidence and
identifies gaps requiring verification. Supported research findings inform the
working drafts. Submission preparation has priority over historical editorial
cleanup. Final manuscript approval and submission are separate milestones.

The initial inspected revisions are:

- Software repository: `origin/codex/datagov-validation-watch`, `a7eef1a`.
- Paper repository: `origin/icde2026`, `652e7ac`.
- Local OneDrive runner checkout: `experiment/check-performance-onedrive`,
  `77cdd09`. Its branch is independent of the manuscript development branch.

The software snapshot was exported without checking out its branch or changing
the runner. Repository-relative paths and commit IDs identify evidence below.
Remote branches were fetched on 2026-11-09. A branch tip's commit timestamp does
not prove the server-side time of its last push.

## Mandatory execution rules

- Never open the original uncleaned Wanderungen workbook with Python/Pandas or
  export that original through Pandas to CSV, on this or any other machine.
  Historical records may be read. A cleaned-workbook test is a separate task.
- Source-code orchestration must follow Single Layer of Abstraction and read
  like prose. Apply the same principle to the second layer when it does not
  introduce overengineering, obscure implementation, or materially harm speed.
- Keep TeX temporary files, caches, generated figures, PDFs and auxiliary output
  under `C:/texlive`; use its configured build environment.
- [x] Keep every new commit human-reviewable, coherent, developer-written in
  tone, and low in cognitive load. Write code comments and commit messages in
  English. Never create merge commits. Preserve local changes before updating,
  inspect incoming changes, and integrate with a linear history. Prefer a small
  local checkpoint commit for coherent work; use a named stash including new
  files when a checkpoint would capture incomplete unrelated work.
- Use factual, project-focused language in comments, documentation and commit
  messages. Omit conversational narration and attribution of instructions.
- Treat unverified historical accounts as investigation leads until supported by retained
  artifacts. Distinguish implementation checks, development observations,
  controlled measurements and independent evaluation.
- Do not finalize or submit manuscripts, freeze unsupported results, or run
  destructive or resource-intensive historical experiments to fill gaps.

## Checklist

- [x] Verify project names and the current development branch.
- [x] Locate existing literature-review prompts.
- [x] Determine whether DeltaSieve uses MCP or general-purpose APIs.
- [ ] Inventory troubleshooting across documentation, tests, commits and issues.
- [x] Verify WebSocket, JSON Patch and other transport mechanisms in code.
- [x] Reconstruct registry, orphaned WebView, resource and unattended-run issues.
- [x] Research the publication role of engineering hurdles and negative findings.
- [x] Inventory experiments across all relevant software branches, including WPF.
- [x] Reconstruct VBA, VSTO, Excel-DNA, COM, Rust COM and Office.js alternatives.
- [x] Trace CopyPicture, clipboard, GDI and disk-based comparison experiments.
- [x] Trace COM geometry findings and limits of parallel execution.
- [x] Trace workbook opening costs and Office.js batching evidence.
- [x] Assess the publication suitability and limits of these experiments.
- [x] Map SheetShrink rules, regressions and source-workbook provenance.
- [x] Map DeltaSieve selection, capture, fallback and expensive-case evidence.
- [x] Inspect historical manuscripts and XML, regex and byte-based implementations.
- [x] Assess historical comparisons without claiming feature-equivalent baselines.
- [x] Inspect earlier size, opening and cleanup-plus-opening figures.
- [x] Assess download, extraction and preparation scripts for reuse.
- [x] Assess chart improvements and the censored Wanderungen observation.
- [x] Verify the forbidden original-workbook Pandas test remains unexecuted.
- [x] Document the abstraction and review rules for subsequent source changes.
- [x] Establish a fast, reproducible LaTeX build within the allowed directory.
- [x] Research minted/Pygments, vector code figures and ICDE/IEEE constraints.
  No verified accepted-paper precedent or specific permission was established.
- [x] Research related concurrent submissions, disclosure and supplementary papers.
  Actual CMT companion-upload options remain unverified.
- [x] Incorporate only established research-relevant findings into working drafts.
- [x] Record remaining literature-review inputs and evidence gaps.
- [ ] Low priority, after submission preparation: review published commit messages
  and historical documentation for conversational narration and attribution of
  instructions. Replace such wording with factual technical rationale. Review
  unpublished changes before publication. Preserve actual evidence provenance.

## Findings at the initial evidence checkpoint

### Names, branches and prompts

The project names are SheetShrink and DeltaSieve. These names are used
consistently in the manuscripts, roadmap and software documentation.

`a7eef1a` is the newest remote software tip by committer date among the 17
fetched branch references. This supports selecting the proposed development
branch, but does not recover server-side push timestamps. The paper branch
remained at `652e7ac` when fetched.

Two complete optional prompts exist in
`docs/planning/optional-deep-research-prompts.md`, headed SheetShrink and
DeltaSieve. They request novelty challenges, primary-source comparisons and
evaluation critiques. No external Deep Research report was found or invoked
during this checkpoint. Prompt existence is established; research output is a
separate input.

### MCP, events and data transfer

The software search covered code, tools and documentation while excluding
vendored JavaScript. The only MCP reference found is the optional adapter in
roadmap Task 08. No MCP server, tool registration or protocol implementation was
identified at this revision. Ordinary HTTP and WebSocket APIs are implemented.
This is a bounded repository inspection, not proof about untracked external apps.

Both WebSockets and a restricted JSON Patch protocol are present:

- `excel-shrink-cli/src/capture_host/resident_events.rs` and
  `office-capture/web/resident-connection.js` implement resident control delivery.
  `docs/visual-validation/resident-events.md` records reconnect, duplicate-result
  handling, durable file reconciliation and bounded in-flight image accounting.
- `src/capture_host/control_room/live_stream.rs` sends an initial snapshot then
  RFC 6902-shaped add/replace/remove operations. It samples every three seconds.
  `office-capture/control-room/live-status.js` applies only allowlisted paths
  and rejects sequence gaps. It is deliberately not a general JSON Patch engine.
- `docs/visual-validation/experiments/observatory-live-stream-20261109.md`
  records a 100-card fixture with a 1,627,309-byte baseline, 17,722-byte changed
  response and 1,415-byte unchanged-comparison response. The retained real-stream
  samples are observations, not sustained throughput measurements.
- PNG uploads and detailed evidence requests remain HTTP operations. Hash-first
  PNG transport and bounded buffering are additional mechanisms, separate from
  the live-status JSON patches.

### Troubleshooting is retained outside commit messages

The principal inventories are `docs/behavior/review-findings.md`,
`docs/visual-validation/investigations/`, the experiment READMEs and receipts,
`docs/behavior/execution-and-evidence.md`, and
`docs/behavior/targeted-validation.md`. Regressions also appear beside code and
in `regression/` manifests. Git history identifies the changes and chronology.

Three evidence boundaries matter for the manuscript:

1. `docs/visual-validation/cleanup-isolation.md` first retains passing peer
   cleanup checks alongside a running corpus, then a clean-desktop failure where
   B reported TASKPANE_RESTARTED after A's cleanup. Per-group WebView isolation
   passed the same two-group check and the four-group sequence. This supports
   recording an ownership/isolation mechanism under the tested conditions.
2. `investigations/addin-registration-accumulation-20261104.md` records 3,085
   registrations in one failure investigation and a completed small pair after
   backed-up owned-entry removal. A later inventory found 617 entries. These
   are separate checkpoints, not a monotonic series on one controlled workload.
3. `investigations/addin-startup-20261108.md` retains the stronger countercheck:
   the workload remained fast after the removed entries were restored. Thus
   cleanup cannot be assigned a general startup speedup from that sequence.
   Registry maintenance and profile isolation address different mechanisms.

Attributing long-run slowdowns generally to orphaned-profile accumulation
requires a retained causal trace. Heartbeats, useful progress and confirmed exit already
have separate semantics in the execution specification.

### Branch and experiment inventory

`research-audit/branches.tsv` records all 17 fetched branch tips, graph-unique
commit counts and non-merge patch-equivalence counts relative to `a7eef1a`.
`research-audit/experiments.tsv` records 338 experiment/readme/result/native
adapter artifacts across those tips with blob IDs and current-path relations.
Identical filenames are not treated as identical experiments or proof of merge.
The inventory uses `git for-each-ref`, `git ls-tree`, `git rev-list` and
`git cherry`. No experiment was executed.

The missing historical GUI was found on `origin/experiments/visual-validation`
at `ce0dc7b`. Its `docs/visual-validation/experiments/README.md` documents the
WPF/CLI interface and a 62-case protocol. `tools/excel-lab/` retains the VBA-local
GDI+ and Rust image references. The native adapter distinguishes bitmap/EMF,
memory-pair and persisted-image modes. Shape-bounds adapters include VBA,
PowerShell COM, Rust COM, VSTO, Excel-DNA and Office.js.

The historical `docs/visual-validation/decisions.md` explicitly warns that the
VSTO/Excel-DNA comparisons measure drawing bounds, not independent screenshot
engine throughput. It also does not assert clipboard corruption without a trace.
Its earlier .NET host selection cannot describe the newer Rust host by inference.
The image-budget replication retains 48 complete and 14 failed cases, with
unresolved drawing bounds, rather than declaring universal qualification.

`origin/experiments/corpus-quartet-benchmark` at `68349cd` retains seven
graph-unique commits and an isolated 1/2/4-group protocol. Its capture-phase
measurement starts only after all 16 processes are ready. Opening cost is a
different denominator. Full semantic reconciliation of the branch-only artifacts
remains an open evidence-reconciliation task.

### Opening experiment suitable for a bounded manuscript observation

The current branch retains
`docs/visual-validation/experiments/parallel-workbook-open/findings-four.json`
and its protocol. Four copies of an eight-sheet fixture, each sheet containing
one million rows and three numeric columns, were opened sequentially or together
in independent native Excel processes. The endpoint is verified Office.js add-in
readiness. Two reversed-order rounds give median batch times 118.388 and 41.714
seconds, respectively. The JSON records a Ryzen 7 7800X3D, Windows 11 and binary
SHA-256. This establishes bounded overlapping startup, not parallel execution of
arbitrary COM commands. It supplies a design observation, not corpus throughput.

### Method documentation map

SheetShrink already has a prose specification with a source map in
`docs/behavior/shrinker.md`, sections S1--S10. The map covers eligibility,
classification order, whitespace/protection, row representatives and promotion,
tails/drawings, references, names/calculation-chain handling and output failures.
`docs/methodology/reference-protection.md` supplements dependency policy.
Examples of executable regression sources are
`excel-shrink-core/src/sheet/{cells,rows,ranges,tail_fill}_regression_tests.rs`
and `src/reference_scanner_regression_tests.rs`. A rule-family, regression and retained-workbook map is now retained in
`research-audit/method-and-experiment-map.md`. Individual provenance links remain
uneven. The map does not assert source-workbook coverage for every test.

DeltaSieve has a specification and source map in
`docs/behavior/targeted-validation.md`, including V5--V11, and execution semantics
in `docs/behavior/execution-and-evidence.md`. `resident-events.md` supplies the
transport and buffering details. Tall-sheet guards live in
`src/regression/office_capture/tall_sheet_{analysis,selection}.rs`. The recalled
five-hour fallback and external Rust-COM geometry defect are now identified in
`research-audit/method-and-experiment-map.md`, with source revisions and limits.

### Historical code, charts and the forbidden workload

The paper repository retains the lxml-based analyzer in
`scripts/excel_shrink_analyzer/excel_shrink_analyzer.py`, historical cleanup under
`scripts/excel_shrink/`, acquisition scripts and chart/benchmark routines.
`origin/codex/corpus-preparation-snapshot` at `98c4e9d` is another migration source.
These have not been executed or represented as current, equivalent baselines.

The existing `scripts/loading_times_size_reduction/excel_library_execution_times.csv`
records `wanderungen-2010120217005.xlsx` with Pandas status `failed_timeout`,
missing successful original opening/saving durations, and failure time **690.1 s**.
The cleaned side records opening 0.277701 s and saving 0.514670 s. Separate
`excel_shrink_benchmarks.csv` rows carry historical Rust measurements without a
fully reconciled current producer/input/environment identity. Those numbers must
not be combined into a current end-to-end speedup.

The earlier interruption after more than 48 hours and the recorded 690.1-second
timeout are separate attempts. The earlier attempt's environment and result
documents remain unverified.
A later chart should display a verified timeout as a lower-bound/censored event,
not a completed runtime or invented speedup. The original-workbook Pandas test
has not been run, imported, or scheduled by this audit.

### Primary-source publication research

ICDE's EAB category explicitly recognizes new insights into existing methods.
This supports investigating systematic failure modes and alternatives as research
material. Effort spent debugging alone does not demonstrate novelty. The audit's
editorial inference is to retain mechanisms, controlled counterexamples, scope
and reproducibility consequences, with operational details in the artifact.
The PVLDB reproducibility tutorial independently motivates documented, repeatable
experimental environments. Neither source certifies acceptance of these papers.

- [ICDE 2027 research CFP](https://icde2027.github.io/cf-research-papers.html)
- [Fischer, Hirn and Kul, PVLDB 2024](https://www.vldb.org/pvldb/vol17/p4221-hirn.pdf),
  DOI `10.14778/3685800.3685840`.

The current ICDE instructions specify single-blind review and a 12-page body,
excluding references and the AI-content acknowledgment, with no appendix.
The guidelines page measures the interval from the last submission; the CFP
instead says rejection. This discrepancy between official pages was rechecked
on 2026-11-09. Public artifact URLs are expected. The inspected instructions do
not explicitly establish a confidential companion-manuscript upload procedure.
The submission planning record appears below.
The actual CMT options remain unverified.
[ICDE 2027 submission guidelines](https://icde2027.github.io/submission-guidelines.html)

No explicit minted ban or ICDE-specific minted approval was found in those
instructions. The official minted manual documents Pygments and cached
highlighting. IEEE's PDF preparation guidance accepts vector graphics generally,
subject to publication fonts and PDF checks. This supports a technically
plausible pre-rendered code-figure route, not an ICDE permission guarantee.
The ACM accepted-package page returned HTTP 403. A public IEEE VIS submission
source loads minted, but its accepted IEEE publication status was not verified.
It is not recorded as an accepted-paper precedent.

- [minted manual](https://tug.ctan.org/macros/latex/contrib/minted/minted.pdf)
- [IEEE PDF eXpress source preparation](https://ieee-pdf-express.org/External/PreparingSourceFiles)
- [IEEE Xplore PDF requirements](https://conferences.ieeeauthorcenter.ieee.org/write-your-paper/meet-ieee-xplore-requirements/)
- [ACM package list, inaccessible during this check](https://authors.acm.org/proceedings/production-information/accepted-latex-packages)
- [PyQtGraph IEEE VIS submission source](https://github.com/pyqtgraph/IEEEVIS2021-pyqtgraph/blob/main/template.tex)

## Submission and measurement decisions, 2026-11-09

Both manuscripts target ICDE 2027 with planned submission on 2026-11-11.
The submission history records rejection of the related historical work in the
previous year's second round and no other current review. The exact historical
submission date remains undocumented. Retain the policy provenance above as part
of the submission record. Eligibility review is separate from draft preparation.

The shorter Wanderungen timeout was selected in a later attempt for chart
presentation. The current chart evidence is the recorded 690.1-second censored
observation. The earlier attempt exceeding 48 hours requires its machine,
environment and result documents before use as a measured manuscript result.
Keep both attempts separate and retain the prohibition on rerunning the
original-workbook Pandas workload.

## Open evidence requirements

1. Locate the machine/setup and retained report for the earlier more-than-48-hour
   Wanderungen/Pandas attempt when available. Add a provenance note before using
   it as a measured manuscript result.
2. The DeltaSieve report has been reviewed at the literature checkpoint below.
   The SheetShrink report remains pending. Continue independent evidence and
   submission preparation in the meantime.

The confidential-companion-paper route and a verified accepted-paper minted
precedent remain open research items. Unchecked entries identify outstanding work.

## Working-draft verification

Both manuscripts were rebuilt with the local minimal TeX Live installation.
The earlier checkpoint contained four and six pages. The expanded drafts now
have five and seven pages. Compilation has no undefined references or overfull
boxes. Underfull-box diagnostics remain, including long bibliography URLs.
All twelve pages were rendered and inspected at the expanded checkpoint.
The recorder output paths stay under C:/texlive. These are working drafts.

The working drafts and remaining evidence requirements are retained together
for the next research and submission-preparation checkpoint.

## Expanded checkpoint

`research-audit/method-and-experiment-map.md` records rule families, regression
sources, retained workbook examples, historical alternative protocols and their
measurement limits. `submission-readiness.md` separates prepared draft material
from final author, artifact, measurement and disclosure inputs.

The DeltaSieve draft now includes the retained geometry/capture batch experiment
and the two still-progressing whole-grid object fallbacks. SheetShrink's method
now explains masking blanks, font witnesses, cross-coordinate visual anchors
and the admitted hidden integer-tail structure. Both compiled drafts were
rendered and inspected after these changes.

The chart-only script reads historical CSV rows without importing a workbook
loader or old benchmark module. Matplotlib 3.11.2 and its dependencies are
installed under `C:/texlive/python-packages`; its configuration and generated
assets remain within `C:/texlive`. The preview remains separate from the paper
pending historical producer/environment reconciliation.

Repository documentation, regression sources and commit history were searched
for operational failures. The GitHub issues collection returned HTTP 404 through
the available connector. Issue coverage is therefore unverified, and the
troubleshooting inventory checkbox remains open. This response does not establish
that the repository has no issues. The branch inventory covers all 17 fetched
references; selected retained experimental protocols have been inspected in
depth, with full semantic reconciliation of branch-only code still outstanding.

Deep Research results and the earlier Wanderungen report remain pending external
inputs. Published-history wording review retains its lower priority. Final
evaluation freezes, paper finalization and submission have not been performed.

### Published wording search

An exact-phrase search of published paper history and tracked Markdown/TeX found
no occurrences of `the author requested`, `the user requested`,
`the user instructed` or `the actor requested`. The same search of software
remote commit subjects found no matches. The current software documentation
contains six matching lines, retained at `a7eef1a`:

- `docs/planning/sheetshrink-deltasieve-roadmap.md`, lines 71, 117, 178 and 996.
- `docs/visual-validation/experiments/comparator-pilot/README.md`, line 181.
- `docs/visual-validation/investigations/enron-capture-idle-recovery-20261104.md`,
  line 35.

These are already published software documents. The current unpublished paper
changes contain no matching phrasing. This bounded phrase search does not cover
every possible paraphrase. The lower-priority cleanup checkbox remains open for
a broader editorial pass after submission preparation. Existing published Git
history has not been rewritten.

### Algorithm presentation checkpoint

Both method sections now include a numbered logical-stage summary using
`algpseudocode` from algorithmicx. Ordinary IEEEtran figure environments retain
the class's captions and numbering. The summaries describe method-level
operations and do not assert literal implementation call order or replace the
detailed guards. The IEEEtran HOWTO recommends algorithmicx-compatible content
inside standard figures.
[IEEEtran HOWTO](https://ctan.math.illinois.edu/macros/latex/contrib/IEEEtran/IEEEtran_HOWTO.pdf)

The local TeX installation gained algorithmicx and the AMS Type 1 fonts required
for its mathematical input notation, installed from the verified ISO. Both
drafts rebuilt successfully with resolved references and no overfull boxes.
The two algorithm pages were rendered and inspected. Page counts remain five
and seven. All build output remains under `C:/texlive`.

CMT companion-paper entries are handled during submission. No public companion
PDF service URL has been selected. The local manuscript PDFs remain at
`C:/texlive/builds/sheetshrink.pdf` and `C:/texlive/builds/deltasieve.pdf`.

### Companion-material policy checkpoint

The researched local companion package now lives at `papers/supplementary/`,
including both working PDFs under `related-manuscripts/`, a package index and
a hash manifest. The two manuscripts cite each other and identify their
respective methodological roles. Build intermediates remain under `C:/texlive`;
the supplementary directory contains authorized distribution copies.

`research-audit/companion-manuscripts-policy-20261109.md` records ICDE and IEEE
guidance, the explicit ICML procedure, relevant PVLDB practice and two community
accounts. Its conclusion distinguishes a supported submission practice from
an explicit ICDE approval. Companion PDFs provide optional detail; each main
manuscript retains its own necessary method and evidence. Working snapshots and
preparation status must be updated to the actual submission versions at freeze.

### DeltaSieve literature checkpoint

The supplied DeltaSieve report was inspected in Markdown, PDF and DOCX forms.
Input hashes, verified sources, corrected claims and manuscript changes are
retained in `research-audit/deltasieve-literature-review-disposition.md`.
Private source copies remain with the audit inputs outside the paper artifact.
The report's unresolved numbered references required independent primary-source
verification.

Related work now covers SheetDiff and RowColAlign, and identifies Synkronizer
and SpreadsheetsDiff as further vendor-documented comparator candidates. The
existing Ekstazi and ReDeCheck references correctly use ISSTA 2015 and 2017.
The report's alternative venue/year and mechanism descriptions were corrected
in the disposition. No exhaustive priority or universal soundness claim follows
from this search.

The evaluation now defines pair-level detection, partial coverage, usable native
reference fractions, selection overhead, mutation strata, repeated-state
controls, cluster-aware uncertainty and matched end-to-end cost accounting.
The illustrative zero-miss binomial bounds are calculations under stated
sampling assumptions, separate from measured results.

DeltaSieve rebuilt to eight pages. All eight pages were rendered and inspected.
The build has resolved references and no overfull boxes. Underfull-box spacing
diagnostics remain. Supplementary working snapshots and their hash manifest were
updated to the rebuilt PDFs. The SheetShrink report remains pending.
