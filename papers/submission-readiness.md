# ICDE 2027 draft preparation checkpoint

## Prepared material

The two IEEEtran working drafts contain separate preservation contracts,
methodology, evidence boundaries and evaluation plans. The expanded checkpoint
contains eight SheetShrink pages and nine DeltaSieve pages including references.
Neither draft is a completed twelve-page submission. The source and experiment
maps identify established results that can be used during further writing.

SheetShrink contributes cleanup decisions and preservation structures. DeltaSieve
contributes candidate selection, native capture and retained comparison evidence.
Each paper must state enough of its own method and evaluation to stand alone.
Shared corpus preparation, native observations and companion methods need explicit
attribution. Reusing an observation requires the same source identity and scope.
The companion paper's existence cannot discharge an independent evaluation duty.

## Submission checks requiring final inputs

- [ ] Replace working-draft author blocks with confirmed names, affiliations,
  contact details and ORCIDs.
- [x] Review the DeltaSieve literature input, verify primary sources and update
  related work and evaluation definitions. The disposition is retained in
  `research-audit/deltasieve-literature-review-disposition.md`.
- [x] Review the SheetShrink literature input against pre-native-loading cleanup
  and verify its sources. The disposition is retained in
  `research-audit/sheetshrink-literature-review-disposition.md`.
- [ ] Once manuscript methods and results are sufficiently settled, perform
  another related-work search using the manuscript as the primary attachment
  and checking close approaches against its explicit transformation contract.
- [ ] Freeze current corpus outcomes and independent-reference denominators,
  with attempt identities, software versions and unfinished outcomes.
- [ ] Complete matched cleanup, size and end-to-end cost measurements for
  SheetShrink. Historical CSV results cannot replace this experiment.
- [ ] Reconcile the DeltaSieve mutation matrix, reference sample and matched
  cost evidence before drafting aggregate detection or speedup claims.
- [ ] Select an openly accessible standard-service artifact URL and review
  contents for sensitive workbook data and local paths before release.
- [ ] Complete the CMT fields for related-work disclosure and companion
  manuscript handling during submission. The local companion snapshots live in
  `papers/supplementary/related-manuscripts/`, with a package index and hashes.
  The final uploaded or linked package must preserve the cited paths.
- [x] Add reciprocal working-manuscript citations and explain the contributions.
- [ ] Confirm author metadata, update companion status after submission and
  replace working snapshots with the corresponding submitted PDFs.
- [ ] Verify the required content-generation acknowledgment against the official
  ICDE policy and the material included in the final submission. Routine editing
  and grammar assistance have a different stated scope in that policy.
- [ ] Run final font, page-limit, link, bibliography and PDF compliance checks.

The public guidelines permit up to five submissions per author in a round.
They require a twelve-page body limit, no appendix, ORCIDs and a PDF submission.
The artifact instructions require a publicly accessible standard-service URL.
Their wording does not establish approval for overlapping companion submissions.
These requirements were rechecked on 2026-11-09.
[Official submission guidelines](https://icde2027.github.io/submission-guidelines.html)

The companion-material policy investigation found no explicit ICDE prohibition
on the prepared approach. A documented concurrent-PDF procedure exists at ICML,
and database-venue guidance supports explicit related-submission disclosure.
The source-specific findings and their limits are retained in
`research-audit/companion-manuscripts-policy-20261109.md`.

The announced second-round deadline is November 11, 2026, at 5 PM Pacific Time.
Submission eligibility has a separately recorded policy discrepancy in the audit.
Draft preparation follows the planned submission date.
[Official research CFP](https://icde2027.github.io/cf-research-papers.html)

## Code presentation and build

The drafts use algorithmicx's algpseudocode layout for logical method summaries,
inside ordinary IEEEtran figure environments with native captions. This follows
the IEEEtran HOWTO's algorithm guidance. Algorithm source files live under each
paper's `figures/` directory. The package was installed from the verified local
TeX Live ISO, with temporary and build files retained under `C:/texlive`.
[IEEEtran HOWTO](https://ctan.math.illinois.edu/macros/latex/contrib/IEEEtran/IEEEtran_HOWTO.pdf)

The drafts require no shell escape. Minted remains technically plausible
for a later code figure, with a pre-rendered vector asset and embedded text fonts.
No verified accepted ICDE/IEEE precedent or ICDE-specific permission was found in
the bounded search. An accessible submission source alone does not establish
acceptance. Record this uncertainty when selecting the final presentation.

The local build uses `C:/texlive/environment.ps1` and
`C:/texlive/build-icde.ps1`, with output in `C:/texlive/builds`.
Sections use `input`, avoiding nested-include restrictions. Figures are separate
assets and can be regenerated only when their inputs change. Bibliography changes
require BibTeX and enough subsequent LaTeX passes to resolve references.
The current small drafts build in a few seconds. Introducing a new build package
or externalization dependency is unnecessary at this checkpoint.

The historical Wanderungen chart preview is generated only from retained CSV
rows by `scripts/loading_times_size_reduction/plot_retained_wanderungen.py`.
Its PDF and PNG are under `C:/texlive/figures`. A logarithmic axis separates the
range of observed durations. The censored timeout is explicit, and cleanup time
is excluded. Keep this preview outside current-results tables until producer and
environment provenance are reconciled. Never run the original-workbook Pandas
benchmark to recreate the missing historical record.

## Outstanding editorial review

Long bibliography URLs generate underfull-box diagnostics. The rendered drafts
have no observed clipping or overlap. Bibliography consolidation and final page
balancing belong to the final writing pass. The current checkpoint has no
undefined references or overfull boxes.

Published historical commit-message wording has lower priority than submission
preparation. Unpublished changes are reviewed for factual English wording and
linear, human-reviewable commits. Historical rewriting remains a separate task.
