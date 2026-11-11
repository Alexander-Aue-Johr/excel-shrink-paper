# Working manuscripts and author review

The fresh `sheetshrink` and `deltasieve` projects contain independent narrative
drafts, editable method figures and verified initial bibliographies. Evaluation
sections distinguish observed checks from measurements still required. Historical
manuscripts inform requirements and engineering history. Their structure and
wording do not constrain these drafts.

## Build on this machine

Run the manuscript builder from the repository root with Python 3.

```powershell
python scripts/build_manuscripts.py
```

The build produces `papers/sheetshrink/SheetShrink-ICDE2027.pdf` and
`papers/deltasieve/DeltaSieve-ICDE2027.pdf`. These filenames identify the system
and target conference for uploads. Use `--paper sheetshrink` or `--paper deltasieve`
to build one manuscript. The builder also refreshes each existing `main.pdf`
for links and viewers that use that path.

TeX Live is discovered on PATH or under `C:/texlive`. Override its executable
directory with `--tex-bin`. On this machine, all temporary and auxiliary files
stay under `C:/texlive/builds/icde2027`. Elsewhere the default is the system
temporary directory. Use `--build-root` to choose another build directory.
The builder runs PDFLaTeX, BibTeX and two further PDFLaTeX passes, publishing
PDFs after successful compilation. Commit and push the named PDFs and updated
`main.pdf` copies with manuscript changes. Inspect the logs and rendered pages
after layout changes.

## Author-review checklist

- [ ] Confirm the contribution and supported preservation/comparison contract in
  each paper, including intentional hidden-tail policy and visual-area boundaries.
- [ ] Freeze corpus manifests, source and binary versions, configuration,
  environment, coverage cutoff, exclusions and unresolved findings.
- [ ] Replace pending evaluation text with verified tables. Keep successful
  workbook observations distinct from attempts, failures and unexamined regions.
- [ ] Complete independent functional checks, mutation/reference observations,
  rendering controls, and matched end-to-end performance measurements.
- [ ] Verify every citation, comparison-tool claim and figure at publication size.
  Keep each paper self-contained and map any overlap between the two submissions.
- [ ] Confirm the applicable submission category, author list, affiliations,
  ORCIDs, conflicts, artifact access and the actual CMT declarations.
- [ ] Resolve any prior-ICDE rejection waiting-period question before submitting a
  related manuscript. The current CFP includes revisions in its one-year rule.
  Record the prior decision date and the chairs' interpretation where needed.
- [ ] Check the current IEEE template, 12-page body limit, excluded reference and
  AI-acknowledgment pages, and prohibition on appendices. Recheck these rules when
  preparing the submission PDF.
- [ ] Record actual generative-AI assistance and the author's subsequent revisions.
  ICDE requests disclosure of generated content and experimental assistance.
- [ ] Read every instruction received for this particular submission or rebuttal.
  For a requested revision, map each requirement to its response, required upload,
  change marking and page accounting. Earlier-cycle instructions are historical.
- [ ] Review exported metadata, workbook screenshots and corpus distribution rights
  before making an artifact link public. Keep confidential reviews and emails in
  private storage.

The current ICDE 2027 pages were checked on 2026-11-08. The second-round deadline
is November 11, 2026 at 5 PM Pacific. Current pages describe accept/reject outcomes
without a revision round. Eligibility of a related previously rejected manuscript
remains an author/chair decision pending the facts of that earlier submission.
[Research CFP](https://icde2027.github.io/cf-research-papers.html),
[submission guidelines](https://icde2027.github.io/submission-guidelines.html)

## Assistance provenance for the author's final declaration

Codex assisted with architecture, implementation, tests, experiment execution,
analysis of retained findings, literature retrieval, and manuscript drafting.
The current drafts' prose and editable TikZ method figures include AI-generated
material. They are supplied for substantial author review and rewriting. The
user also reports earlier ChatGPT assistance. Exact model/version information
must follow the available session records. It is not inferred from the tool's
product name.

Keep an assistance log as experiments and manuscript sections change. The final
acknowledgment should identify the systems and actual scope retained in the
submitted work. Later human revision does not itself describe the preceding
experimental or code assistance. Author responsibility includes verifying the
claims, references, data interpretation, originality and disclosure.
