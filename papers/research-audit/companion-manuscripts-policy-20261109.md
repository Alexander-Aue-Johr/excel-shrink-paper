# Companion manuscripts and supplementary access

## Finding

The inspected ICDE 2027 pages do not prohibit including a related, concurrently
prepared manuscript in supplementary material. They specify single-blind review,
allow several submissions per author, and request artifacts through openly
accessible standard services. Reviewer anonymity still constrains tracking links.
No explicit companion-PDF upload procedure was found. The inference is that
transparent companion access is compatible with the inspected instructions,
subject to distinct contributions and the actual submission form.
[ICDE 2027 submission guidelines](https://icde2027.github.io/submission-guidelines.html)

IEEE's general journal policy calls for disclosure and citation of related
submitted work, together with an explanation of the difference. Its scope is
broader publishing guidance rather than an ICDE-specific companion procedure.
[IEEE submission and peer review policies](https://journals.ieeeauthorcenter.ieee.org/become-an-ieee-journal-author/publishing-ethics/guidelines-and-policies/submission-and-peer-review-policies/)

IEEE also permits preprints in specified author, institutional and approved
preprint locations. Its definition of a personal server includes a controlled
repository. This supports public manuscript sharing in principle. Treating every
third-party GitHub location as automatically approved would exceed that evidence.
[IEEE posting policies](https://journals.ieeeauthorcenter.ieee.org/become-an-ieee-journal-author/publishing-ethics/guidelines-and-policies/post-publication-policies/)

## Documented procedures at other venues

ICML 2026 expressly requires relevant concurrent submissions to be cited,
discussed and included as supplementary PDFs. Its sample bibliography includes
the manuscript title, concurrent status and supplement filename. ICML requires
anonymization under its own review process. That requirement does not transfer
to ICDE's single-blind process. This is a concrete procedural precedent, without
establishing an ICDE-specific permission.
[ICML 2026 peer review FAQ, concurrent submissions](https://icml.cc/Conferences/2026/PeerReviewFAQ#concurrent-icml-submissions)

PVLDB 2027 requires related concurrent submissions to be declared and cited with
their submission status so incremental contributions can be assessed. This is a
relevant database-venue convention, with its own policy scope.
[PVLDB 2027 submission guidelines](https://www.vldb.org/2027/submission-guidelines.html)

Two community discussions describe the same practice: cite a submitted
manuscript accurately and optionally provide its PDF for review. These are
experience-based accounts, not venue rules or proof of acceptance. They also
describe the risk of relying on a method that may change during review.
[Dependent manuscripts discussion, 2018](https://academia.stackexchange.com/questions/107367/submitting-a-manuscript-dependent-on-another-one-also-submitted/107377)
[Concurrent methods papers discussion, 2022](https://academia.stackexchange.com/questions/183051/is-it-accepted-that-i-submit-two-novel-papers-at-the-same-time-but-one-uses-the/183052)

## Application to SheetShrink and DeltaSieve

SheetShrink should summarize the native validation protocol sufficiently to
interpret its own evidence, and cite DeltaSieve for the broader selection and
capture method. DeltaSieve should identify SheetShrink as the motivating cleanup
application and explain its own transformation-independent method and evaluation.
An appropriate reference identifies a source or shared evidence even when its
full technical method is unnecessary to the current contribution.

The bibliography identifies each companion manuscript. The main text states
its relevance and the exact supplementary filename. A footnote can locate an
artifact, but cannot replace attribution of a method. CMT IDs are optional
identifiers until they exist. They do not establish reviewer access.

The prepared local paths are
`papers/supplementary/related-manuscripts/sheetshrink.pdf` and
`papers/supplementary/related-manuscripts/deltasieve.pdf`.
Within an uploaded artifact package, `related-manuscripts/` should be at its
root. Both papers can cite the same versioned package. The current copies are
working drafts; the final submission snapshot must match the actual submitted
PDFs and carry their hashes. Current bibliography entries truthfully state
preparation status. Update to submitted status only after submission, with
confirmed authors and titles.

An artifact-availability sentence is ordinary scientific documentation. It can
remain useful in the published paper if it links to a permanent version.
Temporary review wording should be updated when publication status changes.
The current main-text sentences use factual supplement locations, without
assuming that reviewers can open another CMT submission or that a temporary
review link is permanent.

Each main paper must contain enough method and evidence to assess its claims.
Additional method detail can live in the companion. The supplement cannot
resolve missing controls, unsupported aggregate claims or an unclear boundary
between the two contributions. The experimental protocol should apply the same
observation and outcome rules to SheetShrink and Check Performance, accounting
for their different transformation contracts.

No source inspected establishes an ICDE ban on factual artifact-access notes
in a review manuscript. No rule found guarantees that every reviewer will read
the supplement. A visible main-text location and a matching package index make
the material discoverable.

## Verification boundary

Checked on 2026-11-09 against the linked primary policies and community reports.
This research does not constitute a response from the ICDE chairs. The actual
CMT fields remain to be completed during submission. Publicly archiving a
companion manuscript is distinct from submitting it for independent review.
The exact final overlap between the completed papers still requires editorial
assessment when their results and claims are frozen.
