# DeltaSieve page and evidence review

## Reviewed snapshot

Paper revision `a0605bf`, 11 November 2026. The PDF has 17 physical pages.
The conclusion ends on page 16. The AI declaration and references follow.
The PDF SHA-256 is
`f466df85493a820461ee8db35b5666d19615a3cc0b4384d6927015b46fab8fc8`.
Every page was text-extracted and visually inspected in contact sheets.
All manuscript sections, figure placements and cited local evidence summaries
were inspected. Selected native-control receipts and inverse-reference tests
were cross-checked. This review adds no experiments or manuscript edits.
The final refresh includes the Ryzen cell-stream follow-up in `a0605bf`.

The official ICDE 2027 limit is twelve pages excluding references and the
AI-generated-content acknowledgment. Appendices are prohibited. Title, abstract,
figures, tables, related work, discussion and conclusion consume the allowance.
The current main text therefore needs to finish four physical pages earlier.
Actual savings require rebuilding after edits because floats move.

## Review criteria

The following distinguishes submission rules from evidence expectations.
ICDE supplies the applicable format and scope rules. ICSE, SIGMOD and PVLDB
provide closely related review and evaluation standards.

| Criterion | Consequence for DeltaSieve |
| --- | --- |
| Scope and contribution | Explain the data-engineering problem, spreadsheet-specific selection obligations and advance over existing selective visual verification. |
| Originality and positioning | Identify which combination or mechanism is contributed. Ordered merging, XML tokenization, heartbeats and queued execution are established techniques. Their presence alone supplies limited novelty. |
| Technical soundness | State supported semantics, selection invariants, unsupported features and failure outcomes. Correctness guards need arguments and controls. |
| Evidence matched to claims | A claimed speed or memory improvement needs an appropriate comparator and matching work. A required safety check can be justified through correctness evidence. |
| End-to-end utility | Include planning, preparation, native observation and completion costs when claiming practical acceleration. Preserve coverage and output quality across compared modes. |
| Independent validation | Use references that bypass the selector and mutations with observed native effects. Separate missed pixels, detected pairs and unsupported references. |
| Experimental validity | Report workload selection, repetitions, variation, cache state, hosts, versions, concurrency, resource scope and censored attempts. |
| Reproducibility | Preserve inputs, identities, exact settings, scripts and receipts. ICDE expects supplemental artifacts and explicitly requires reproducibility artifacts for EAB submissions. |
| Clear and self-contained exposition | Keep the essential method and evidence in the main paper. Move operational detail and redundant examples to the artifact. |
| Accurate scope and disclosure | Distinguish selected controls from corpus estimates, completed results from plans, and externally assisted observations from author-verified evidence. Keep the required AI disclosure. |

Official sources consulted on 11 November 2026 follow.

- [ICDE submission guidelines](https://icde2027.github.io/submission-guidelines.html) establish the page limit, format, disclosure and submission rules.
- [ICDE research call](https://icde2027.github.io/cf-research-papers.html) establishes data-engineering scope, research/EAB categories and artifact expectations.
- [ICSE review criteria](https://conf.researchr.org/track/icse-2027/icse-2027-research-track#review-criteria) name novelty, rigor, relevance, verifiability/transparency and presentation.
- [SIGMOD research call](https://2027.sigmod.org/calls_papers_sigmod_research.shtml) requires a self-contained submission and limits reliance on supplementary explanations.
- [PVLDB research categories](https://vldb.org/2027/call-for-research-track.html) distinguish algorithmic contributions from systems papers with empirical end-to-end evaluation.
- [SIGMOD reproducibility reviewer guide](https://reproducibility.sigmod.org/documents/Guidelines-Reviewers.pdf) addresses executable documentation, environmental interference and explicit partial reproduction.

An isolated performance experiment becomes necessary when the manuscript
attributes a performance advantage to a mechanism. These sources provide no
universal rule demanding a speedup benchmark for every pseudocode block,
implementation detail or correctness guard.

## Lessons from the SheetShrink edits

The relevant recent changes are `ada6eaa`, `21996a7`, `3672e29` and `2fa6b01`.
They remove the long SIMD exposition and mixed dispatch ablation from the main
paper, remove the external-parser repair narrative, remove a superseded cleanup
panel and replace a long opening/memory matrix with a compact summary.
The detailed measurements remain available in artifacts.

Apply the same selection principle here. Prioritize evidence for the central
method. Condense useful component results. Move development history and
peripheral mechanisms when their exposition costs more than their scientific role.

## Every-page inventory

Page numbers refer to this snapshot. The evidence column identifies the scope
of support and the action required if a stronger claim is retained.

| Page | Content | Editorial action | Evidence assessment |
| --- | --- | --- | --- |
| 1 | Title, authors, large Atlas/Adage illustration, abstract | Retain the independently verified fill-loss example. A smaller figure is a layout candidate. | Real localized native evidence supports the example. Atlas size and UI polish supply no measured research benefit. |
| 2 | Introduction, companion-paper relationship, comparison contract | Retain the selection obligation, effective styles and supported domain. Merge repeated selection/equality/attribution distinctions. | Central detection and cost claims require the independent reference and matched-cost results discussed below. |
| 3 | Contract continuation, overview, logical algorithm, artifact lineage and preparation | Combine the overview and stage pseudocode. Shorten the lineage to recipes needed for interpreting results. | State-pair provenance is essential. Broad preparation-cache savings remain unmeasured in this manuscript. |
| 4 | Disk records, Figure 5, style identity and dependency selection | Shorten the 31-line pseudocode and byte-level contract. Keep the two-stream design and its measured scope. | New tests and timings support the implementation. The 95.2% result isolates immediate consumption versus retained changes. Both modes already use disk streams. |
| 5 | Lexical state diagram, changed-precedent XML panel, inverse-reference equation, selection-obligation table, envelopes | Move the lexical diagram. Retain a compact dependency example and the inverse mapping with its supported domain. | Small-grid tests check inverse mapping. Current INDIRECT native controls exercise fallback. Native supported-reference branch coverage and an enumeration performance comparison remain open. |
| 6 | Adaptive-partition schematic, all-state geometry code, finite envelopes, image qualification | Keep one coverage-preserving adaptive schematic. Move the accompanying low-level code panel. | FRPP supplies a planner probe and a visibility-transition control. A completed matched fixed/adaptive comparison remains open. |
| 7 | Office.js synchronization code, run-state diagram, pixel/geometry qualification, attribution and execution | Move the synchronization and generic run-state figures. Retain image completeness and attribution invariants. | Batching has bounded component measurements. Ordinary async synchronization and run phases need little main-paper exposition. |
| 8 | Repeated cache gates, report/replay/Atlas details, monitoring APIs, corpus introduction | Condense cache admission to its compatibility facts. Move UI, replay and API details. Shorten source-history prose. | Compatibility protects evidence meaning. Cache speed, UI usability and monitoring efficiency have no integrated advantage evaluation here. |
| 9 | Large three-panel protocol Figure 12, source table, Graph OLE conversion qualification | Move Figure 12 and detailed Graph-conversion history. Keep one source table and the eligibility boundary. | Protocol traces are explicitly illustrative. Source inventories describe populations. Neither establishes completed visual coverage. |
| 10 | Independent Enron/Data.gov/Destatis references, metrics, mutations and statistical scope | Protect the reference outcomes, unsupported denominators and mutation table. Condense prospective metric instructions. | Mostly equal samples qualify bounded observations. Additional changed-input references are needed for broader selector detection claims. |
| 11 | Cost accounting, Intel and completed Ryzen stream results, batching, live package A/B/A, independent pairs | Keep the Ryzen stream result and concise qualified batching evidence. Move the preliminary Intel lane and historical tuning detail. | The Ryzen repeat passed all 63 trials and 27 focused tests. Its dense-change result is 94.0%. One real-input repetition has a 42.121 MiB working-set peak that remains in the receipts. |
| 12 | Capture-limit failures, adaptive planner, candidate compression, finite backgrounds, far-right border, fallback costs and comparator pilot | Retain capture truncation, finite-bound safety and the far-right border. Consolidate long before/after histories. | Candidate counts, rendered work and elapsed time have different endpoints. Several earlier runs are censored or stop before rasterization. |
| 13 | Comparator pilot continuation, chart controls, font/renderer/object attribution and counter-scope explanation | Keep chart-gate controls and a compact attribution example. Move redundant measurement caveats and single-run tool timings. | Chart exclusion has three native controls. Broader gate coverage and matched time savings remain open. Repeated chart variability is useful validity evidence. |
| 14 | Counter-scope continuation, 135-slot coverage table, border/PFAS/Adage examples, finite-row and zero-width controls, operational failures | Keep stage denominators and two complementary native counterexamples. Move PFAS cost anecdotes and dialog-recovery census. | Distinct selection scopes prevent a tool ranking from raw difference counts. Finite-row native masks support a core boundary. |
| 15 | Operational outcomes continuation and most related work | Move operational chronology. Preserve direct spreadsheet-dependency, selective-testing and visual-verification precedents. Shorten peripheral parsing/database analogies and repeated product-scope qualifications. | These precedents make precise spreadsheet-specific novelty important. They should survive the page cut. |
| 16 | Related work continuation, limitations, conclusion, AI declaration and first references | Condense repeated UI/replay limitations and retain scientific limitations. Keep AI disclosure and relevant references. | The main text reaches this page. The declaration and references are excluded from the twelve-page budget. |
| 17 | Remaining references | Maintain relevant and accurate references. | Excluded from the main-text limit. Removing citations supplies little page-budget benefit. |

## Mechanism-to-evidence inventory

The locations below refer to `sections/method.tex` and `sections/evaluation.tex`.
An open measurement means qualified support is absent from this reviewed mapping.
Additional historical artifacts may exist elsewhere. This audit does not certify
their absence or substitute implementation tests for native observations.

| Mechanism or claim | Existing support | Additional evidence or editorial decision |
| --- | --- | --- |
| Selected native comparison overall | Independent reference areas, nine designed mutation pairs, five comparator pilot pairs and scoped corpus samples | Highest priority. Matched targeted/reference trials with identical observation policy, native changed-pixel coverage, complete phase costs and censored outcomes. |
| Stored/save/rebuilt state controls | Generation recipes and native attribution controls | Keep the minimum recipes. Quantify preparation/cache benefit only if advertised as a contribution. |
| XML byte scanning | Scanner tests and explicit errors | Move lexical details. A parser-speed claim needs an equivalent parser baseline and matched semantics. |
| Fixed-record streams and ordered merge | 6,561 oracle comparisons, focused tests and 63 samples per host, including the completed Ryzen repeat | Consolidate on the Ryzen lane. A disk-versus-heap superiority claim additionally needs an equivalent heap-index implementation. |
| Immediate consumption versus collected changes | Ryzen peaks of 5.703 versus 94.973 MiB for one million changes, giving a 94.0% reduction | Preserve this result and resource scope. Retain the real-input outlier in reported maxima. The full pipeline stores richer evidence. |
| Effective style identities | Permutation, implicit-default and dependency tests, native representation controls | Keep. Broaden native controls only to support additional style/extension semantics. |
| Shared-string identity | A native permutation control retained one unnecessary equal capture | State the observed conservative work. Avoid a claim of eliminating every representation-only capture. |
| Layout propagation and row-run coalescing | 5,192,711 to 603,267 candidates with equal content envelopes | Candidate reduction is supported. Time/RAM improvement requires completed matched measurement if claimed. |
| Conditional rule dependency selection | Native direct/dependent pilot, rule-change controls and conservative fallbacks | Expand unchanged-rule/changed-precedent native tests for the specifically claimed supported reference classes. |
| Inverse A1 mapping | Independent forward mapping on small grids, mixed axes and reversed endpoints | Keep the equation. Add native relative/absolute/range/disjoint-area controls. Compare enumeration with inversion only for a speed claim. |
| Chart exclusion | Equal/different-source native controls and a 69-chart gate census | Test source-style, visibility, layout and conditional dependencies. A conservative-export comparison can measure avoided work and elapsed time. |
| Drawing/OLE ownership and previews | Frame, preview and zero-size-shape native mutations | Keep safety distinctions. General OLE coverage remains outside these controls. |
| Finite backgrounds and row-border envelopes | Fresh equal pairs, explicit distant-border references and matching mutation masks | Keep the finite-bound argument and negative controls. Completed same-input comparisons are needed for quantitative speedup. |
| Zero-area geometry | Visibility-transition and unchanged-input native controls | Keep as a correctness condition. Additional timing evidence is optional unless an efficiency gain is attributed to it. |
| Hole-preserving coalescing | Coverage invariant and implementation controls | A compact invariant is sufficient for method description. Preserve gaps in any paired pixel-reference test. |
| All-state adaptive splitting | FRPP planner result, large hidden-row counts and transition control | Completed fixed/adaptive comparison with equal selected union, context, pixels, probes, memory and elapsed time. |
| Context, cropping and seams | Declared seam controls and bounded native references | Include split/whole-region native controls at merge, overflow and object boundaries for the supported coverage claim. |
| Geometry batching | 64 equal scalar/batched measurements and bounded timings | Already measured. Condense and retain workload, repetition and polling qualifications. |
| Image request batching | Twelve-region identical-image control and live package observation | Preserve the controlled result. A broad gain needs matched mixed-size packages and repetitions. |
| Office.js synchronization | Required consume-after-sync behavior | Move the code illustration. Correct ordering needs correctness checks. Separate optimization claims would require measurement. |
| Image completeness and exact pixel comparison | Successful-but-truncated image probe and independent native masks | Keep. Successful API completion alone leaves completeness unqualified. |
| Extent classification | Opaque-white exterior strips, zero-area checks and far-edge controls | Keep scoped geometry/containment tests and overlapping-pixel findings. |
| Volatile/environment attribution | Formula tracing, original/original controls and native counterexamples | Keep one causal example and unresolved classifications. Diagnostic paths have a supported subset and bounded search. |
| Parallel native pairs and private ownership | Repeated one/two-pair trials and local isolation controls | Current results support the completed equal-image workload. Mixed workloads and full startup costs are needed for broad scaling claims. |
| Hash-first transport and equal-image recompression | Equal-only throughput experiment and exact decoded-pixel checks | Move historical rates unless central. General transport benefit needs mixed image outcomes, bytes, encode/decode CPU and peak memory. |
| Pixel/count credits and bounded upload lifetime | Protocol and schematic ownership model | Move detailed diagram. A bounded-memory performance claim needs asymmetric-client stress and measured pending occupancy. |
| Watchdog, heartbeat, durable retries and cleanup | Profile hazard/control and 741-pair watchdog follow-up | Retain short failure invariants. Systematic reliability claims need fault injection covering stalls, disconnects, duplicate delivery and owned exit. |
| Byte-identical bypass and rebuild eligibility | Identity and conservative feature gates | Retain the admissibility rule. Test invalidation boundaries. Add avoided-work timings only for a savings claim. |
| Completed-result cache reuse | Completeness, part coverage and renderer/recipe compatibility contract | Condense. An invalidation matrix and cold/warm trials support a future cache contribution. |
| Observatory, replay, HTTP/WebSocket monitoring and agent access | Implemented interfaces and operational use | Move UI/API exposition. A usability, monitoring-efficiency or reliability contribution requires its own evaluated research question. |
| Comparator feasibility and coverage census | Frozen input identities, exported reports, native pilot and 135-slot table | Keep endpoints and unavailable/failed outcomes. Matched scope and cost are required for cross-tool performance or preservation ranking. |

## Ranked shortening candidates

Savings below are rough layout estimates. They overlap and should not be summed
as a verified page count.

| Order | Candidate | Pages | Estimated saving | Reason |
| ---: | --- | --- | --- | --- |
| 1 | Operational execution, protocol, replay and monitoring detail, including Figures 11 and 12 | 7--9, 14, 16 | 0.8--1.2 pages | Generic infrastructure occupies substantial space. A short correctness contract preserves the scientific obligations. |
| 2 | Lexical-state and low-level geometry/synchronization illustrations, Figures 6, 9 and 10 | 5--7 | 0.6--0.9 pages | They repeat prose and the adaptive schematic, with limited independent contribution. |
| 3 | Historical budget tuning, old adapter comparisons, original-only sweep and interrupted diagnostic chronology | 11--12 | 0.4--0.6 pages | Workloads, policies and endpoints vary. Retain the matched batching control and image-truncation counterexample. |
| 4 | Extraction history, Graph recovery and opening-dialog census | 8--10, 14 | 0.3--0.5 pages | These qualify corpus preparation. A compact eligibility statement can carry that role. |
| 5 | Repeated contract, measurement-scope and future-evaluation prose | 2--3, 10--11, 13, 15--16 | 0.4--0.7 pages | State each limitation beside its result and summarize remaining scientific limits once. |
| 6 | Overlapping post-hoc examples and unmatched comparator timing anecdotes | 12--14 | 0.3--0.5 pages | Keep complementary native failure mechanisms and the main coverage table. Move per-file histories. |
| 7 | Long Figure 5, serialization detail and superseded Intel measurement prose | 4, 11 | 0.4--0.6 pages | Retain a compact merge outline and the completed Ryzen memory result. Move the earlier lane, record offsets and full loops to the artifact. |
| 8 | Cover size, duplicated overview/stage diagrams and peripheral related-work analogies | 1, 3, 15 | 0.3--0.6 pages | Retain the verified fill-loss image and direct prior work while reducing repeated exposition. |

Prioritize selector detection and common-endpoint cost experiments over further
optional infrastructure benchmarks. Negative or mixed component results can
remain scientifically useful when they answer a central question. Peripheral
results can move with their evidence even when substantial effort produced them.

## Recommended first edit

Start with candidate 1. It has a favorable reduction in space relative to lost
scientific explanation and requires no new experiment to justify shortening.

Move the following material to the artifact.

- The generic lifecycle paragraph and Figure 11 in Method J, starting at
  `Execution, omission and evidence reuse`.
- Figure 12, its protocol walkthrough, illustrative heartbeat/occupancy traces
  and detailed acknowledgment/retry narration.
- The repeated cache-admission explanation after its four compatibility facts
  have been stated once.
- The preparation-replay paragraph, Atlas theme/badge details and full
  `Machine-Accessible Campaign Monitoring` subsection in Method K.
- The dialog-recovery census and profile/watchdog development chronology under
  `Operational outcomes and evaluation limits`.
- The repeated Observatory/replay paragraph in Discussion.

Retain a short operational contract stating that each comparison owns its native
processes, records verified work separately from reachability, requires paired
qualified observations, and preserves failed or incomplete outcomes. Retain
the cache requirements for completion, input/part identity, component coverage
and compatible renderer/preparation recipes. Preserve image receipts and their
provenance in the report contract. The main coverage table keeps its failure
and unavailable denominators.

If the protocol itself remains an emphasized contribution, evaluate bounded
fault injection for disconnects, duplicate commands, stalled native calls and
cleanup, plus asymmetric-client memory behavior. A claimed throughput benefit
also needs matched equal/different image workloads with identical capture scope.
Those experiments have lower priority than the selector's native detection
coverage and matched targeted/reference cost.

## Completion

All 17 physical pages, 31 mechanism/claim groups and the recent SheetShrink
shortening decisions have an inventory. Main-paper content remains unchanged.
The proposed savings are editorial estimates pending an authorized edit and
rebuild. Existing receipts were inspected rather than rerun for this review.
