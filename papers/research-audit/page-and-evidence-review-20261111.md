# Manuscript page and evidence review

## Snapshot and scope

The review covers every physical page of the two compiled manuscripts at paper
revision `655e287`, dated November 11, 2026. Sheet Shrink has 22 PDF pages.
DeltaSieve has 17 PDF pages. Text was extracted from every page and all pages
were rendered for a visual layout overview. The LaTeX sources, retained local
result summaries and earlier method-to-experiment map were cross-checked.

This is a manuscript and retained-evidence audit. Native experiments were not
repeated during this review. The initial 449956e inventory was updated after a
fast-forward to 655e287 brought seven new commits, including the dispatch and
higher-memory follow-ups. The final inventory uses the updated compiled PDFs. A missing result below means that a qualified result is absent from
the reviewed manuscript/evidence mapping. It does not establish that no such
experiment exists on another machine or in an unexamined historical artifact.

- `papers/sheetshrink/main.pdf` SHA-256 `a135f3aa1717b87d622eab2300458d6a861d113210d446c34d2e6e22c26faee3`
- `papers/deltasieve/main.pdf` SHA-256 `3de4a5e9576cb0a7ba335e14163322a23311f80178f241ed5b1713983ab2da58`

The page numbers below refer to physical PDF pages. Figures and floats can move
after editing. Recommendations concern scientific content and claim scope.

## Sources and review standards

The official [ICDE 2027 research call](https://icde2027.github.io/cf-research-papers)
states a twelve-page limit excluding references and the AI-generated content
declaration. It prohibits an appendix. Title, author information, abstract,
figures, tables, related work, discussion and conclusion belong within the
twelve-page main paper. The explicit exclusion does not generally exempt every
kind of acknowledgment. The current main text reaches page 21 for Sheet Shrink
and page 15 for DeltaSieve, before their AI declarations and references.

ICDE also requires format compliance and a contribution positioned within data
engineering research. Supplemental artifacts are expected. EAB submissions must
provide all artifacts necessary to reproduce their results. The research call
specifies an openly accessible standard hosting service for supplemental material.
These are submission requirements with possible desk-rejection consequences.

The [SIGMOD reproducibility guidelines](https://reproducibility.sigmod.org/)
require central claims to be supported by experiments and document environment,
configuration, execution phases, raw data and graph-generation scripts. They
provide evidence standards for this audit and are not an additional ICDE rule.
The [SIGMOD reviewer quick guide](https://reproducibility.sigmod.org/documents/Guidelines-Reviewers.pdf)
also asks evaluators to eliminate interference and identify partial reproduction.

[Hoefler and Belli, SC 2015](https://spcl.inf.ethz.ch/Publications/.pdf/hoefler-scientific-benchmarking.pdf)
explain the importance of experimental configuration, variation and interpretable
performance results. The [SIGMOD research call](https://2026.sigmod.org/calls_papers_sigmod_research.shtml)
emphasizes original contributions and a self-contained main paper. It does not
supply a universal requirement to benchmark every implementation detail.

The resulting substantive review criteria are as follows.

1. Define the contribution, supported input domain and observable contract.
2. Support comparative performance claims with matched measurements. An end-to-end
   result does not isolate one component's causal contribution.
3. Support correctness with an independent oracle, positive and negative controls,
   and explicit unsupported cases. Passing implementation tests alone cannot
   establish corpus-level preservation or selector recall.
4. Use baselines with an appropriate task, configuration, endpoint and policy.
   Account for output quality together with speed or size.
5. Separate population samples, extreme-case selections and post-hoc diagnostics.
   Retain failures and censored attempts in the selected denominator.
6. Report repetition counts, variability, resource counter semantics, host,
   versions, worker counts, cache policy and environmental interference.
7. Distinguish work counts, transmitted bytes, rendered pixels and wall time.
   One cannot substitute for another without an established relationship.
8. Make parameters, fallback policies, exclusions, reference coverage and identities
   reproducible. A companion manuscript cannot carry an essential missing premise.
9. State limitations once at the appropriate scope and avoid general claims based
   on selected successful examples.
10. Preserve readability and self-containment under the page limit. Move detailed
    receipts and redundant illustrations to supplemental artifacts.

A correctness guard needs a semantic argument and regression controls. It needs
an isolated timing experiment only when its overhead or speed benefit is claimed.
Likewise, merely disclosing memchr's SIMD dispatch does not require an ablation.
Attributing a measured speedup to SIMD does require one. Missing an ablation is
not, by itself, a universal automatic-rejection rule.

## Sheet Shrink page inventory

| Page | Present content | Retain and shorten | Evidence assessment |
| --- | --- | --- | --- |
| 1 | Title, affiliations, large 144-workbook execution illustration, abstract with corpus reductions and Top49 coverage | Retain abstract and concise contribution. Strongly shrink or relocate the execution illustration. Its 17.077 s trace is one selected development run. | Abstract gives Data.gov 5.29% and Enron 18.92%, while current results give 5.27% and 18.87%. The local historical screening and fresh distribution summaries explain the different result versions. Reconcile the chosen lane. The Thuenen affiliation contains visible character-encoding corruption. |
| 2 | Cleanup-family schematic, motivation, Wanderungen hidden zeros, pre-opening distinction, start of preservation contract | Retain pre-opening obligation and one motivating example. Merge repeated contract/motivation language. Reduce the overview figure. | Explain that guarded hidden-value deletion is an explicit policy. Repetition does not prove user intent. The pre-opening distinction contributes to positioning but does not alone establish novelty over package-level cleaners. |
| 3 | Contract, witness strategy, validation boundaries, pipeline pseudocode, byte-oriented parser and memchr dispatch | Retain contract and pipeline. Condense low-level search implementation details. | End-to-end cost does not isolate byte scanning, reused Finder objects or SIMD. The newly added dispatch ablation measures selected memchr architecture paths on three files. It does not isolate the whole parser or all SIMD. No additional speed claim should be inferred from library dispatch alone. |
| 4 | Stream-state diagram, XML parsing regressions and test counts, content/style witnesses, promotion introduction | Keep attributed-value failure as one concrete parser counterexample. Move exact test counts and lexical state detail to artifact. | 31 and 155 test functions document test inventory, not independent native validation coverage. Parser allocation/memory benefits lack a matched parser baseline in the paper. |
| 5 | Native font-witness illustration, promotion thresholds and guards, tail pruning, hidden integer tails, deferred dependencies | Keep one witness experiment, hidden-tail policy and reference protection. Move detailed threshold variants and full XML expansion to artifact. | Native witness control supports the pictured effect. Thresholds 1,024/64/4,096 have no displayed sensitivity study. Such a study is needed if optimality is claimed, otherwise document them as policy parameters. Quantify retained witness cost if it becomes a central efficiency claim. |
| 6 | Enron fill-mask control, metadata rewriting, parallel chunks/admission/output ownership, validator and example atlas | Keep fill inheritance and concise execution model. Move operational detail and atlas inventory to artifact. | No matched worker/chunk/writer ablation isolates the scheduler's benefit. CRC32/length shortcut and built-in-name deletion without formula-body equivalence are correctness limitations, requiring qualification, a safer evaluated policy or evidence for the claimed domain. |
| 7 | Dense-row XML figure, corpus history and inventories, Top49 definition, fresh cleanup screening totals and times | Keep frozen populations, eligibility, outcomes and current summary. Move historical Enron preparation chronology. | Full-corpus reduction is measured. Corpus wall times include statistics/replay and concurrent native validation, so they do not isolate cleanup engine throughput or parallel speedup. |
| 8 | Scheduling schematic, namespace outcomes and new byte-search dispatch ablation protocol | Shrink scheduling illustration. Retain concise ablation protocol and eligibility. | Two same-source builds, one worker, three repetitions per configuration, 18 member-equivalent outputs. Native and memory workloads ran concurrently. This is a bounded ablation, not an unmeasured mechanism. |
| 9 | Three-corpus distributions, Table II dispatch timings, sampling protocol | Retain distributions and the compact ablation table. Condense repeated sampling promises and the figure caption. | Portable/normal median ratios are 0.930, 1.039 and 1.064. The normal path is not uniformly faster. Three repetitions under contention do not justify a general SIMD speedup or a precise corpus effect. |
| 10 | Laptop open/save protocol, larger-memory follow-up protocol and outcomes, representative native identity joins, generic evaluation-plan prose | Keep the new memory-limit experiment and one concise setup paragraph. Remove prospective instructions already answered by results. | Sixty follow-up attempts cover five endpoints at 4 and 24 GiB. All 30 fresh pipelines complete, nine Original trials complete. Post-trial representative outputs match native candidates member-wise, but per-trial member inventories are absent. Migration uses a normalized native baseline and QT/RR has zero selected captures. |
| 11 | Distributed comparator host, formatting-only NXPowerLite profile/audit, XLStylesTool inclusion, nine-repeat Sheet Shrink timing protocol | Keep profile, protected-part audits, host and endpoint boundaries. Shorten repeated provenance wording. | GUI observations have one primary repetition and include orchestration. The one-worker CLI has nine repeats. Same-host pairings do not make their endpoints identical. Protected-part exclusions are correctly retained. |
| 12 | Workflow timing Figure 10, external SheetReader bug and upstream patch, start of older six-file panel | Retain the current workflow result. Move upstream parser debugging and historical setup detail to artifact. | The parser patch control supports a specific upstream correction. It is not a rule-equivalent Sheet Shrink parser performance baseline. Older Balanced outputs remain a separate lane. |
| 13 | Historical native comparator findings, Traffic correction, Migration validator correction, preparation recovery, background and no-fill regression controls | Keep one cleaner correction and one validator-boundary example, with concise supporting controls. Move per-file recovery chronology. | Traffic and no-fill corrections have native controls. Single durations do not measure correction overhead. Historical raw pixel totals have differing scopes. |
| 14 | applyFont/automatic-height correction, historical Inquire workflow ratio, final Top49 outcomes and evidence boundary | Retain compact native outcome census and representative corrected control. Move historical Inquire lane. | Twelve workbooks have recorded differences with unresolved attribution, not twelve established cleaner defects. The evidence-boundary wording incorrectly leaves final reduction and matched opening cost generally open after reported results. |
| 15 | Table III, sizes and native observations for Migration, Finance, QT/RR and reserved Enron case | Retain a compact finalized subset in main paper. Move full case inventory and pending Enron rows to artifact until results exist. | Only two exact Finance comparator pairs have qualified native reports here. D counts pixel observations across captures. N is pending, scopes differ and overlaps can repeat pixels. Fresh Sheet Shrink uncompressed payload and visual qualification are pending. |
| 16 | Table IV part 1, Migration, six library/export endpoints, six variants, times and memory | Keep Migration as a strong motivating case in one combined figure/table. Move detailed rows and repeated notes. | Three repetitions per row, 900 s and 4 GiB cap. Original pandas failures are censored stops. Fresh Sheet Shrink completion is a demonstrated feasibility result. The capped original has no completed-time speed ratio. |
| 17 | Table IV part 2, Municipal Finance, five endpoints and six variants | Preserve the Excel counterexample and selected representative rows. Move complete matrix. | Fresh cleanup plus Excel takes 69.76 s versus 42.46 s directly, despite faster open/save phases. Keep this counterexample and explain phase/lifecycle cost before generalizing. |
| 18 | Table IV part 3, QT/RR library/resource rows, and new Table V with 24 GiB follow-up | Keep a compact follow-up table and representative QT/RR results. Move complete matrix. | New Table V gives completed ratios 2.97, 16.46 and 24.28 for three endpoints. Both Migration pandas endpoints still stop at 24 GiB. These are censored outcomes, without completed-time ratios. |
| 19 | Table VI CPU/working-set summary and Table VII older six-file reductions, with large whitespace | Move detailed resource table and historical size panel to artifact. Reflow floats. | Different peak counters are disclosed. The old six-file panel uses a7eef1a and Balanced, so it has low priority for current preservation-focused comparisons. |
| 20 | Related work, cleanup products, spreadsheet systems/parsing, SIMD methodological connection, start of limitations | Keep direct cleanup baselines, pre-opening obligation, SheetReader, simdjson and strongest data-management links. Condense broad systems survey and repetition. | Vendor capabilities need version-specific qualification. simdjson is a methodological reference, not an equivalent cleaning baseline. |
| 21 | Limitations, conclusion, AI disclosure and first references | Keep concise limitations and an actual-result conclusion. Remove prospective 'completed manuscript must quantify' wording. | Main content still reaches this page and therefore exceeds twelve pages. AI declaration and references are exempt, while discussion/conclusion are not. Existing name-equivalence and CRC limitations remain material. |
| 22 | References continued | Retain cited sources and correct metadata. Prefer compact stable evidence landing pages where traceability permits. | Exempt from the main-text page limit. Bibliographic maintenance is needed, not additional benchmarking. |

## DeltaSieve page inventory

| Page | Present content | Retain and shorten | Evidence assessment |
| --- | --- | --- | --- |
| 1 | Title/affiliations, large Atlas screenshot, abstract and 135-slot snapshot | Keep concise problem/result abstract. Reduce the UI cover and retain one scientifically informative native difference. | Snapshot counts are campaign outcome counts. They do not establish selector recall or reduced wall time. |
| 2 | Motivation, contribution framing, companion paper, preservation/coverage/verdict contract | Keep one formal contract and concise contribution list. Merge repeated statements about scope. | Define the independently justified relevant area and distinction between finding a changed pair and covering all changed pixels. The paper can remain self-contained with a brief companion reference. |
| 3 | Pipeline pseudocode, workflow diagram, native state/control lineage, formal A/B/E/C model | Keep pseudocode and lineage. Combine the redundant overview diagram into one figure or remove it. | Independent references must bypass the selector. Rebuild/Save controls constrain attribution. Neither guarantees that all supported visual effects are selected. |
| 4 | Forty-byte scratch evidence stream, new two-way-merge pseudocode, effective styles and dependencies | Retain algorithm and stream invariant. Compact repeated prose. | The merge complexity follows for valid ordered inputs. No matched scratch-versus-heap RAM/time/disk comparison is reported. |
| 5 | Token-reader diagram, obligations table, conditional reference example/inverse equation, chart/object gates and envelope start | Retain inverse mapping and gate conditions. Move lexical state machine and shorten overlapping example/table. | Small-grid controls support mapping. Native branch coverage for supported conditional and chart source changes needs consolidation. |
| 6 | Adaptive partition diagram/code, finite-background conditions, all-state image qualification, native pixels | Keep coverage invariant, all-state check and image-completeness requirement. Use one concise visual. | FRPP has planner measurements but no completed paired adaptive-versus-fixed image/time trial. Native seams, overflow and object-edge controls are important to establish coverage preservation. |
| 7 | Office.js synchronization snippet, run-state diagram, extent classifications, attribution, watchdogs/ownership/reuse | Keep native image/geometry qualification and concise execution safety. Move routine synchronization code and lifecycle state chart. | Batching has bounded measurements later. Reliability benefits of progress watchdogs and profile isolation need controlled outcomes if promoted as central performance contributions. |
| 8 | Default image retention/recompression, cache and rebuild omission, Atlas/status monitoring, corpus frame and evaluation questions | Keep cache admission invariant and short retention description. Move monitoring API and preparation chronology. | The current default uploads images and retains equal captures. Hash-first is a measured optional historical path. Current-default transport and lossless repackaging lack a matched integrated cost result. |
| 9 | Large revised default-transport/supervision figure, Graph-object preparation caveats and frozen inventory table | Strongly shrink or move transport figure. Keep current source frame in a compact table. | Figure describes current live-image default and retained equal PNGs. Illustrative timings/occupancy are not measurements. The figure displacement creates substantial layout cost. |
| 10 | Enron/Data.gov/Destatis reference samples, detection metrics, nine mutations, statistical scope, cost protocol, opening/batching results | Preserve reference outcomes and mutation table. Condense generic planned-protocol instructions. | Four Enron and four Data.gov all-equal references support bounded agreement, not recall on changed inputs. Eleven of twelve Destatis inputs remain unsupported by the extended reference preparation. Mutations are designed controls, not independent population trials. |
| 11 | Batching continuation, concurrency/hash-first results, image-limit failure, historical coverage, adaptive and background cases, long fallbacks | Keep strongest controlled batching/concurrency result and finite-envelope counterexample. Move historical budget sweep and repeated attempt chronology. | Several before/after measurements exist. Two-pair throughput is measured, three-to-eight pairs were rejected before launch. Candidate and region reductions with interrupted baselines do not quantify end-to-end speedup. |
| 12 | Far-right border and long fallback examples, five-pair product/reference pilot, representation permutations, chart controls, start of attribution controls | Keep strongest border witness, pilot scope and chart positive/negative controls. Move single-run large-product timing details. | Pilot is option-specific and does not rank whole products. Missing matched DeltaSieve timings prevent a cross-product speed claim. Chart gate cost remains unmeasured. |
| 13 | Migration context correction, chart renderer variation, object mutations, metric interpretation, Table V campaign snapshot, Enron border/PFAS examples | Keep campaign denominator and one attribution counterexample. Condense repeated metrics/caveats and post-hoc examples. | Counts from distinct selection scopes cannot rank visual preservation. Chart variability controls explain a possible source, not exclusive causation of original findings. |
| 14 | Adage reference, finite-row/zero-width controls, operational failures, related work through comparison tools | Keep native edge controls and compact reliability limitation. Move detailed recovery census. | v17 has independent matching finite-row masks. Isolated-profile and watchdog controls support local mechanisms, not a general overnight reliability gain. |
| 15 | Related work through selective/document visual testing, discussion, conclusion, AI declaration and first references | Keep strong visual-testing precedents and concise limitations. Reduce repeated baseline-scope prose. | Main text still reaches page 15. Independent detection coverage and matched targeted/reference wall-time results remain the central open evidence. |
| 16 | References continued | Retain relevant references and correct metadata. | Exempt from main-text limit. Reference [29] misparses Marc Fisher II into author components. Original authors are Marc Fisher II, Dalai Jin, Gregg Rothermel and Margaret Burnett. |
| 17 | Final references, including ReDeCheck, VISER, VERVE and document visual regression | Retain these direct methodological precedents. | Exempt from main-text limit. They substantiate prior structural filtering plus visual verification, making the spreadsheet-specific contribution important. |

## Claim-to-evidence actions

Priority A affects a central result or correctness interpretation. Priority B
supports an emphasized component contribution. Priority C supports optional
engineering detail and can be avoided by narrowing that detail's claims.

| Paper and mechanism | Existing support | Remaining action and priority |
| --- | --- | --- |
| Sheet Shrink overall reduction | Completed corpus sizes and distributions | Reconcile historical abstract values with fresh results and freeze one identified lane. A, reconciliation of existing evidence. |
| Sheet Shrink preservation | Targeted Top49 observations, independent local native controls, corrections | Classify unresolved findings where possible and report unclassified outcomes explicitly. Join exact timing-case outputs to qualified visual reports. New representative member-equality joins partly address this, but normalized/zero-capture scopes and absent per-trial inventories remain. A. |
| Sheet Shrink combined cleanup/use | 288 laptop trials plus 60 new higher-memory trials, fresh cleanup, completed/capped outcomes | Retain negative Finance Excel result and host-specific ranges. Enron remains pending. No need to repeat the delivered high-memory follow-up. A, reconcile and summarize existing evidence. |
| Sheet Shrink byte parser and SIMD | New controlled memchr dispatch ablation, 18 equivalent outputs, median/range and CPU observations | This result is now present. Repeat under isolated workload and broaden files only if a general or precise dispatch speed claim is retained. Parallel/core-level and whole-parser claims remain unisolated. B, conditional additional measurement. |
| Sheet Shrink parallel scheduling | Architecture and selected trace, concurrent corpus wall times | Matched 1/2/4/... worker measurements with fixed workbook/writer settings, plus chunking or admission control if their specific benefit is emphasized. B. |
| Sheet Shrink row/witness rules | Native positive/negative controls and test inventory | Consolidate retained/removed records and bytes by rule family. Run rule-disable ablations only for claimed independent size/time contributions. B. |
| Sheet Shrink thresholds | Explicit guards and fixtures | Sensitivity over threshold/gap sizes if optimal efficiency is claimed. Otherwise retain them as conservative chosen parameters. C. |
| Sheet Shrink guard overhead | Corrected native Traffic control, one historical duration | Matched safe input suite before/after if 'near-zero overhead' is stated quantitatively. Correctness itself already has direct local evidence. C. |
| Sheet Shrink repeated cleanup | Separate second-pass receipts, changed archive hashes on several outputs | Inspect member-level changes before claiming idempotence or attributing second-pass work to ZIP serialization. C unless idempotence becomes a claim. |
| Sheet Shrink name/identity policies | Explicit limitation descriptions | Establish formula-body equivalence or a safe evaluated exclusion for alias removal. Record exact-byte identity policy for preservation-sensitive results. A for a broad preservation claim. |
| DeltaSieve selector completeness | Nine designed mutations, five pilot pairs, bounded controls and mostly equal random references | Expand independent native changed-input controls across the supported model, report pair detection and changed-pixel coverage separately, retain unsupported sample fractions. A. |
| DeltaSieve net selection benefit | Enron 14 versus 82 rendered images and 78 versus 398 million pixels | Common-input policy-matched targeted/reference trials including planning, native work, memory, completion and censored outcomes. Existing render-work reductions do not supply total wall-time reduction. A for a speed claim. |
| DeltaSieve inverse conditional mapping | Equation, exhaustive small-grid tests, INDIRECT fallback | Add or identify native changed-precedent controls for supported mapping branches and a matched enumeration comparison if speed is attributed to inversion. B. |
| DeltaSieve finite envelopes | Native finite-row masks and large region-count reductions | Retain edge/border/object counterexamples. Matched completed old/new runtime is needed only for a numerical speedup. B. |
| DeltaSieve all-state adaptive partitions | FRPP planner probe and visibility-transition control | Completed paired fixed/adaptive trial with same coverage, geometry, seams and pixels. Measure images/probes/wall/memory together if adaptive efficiency is central. B. |
| DeltaSieve disk evidence stream | 40-byte serialization and linear ordered merge argument | Matched RAM/time/scratch bytes against an otherwise equivalent heap representation if measured memory superiority is claimed. B. |
| DeltaSieve Office.js batching | Identical geometry/images and bounded elapsed measurements | Already has a component comparison. Improve repetition/order/error reporting if making a broad claim, rather than calling it wholly unmeasured. B. |
| DeltaSieve parallel capture/hash-first | Repeated one/two-pair capture-only experiment and peak memory | Preserve equal-image workload limitation. Measure mixed equal/different captures and transferred bytes if general transport gains are claimed. Separate preparation from capture-only rate. B. |
| DeltaSieve chart exclusion | Three native chart controls, 69-chart dependency census | Native source-style/layout/visibility/CF/object-related controls for claimed gates. Time savings require a matched conservative-export run. B. |
| DeltaSieve cache/rebuild omission | Compatibility conditions and regression logic | Invalidation matrix plus cold/warm measurements if reuse speed is a contribution. C. |
| DeltaSieve operational reliability | Shared-profile hazard/control, watchdog follow-up, durable evidence | Repeated fixed-workload runs or bounded fault injection for disconnect/duplicate/profile cleanup/process exits if claiming systematic reliability improvement. Otherwise retain as operational method/limitations. C. |
| DeltaSieve monitoring/WebSocket/JSON Patch | Monitoring implementation and earlier synthetic status-byte experiment in the audit map | The paper has no integrated comparison. Measure traffic/latency/CPU only if those efficiencies are claimed. UI/API detail can move to artifacts. C. |

## Existing evidence that should survive shortening

Sheet Shrink should retain the pre-opening preservation obligation, representative
font/inheritance witnesses, the guarded hidden-tail policy, frozen three-corpus
distributions, same-host cleanup workflow measurements, selected end-to-end
opening/memory results including failures and the Finance Excel slowdown, and a
compact final native outcome summary. Replacing the detailed resource matrix with
a selective summary must preserve the full matrix in the artifact.

DeltaSieve should retain its selection obligation, effective-style and dependency
model, inverse reference mapping, finite-envelope and all-state coverage argument,
native image qualification, independent-reference and mutation results, matched
batching evidence, the campaign's explicit incomplete outcomes and direct prior
visual-testing comparisons. The Atlas illustrates findings but cannot substitute
for selector evaluation.

Troubleshooting is scientifically useful when it establishes a reusable failure
mechanism, a contract boundary or a controlled corrective result. Row-font
witnesses, far-right inherited borders, apparently successful truncated PNGs,
zero-height visible strokes and watchdog progress misclassification fit that role.
Per-file retry histories, tool-launch instructions and unmeasured cleanup folklore
belong in operational artifacts. Time spent implementing a remedy does not itself
demonstrate novelty or benefit.

## Shortening plan and layout findings

Sheet Shrink requires its main text to finish by page 12 instead of the current
page 21. The current page count is inflated by presentation as well as content.
The historical six-file Table VII and resource Table VI occupy a sparsely filled float page. The detailed opening/memory
matrix and input table occupy pages 15--18. A compact combined summary could
recover several pages while preserving a selected main-paper result and complete
artifact rows. Additional savings come from shrinking the cover, removing the
state-machine/implementation illustration overlap, consolidating historical
experiments and repeated contract caveats, and rewriting prospective evaluation
instructions. These are estimates, not a verified twelve-page rebuild.

DeltaSieve requires its main text to finish by page 12 instead of page 15.
Reducing the UI cover, removing low-level lexical-state/synchronization and run
state figures, compressing the large transport figure, moving the monitoring API
and historical protocol comparisons, and consolidating overlapping post-hoc
examples are the most direct candidates. Keep the actual coverage and mutation
evidence ahead of implementation exposition when choosing between them.

An indicative twelve-page allocation for Sheet Shrink is two pages for motivation
and contract, three for method, four and a half for evaluation, one and a half for
related work and one for discussion/conclusion. An indicative DeltaSieve allocation
is two pages for motivation/contract, four for method, four for evaluation, one
and a half for related work and half a page for discussion/conclusion. These
allocations include figures, tables, title and abstract within the totals.

The `main-no-references.tex` route in both projects also removes Related Work.
It therefore cannot verify the ICDE page cap. Count the end of all main content
in the full manuscript, including Related Work and conclusion. Any ordinary
acknowledgment should fit within the limit unless explicitly exempted.

Additional repair items include the corrupted Sheet Shrink affiliation, stale
future-tense statements, mixed software/capture-policy revisions, an old method
scope README incorrectly saying references count within twelve pages, a table
README still mentioning P, and the DeltaSieve Fisher bibliography entry. The
original Fisher publication is available at
<https://web.engr.oregonstate.edu/~burnett/Reprints/issre02.testReuseSpreadsheet.pdf>.

## Completion

All 22 Sheet Shrink pages and all 17 DeltaSieve pages have an inventory and an
editorial/evidence assessment above. No manuscript text, figures or measured
results were changed during this review. The principal outstanding scientific
work is exact-output preservation qualification for the chosen Sheet Shrink
performance cases and independent selector/cost evidence for DeltaSieve. Optional
component benchmarks should follow the strength of the claims retained in the
shortened manuscripts.
