# Diagram opportunities in DeltaSieve and SheetShrink

Reviewed on 11 November 2026 against source revision
`1b9ee98e3409c3584ffd12a661d0c19f8a5f0788` in `F:/repos/excel-shrink-rust`.
This extends the [state-machine inventory](../state-machines/README.md).

The request's word “Styling” can refer to Observatory presentation or to work
stealing. Both are covered below. The inventory follows the relevant source
contracts, orchestration, rendering and regression cases. Diagram proposals
describe what the implementation can support and what further observations
would be required.

## Recommended figures

| Priority | Figure | Question answered | Suggested placement |
| --- | --- | --- | --- |
| 1 | DeltaSieve artifact provenance DAG with a compact state-by-operation matrix | Which controls were generated, observed, reused or omitted, and why? | Method, two columns |
| 2 | DeltaSieve capture-coverage map with adaptive partition detail | How does a small capture plan retain the selected coverage across every enabled state? | Method, one column with an enlarged inset |
| 3 | SheetShrink hierarchical fork/join graph beside aligned thread lanes | How do workbook concurrency, worksheet work and row chunks expose parallelism? | Method, two columns |
| 4 | DeltaSieve command/acknowledgment sequence | How does the host coordinate native Excel work and recover transport interruptions? | Implementation, one column |
| 5 | SheetShrink linked timeline and before/after atlas | Which worksheet produced these tasks, and what reduction resulted? | Overview or cover |
| 6 | DeltaSieve dependency graph and sparse change/region matrix | Why can a change in one place require an image elsewhere? | Method or appendix |

The first three explain mechanisms central to the contributions. The remaining
forms are useful when their question is not already answered by an adjacent
figure. A paper benefits from a small complementary set.

## DeltaSieve source inventory

Paths below are relative to the source repository. These representations extend
beyond the parser and run transition machines already inventoried.

| Source contract | Representation in code | Suitable diagram and encoding | Evidence boundary |
| --- | --- | --- | --- |
| `excel-shrink-cli/src/regression/targeted_comparison/telemetry.rs`, `LaneId`, `Phase`, `Lane` | Six lane IDs, twelve operation labels, lifecycle, per-step status, process identity and capture receipts | State-by-operation matrix. Rows are actual lane labels, columns follow `Phase::columns()`. Symbols encode done, skipped, failed and unavailable | A completed operation and a visual-equivalence verdict require separate symbols |
| `targeted_comparison/execution_plan.rs`, `InputIdentity`, `StateDecision`, `RebuildCheck` | Raw-byte identity proof, observed/omitted/disabled decisions and a set of rebuild blockers | Provenance DAG with operation-labelled edges and proof badges beside omitted observations | A byte-identity proof and a dependency-based rebuild omission have distinct preconditions |
| `crates/deltasieve-engine/src/reuse.rs`, `PartSelection`, `PairFingerprint` | Selected package categories, per-part content hashes for both comparison partners, previous selection coverage | Two aligned package-part matrices with a reuse gate. Show the policy selection and both partner identities | Reuse is conditional on the selected policy and previous evidence coverage. It does not establish a new visual-equivalence result |
| `targeted_comparison/performance_dependencies.rs`, `reusable_rebuild` | Conservative checks for reusing historical rebuild results | Small decision diagram adjoining the provenance DAG | Unknown dependencies retain a fresh rebuild path |
| `targeted_comparison/capture_layout.rs`, `Geometry`, `coalesce`, `plan_batched` | Coverage-preserving coalescing followed by measured adaptive splitting across live states | Worksheet coverage map, then a partition tree inset. Equal selected coverage is shown before and after | Fan-out is adaptive, up to 128 children per expansion. Label this an adaptive partition tree. Hidden rows and zero-area regions need explicit treatment |
| `targeted_comparison/chart_dependencies.rs` and `conditional_formatting/references.rs` | Supported chart-input proofs and inverse same-workbook A1 dependencies | Directed dependency graph with changed inputs and affected output regions, optionally paired with a sparse adjacency matrix | Unsupported or incomplete dependency analysis retains image evidence. Arrows must distinguish dependence from execution order |
| `targeted_comparison/telemetry/timing.rs`, `PhaseTiming`, `OperationTiming` | Accumulated durations, call counts, failures and operation latency components | Aligned phase-duration bars and per-operation small multiples | These structs aggregate requests. Reconstructing every request interval requires a retained event trace. Batch wall time can remain unassigned to a lane |
| `excel-shrink-cli/office-capture/web/resident-connection.js`, `createJobConnection`, `postResult`, `runJobs` | Authenticated connection readiness, one pending command request, result acknowledgment and bounded reconnect attempts | Sequence diagram with host, transport, add-in and Excel lifelines. Draw retry loops around transport exchanges | Successful delivery and successful Excel work require their own receipts. The reconnect bound is 120 seconds in this snapshot |
| `excel-shrink-cli/office-capture/control-room/worker-monitor-model.js` | Read-only interpretation of process, connection, command and progress observations | Aligned observation tracks, with separate process-alive, heartbeat and verified-progress indicators | Heartbeat age alone supplies no scientific progress verdict |
| `RebuildCheck.blockers` and the dependency inventories | Overlapping reasons for retaining expensive controls | UpSet intersection plot, with workbook counts per intersection | Aggregate a declared cohort and normalize reason categories first. The current blocker strings can contain worksheet names |

`targeted_comparison/` in this table abbreviates
`excel-shrink-cli/src/regression/targeted_comparison/`.

### Capture example to draw

Use a selected area with an intentional hole and a tall formatted band. Show
coalescing without filling the hole, measured geometry in two workbook states,
and a split driven by the larger observation. Give each final leaf the same
identity in both states. Include a zero-area leaf and an indivisible oversized
cell as separate outcomes. The code's edge and pixel budgets are constraints on
native image export, while the candidate grid indexes selected coverage.

Existing regression anchors in `capture_layout.rs` include
`merging_preserves_holes_and_separates_sheets`,
`every_state_controls_splitting_and_no_rows_are_lost`,
`million_uniform_rows_avoid_binary_tree_round_trips` and
`invalid_or_indivisible_oversized_geometry_fails_instead_of_skipping`.

### Quantitative alternatives

A LineUp-style aligned table suits candidate ranking by reduction, elapsed time,
peak memory and capture cost when those measurements exist for the same cohort.
Keep units visible and retain missing values. A cohort alluvial diagram can show
mutually exclusive workbook outcomes with a declared denominator. Artifact
cloning has branching provenance, so a width-conserving flow needs a separately
defined quantity. Time Curves would require a justified state-distance metric
and are lower priority for explaining this pipeline.

## SheetShrink execution and work distribution

| Source | Mechanism | Recommended representation |
| --- | --- | --- |
| `excel-shrink-cli/src/pipeline.rs`, `run_input_files`, `WorkbookLimiter` | Bounded active workbooks, dispatcher outside Rayon, permits retained through writing | Resource ownership diagram with dispatcher, Rayon pool and writer boundaries. Show permit lifetime through finalization |
| `pipeline.rs`, `process_initial_sheets` | Shared size-ordered sheet queue, one sheet claimed at a time, parallel workers, ordered result restoration | Fork/join DAG plus a ready-queue detail. Source-order indices survive execution-order differences |
| `excel-shrink-core/src/work_priority.rs`, `admit_sheet`, `PendingChunk` | Count pending chunks by byte size and defer a new sheet when a larger submitted chunk is pending | Compact admission decision inset using bytes as the comparison unit |
| `excel-shrink-core/src/sheet/sheet_data.rs`, `clean_chunks_in_parallel` | Row-aligned chunks, batches bounded by thread count, parallel cleanup, ordered assembly | Nested worksheet band split into numbered chunks, followed by a join and ordered output. Mark only the submitted batch as pending |
| `excel-shrink-core/src/work_events.rs`, `TaskKind`, `Task`, `TaskGuard`, `Observer` | Workbook boundaries, initial/final sheet tasks, chunk index/count/range, actual Rayon thread index and task lifetime | Thread swimlanes with stable worksheet identity and task-phase glyphs |
| `excel-shrink-core/src/profiling.rs`, `Handoff` | Handoff timing alongside phase profiling | Waiting and execution intervals where the underlying retained observations support them |

The observer exposes execution assignment. Individual steal events, victim
deques and stolen task counts would require additional instrumentation. A
conceptual Rayon deque illustration should be labelled as an execution-model
illustration. Workbook admission and nested sheet/chunk execution have different
owners and should remain visible as separate mechanisms in the figure.

For a compact paper figure, use three workbooks, several worksheets and one
chunked worksheet. Give chunks the worksheet's identity color, print chunk
indices, and align the timeline with an adjacent hierarchy. Use a fixed time
scale. Byte volume, task duration and result reduction each need an explicit
visual channel. The current staircase's vertical coordinate follows workbook
index and step height, so its slope requires that explanation.

## SheetShrink presentation inventory

All paths in this section are under `crates/preparation-web/ui/`.

| Owner | Current responsibility | Diagram implication |
| --- | --- | --- |
| `live.ts`, worksheet projection | Worksheet ID combines workbook job ID and package-part path. Hue combines workbook hue and sheet index. Tasks reference those identities | Preserve identity across filtering and replay. Carry a short worksheet label alongside color |
| `lib/model.ts` | Workbook, worksheet, job, snapshot and playback contracts | Use containment for hierarchy and intervals for work. State and identity can use separate channels |
| `lib/scene.ts`, `sheetColor` | Shared hue mapping with optional chunk saturation parameter | The actual scene call sites use worksheet hue. Chunk index can be encoded by labels or texture without changing worksheet identity |
| `lib/worker-scene.ts`, `buildWorkerScene` | Continuous time, worker lanes, workbook step coordinate, task geometry, selection dimming, batching by appearance | A flat aligned timeline offers the most precise duration comparison. The oblique staircase remains useful as an overview |
| `lib/workbook-scene.ts`, `getAtlasBooks`, `workbookPage`, `buildWorkbookScene` | Historical selection, pagination, before/after sheet geometry, byte sizes and worksheet targets | Paired bars or a slopegraph are compact alternatives to the 3D atlas. Label uncompressed worksheet bytes explicitly |
| `main.tsx` | Hover/pinned worksheet focus, original/optimized focus variant and shared controls | Both atlas variants can refer to the same tasks. The selected variant determines the connection endpoint |
| `components/SheetConnections.tsx` | Focused-sheet projected cubic connections | Draw links for the focus selection and retain quiet context to limit crossings |
| `components/SceneViewport.tsx` | Rendering surface, hit targets, projection updates and tooltip behavior | Screen geometry and semantic worksheet IDs need to stay aligned during camera changes |
| `components/LinkedScenesPanel.tsx` | Shared timeline/atlas panel | Shared selection and caption should explain both views together |
| `lib/presentation.ts`, `lib/camera.ts` | Presentation colors and camera geometry | Theme and projection are presentation transforms. Preserve data scales and identity across variants |
| `lib/renderer.ts`, `lib/export-scene.ts` | Interactive primitive batches and vector export with projected links | Export the same selected identities and declare any layout changes made for print |
| `lib/replay.ts`, `replayAt` | Historical reconstruction from events up to the requested timestamp | Show replay time and speed separately from measured execution duration |

Relevant existing checks in `test.mjs` cover stable hue across filtered pages,
shared original/optimized worksheet targets, actual chunk worker projection,
submillisecond timing, export anchors and escaping, pagination, and replay
without future metadata. These tests were inspected as source during this
inventory. This documentation change does not rerun the application.

## Research and design references

The award category below is the verified recognition. The visual lessons are
our assessment of how each work applies to these systems. The cited records
establish recognition of the research or poster, with aesthetic relevance
assessed separately.

| Work | Verified recognition | Design lesson and application |
| --- | --- | --- |
| HOLA, InfoVis 2015 | Best Paper | Alignment, regular spacing and restrained orthogonal routing suit artifact and execution DAGs. [Paper](https://skieffer.info/publications/kieffer2016hola.pdf), [institutional award record](https://ialab.it.monash.edu/2015/10/) |
| Visualizing Dataflow Graphs of Deep Learning Models in TensorFlow, VAST 2017 | Best Paper | Hierarchical grouping and collapsed repeated structure suit workbook/sheet/chunk fork/join graphs. [Paper](https://idl.cs.washington.edu/files/2018-TensorFlowGraph-VAST.pdf), [official award record](https://ieeevis.org/year/2017/info/awards/best-paper-awards) |
| LineUp, InfoVis 2013 | Best Paper | Aligned attribute bars and linked rankings suit candidate selection across reduction, time and memory. First-page figure inspected in the original PDF. [Paper](https://data.jku-vds-lab.at/papers/2013_infovis_lineup.pdf), [author's publication and award record](https://www.sgratzl.com/publication/lineup/) |
| UpSet, InfoVis 2014 | 2024 ten-year InfoVis Test of Time | A membership matrix with count bars makes overlapping blocker sets readable. First-page figure inspected in the original PDF. [Paper](https://sci.utah.edu/~vdl/papers/2014_infovis_upset.pdf), [institutional publication and award record](https://vdl.sci.utah.edu/publications/2014_infovis_upset/) |
| Provenance-Based Visualization Retrieval, VAST 2017 | Best Poster | Searchable analysis-state provenance motivates an artifact lineage view with operation labels and reuse evidence. [Publication](https://www.vrvis.at/en/publications/PB-VRVis-2017-034), [official poster award list](https://ieeevis.org/year/2017/info/posters), [jury statement](https://www.vrvis.at/en/news-events/awards/ieee-vis-2017) |
| Vega-Lite, InfoVis 2016 | Best Paper | Concise coordinated panels and consistent visual mappings suit a combined method figure. [Paper](https://idl.cs.washington.edu/files/2017-VegaLite-InfoVis.pdf), [official award history](https://www.ieeevis.org/year/2023/info/history/best-paper-award) |
| Animated Vega-Lite, VIS 2022 | Specific design award unverified in this review | Local emphasis and visual/code correspondence remain relevant to the XML method examples. [Author project](https://vis.csail.mit.edu/pubs/animated-vega-lite/) |

HOLA and Vega-Lite figures were inspected in the preceding state-machine review.
The TensorFlow and provenance entries here use the primary publication and
recognition records. LineUp and UpSet add two directly inspected visual
references to that earlier review. Research PDFs and preview PNGs are retained
locally under `F:/corpora/state-machine-design-20261111`.

## Figure production rules

Use manuscript typography and vector geometry. A one-column figure should have
one question and a small worked example. Shared IDs and labels should survive
grayscale reproduction. Keep status symbols distinguishable from task identity
colors, show unknown values explicitly, and print units beside quantitative
axes. State which observations are measured, aggregated, reused or schematic.
Every proposed figure above can then be reviewed against its source contract.
