# State-machine figures

Reviewed on 11 November 2026. The source snapshot is Rust revision `1b9ee98`.
`inventory.json` records source paths, file hashes, enum roles and the transitions
used by the three manuscript figures.

## Research inventory

The verified awards concern scientific contributions, diagram-layout research
and long-term influence. The checked sources establish no award specifically for
the appearance of a paper's state-machine figures. The figure designs apply
readability techniques from the papers below. Their visual quality remains an
editorial assessment.

| Work | Verified recognition | Relevant visual lesson and evidence |
| --- | --- | --- |
| HOLA, InfoVis 2015 / TVCG 2016 | InfoVis 2015 Best Paper | Aligned nodes, restrained orthogonal routes and regular spacing. PDF Figure 1 inspected. [Paper](https://skieffer.info/publications/kieffer2016hola.pdf), [university award record](https://research.monash.edu/en/prizes/ieee-infovis-2015-best-paper-award) |
| STATEMATE, ICSE 1988 | SIGSOFT Impact Paper Award 2008 | Hierarchical state descriptions give a compact vocabulary for complex reactive behavior. Award record reviewed. [SIGSOFT recipients](https://www2.sigsoft.org/awards/impactpaper/) |
| Interface Automata, ESEC/FSE 2001 | SIGSOFT Impact Paper Award 2011 | Label transitions by events and preserve exceptional behavior explicitly. Original paper Figure 1 inspected. [Paper](https://web.cs.wpi.edu/~heineman/html/teaching_/CS562/p109-de_alfaro.pdf), [award record](https://www2.sigsoft.org/awards/impactpaper/) |
| The Immune System as a Reactive System, HCC 2001 | HCC 2001 Best Paper, Visual Languages and Formal Methods | Biological context is expressed through named reactive states and composition. Institutional abstract and conference recognition reviewed. [Institutional record](https://weizmann.elsevierpure.com/en/publications/the-immune-system-as-a-reactive-system-modeling-t-cell-activation-2/), [conference award list](https://conferences.computer.org/VLHCC/awards/) |
| A Framework for the Static and Interactive Visualization of Statecharts, JGAA 2002 | No specific award verified | Hierarchy, label placement, aspect ratio and removal of edge/label overlaps. Figures 8 and 9 inspected. [Paper](https://www.jgaa.info/index.php/jgaa/article/view/paper56) |
| InterState, UIST 2014 | No specific award verified | Put state labels and their behavioral context together. Figure 1 and its explanation inspected. [Author's paper page](https://from.so/papers/interstate_a_language_and_environment_for_expressing_interface_behavior_2014/) |
| Vega-Lite, InfoVis 2016 / TVCG 2017 | InfoVis 2016 Best Paper | Compact panels, consistent semantics and short adjacent explanations. Figures 2 and 3 inspected. [Paper](https://idl.cs.washington.edu/files/2017-VegaLite-InfoVis.pdf), [VIS award history](https://www.ieeevis.org/year/2023/info/history/best-paper-award) |
| Animated Vega-Lite, VIS 2022 / TVCG 2023 | No specific paper-design award verified | Local emphasis and concise visual/code correspondence. Figure 1 inspected. [Author's paper page](https://vis.csail.mit.edu/pubs/animated-vega-lite/) |

## Implementation inventory

The census includes all 78 Rust enum declarations in core/src, cli/src and the
workspace crates. Five are directly traversed parser/run machines. Four further
enums describe lifecycle/outcome vocabularies. Profiling phases, commands,
events, policies and classifications retain their own roles in the index.
Field-based protocols receive a separate targeted review.

| Owner | States or representation | Where it matters |
| --- | --- | --- |
| `models.rs::SheetStreamingState` | Six states, dispatched by `sheet/streaming.rs` | Worksheet root, columns, sheet data, merges and terminal worksheet end. Main SheetShrink figure. |
| `xml/namespaces/header.rs::State` | Prolog, Comment, ProcessingInstruction, Root with quote/cursor fields | Bounded eligibility screening. It does not validate the complete document. |
| `regression/streaming_xml.rs::State` | Text, Open, Tag, Comment, ProcessingInstruction, Cdata | Bounded XML tokenization, quoted attributes, partial delimiters and terminal acceptance. Main DeltaSieve figure. |
| `regression/office_plan/streaming_sheet.rs::Section` | Other, Columns, SheetData, MergeCells | Worksheet-context consumer layered over the token reader, with row/cell/depth fields. |
| `validation/run.rs::RunState` | Eight states with a central `permits` table | Durable orchestration, candidate-source bypass, publication and failure. Main DeltaSieve figure. |
| `validation_process.rs::ProcessStatus` | Pending, Running, Succeeded, Incomplete, Failed, TimedOut, Skipped | Per-invocation receipts. Running can be persisted before spawn, so a spawn error can leave a failed attempt with no child. |
| `validation_process.rs::StageState` | Preparing, Running, Completed, Incomplete, Failed | Stage-level receipt accounting over individual process outcomes. |
| `targeted_comparison/telemetry.rs::Lifecycle` | NotStarted, Opening, Open, Closing, Closed, Failed, Unavailable, Skipped | Native-lane observations. The enum alone imposes no global transition table. |
| `targeted_comparison/telemetry.rs::StepStatus` | Pending, Running, Done, Skipped, Failed, Unavailable | Native-operation observations. Heartbeats and process liveness remain distinct signals. |
| `sheet/hidden_content.rs::HiddenTail` | Current suffix, coordinate continuity, value and rejection fields | Row-local candidate admission, followed by dependency filtering. A derived admission model would require a separate explicit abstraction. |
| `xml/namespaces.rs::NamespaceGuard` | Boundary, declaration and quoting memory | Streaming representation guards after header screening. |
| `web/resident-connection.js` | Connection, pending request, deadline, retry and terminal-error fields | Authenticated reconnect and acknowledged results. One pending job request is enforced. |
| `archive/work_state.rs::Monitor` | Observed phase, event and wait fields | Diagnostic projection of work. These records do not enforce execution transitions. |

`CellStage` names timing buckets. `Phase` names operation telemetry columns.
`StateDecision` classifies whether a native state was observed, reused or disabled.
They are indexed without inventing transition graphs for them. UI view state and
vendored libraries are outside the machine inventory.

## Diagram contracts

- The worksheet graph uses all six enum variants and the eight state-changing
  edges from the dispatcher/tag mapping. Read continuation retains the current
  state. The red band summarizes returned errors and introduces no extra enum.
- The run graph preserves all thirteen permitted edges, including five failure
  edges summarized by the active-phase enclosure. The Processing self-edge is
  allowed by `permits`; it does not claim a transition event per processor stage.
  Supplied `source_b` inputs take the Preparing-to-Pairing path. Completed means
  finalized run accounting, with failed/limited pair observations still reported.
- The token-reader graph uses all six lexical states. Special constructs merge
  into their common Text successor. Quote memory remains inside Tag. The final
  acceptance predicate requires Text, zero depth and exactly one root. It is a
  return condition, not a seventh state.

The manuscript labels shorten implementation names. Full mappings and edge
conditions are retained in `inventory.json`. Source-review anchors include
`SheetXmlPreprocessor::process_current_state`, `streaming_state_for_tag`,
`RunState::permits`, `Run::transition`, `Scanner::feed` and `Scanner::finish`.

## Visual and verification choices

State names use the manuscript font. Blue marks ordinary states, teal emphasizes
payload work and accepted outcomes, and rose marks failure. Failure edges also
use dashes and labels. Terminal run states use an additional outline. Short
labels stay beside their edges, with paired routes for enter/return behavior.
The common style is maintained in `papers/shared/state-diagram-style.tex`.

The figures are native TikZ vectors. The existing algorithm figures remain in
place, with the state diagrams beside the corresponding method explanations.
The diagrams describe implementation behavior and introduce no new execution
or preservation guarantee. The build and focused verification receipts are
recorded with the final source change.
Focused verification passed 16 tests across five recorded commands. The source
review matched 6 worksheet states/8 state-changing edges, 8 run states/13 allowed
edges, and 6 token-reader states/9 state-changing edges. The acceptance/error
annotations retain their distinction from enum variants. `verification.json`
records commands and outcomes.
