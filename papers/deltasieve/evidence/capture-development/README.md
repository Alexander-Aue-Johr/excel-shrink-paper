# Capture development evidence

The review used the pinned sources in `sources.json`. It added no native timing
runs. The experiment branch remains separate from the production branch.

| Question | Retained finding | Interpretation |
| --- | --- | --- |
| VBA and GDI+ | Clipboard picture, bitmap/EMF and owned-buffer pair adapters exist | Separate raster and coverage policies prevent a universal speed ranking |
| VSTO and Excel-DNA | Drawing geometry adapters exist | Their observations do not measure screenshot throughput |
| Local host | Five compatible historical pairs had median Rust/.NET workload ratio about 3.34 | A result about those host implementations and endpoints |
| Image budget | Both controlled 18-case trials select 32 MP under the predeclared 5% rule | 64 MP was fastest in both trials and was not rejected as unstable |
| Historical replication | 48 complete and 14 failed out of 62 nominal cases | Four unresolved worksheet outcomes account for repeated failures |
| Parallel pairs | Two pairs improved capture throughput with higher measured job memory | Startup and RAM admission limit the inference to the completed configurations |
| Image completeness | A million-row request returned a truncated PNG without an API error | Verify expected dimensions before accepting image evidence |
| Geometry batching | All 64 returned geometries matched across scalar and batch runs | Polling and warmup remain part of the recorded timing scope |
| Adaptive partitions | FRPP geometry yielded 64 leaves after 710 probes | This was a planner observation, not a completed paired image comparison |

## Reproduce individual historical experiments

Read the pinned `experiments/README.md` and `PROTOCOL.md` before execution.
The historical branch has a CLI as well as its GUI. `reproduce.ps1 -PlanOnly`
resolves its declared blocks. `run.ps1 -InterfaceMode cli` accepts a named
experiment. Native prerequisites, input hashes and scope remain required.

The historical configuration used selective overflow interventions and disabled
gridlines. Its measurements must retain those settings. Current capture policies
are described separately in the production behavior contract. Replaying an old
protocol is not a production-policy benchmark.

A wholesale merge would import the old suite and its shared hosts alongside
later production changes. A future extraction should preserve one runnable
experiment and its own prerequisites per change, starting with drawing bounds,
VBA/GDI+ controls, then the paired capture matrix. Keep the existing branch as
the reference until the extracted versions reproduce their declared outcomes.

The manuscript adds four abbreviated code panels for dependency selection,
all-state geometry, Office.js synchronization and image qualification. They are
illustrations linked to source anchors, not alternate runnable implementations.
