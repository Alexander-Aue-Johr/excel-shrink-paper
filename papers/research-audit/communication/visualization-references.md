# Visualizing DeltaSieve's communication

Reviewed on 11 November 2026. This research inventory complements the
[implementation inventory](README.md). It maps established representations to
the observed protocol and distinguishes empirical visualization evaluation,
case-study experience and verified paper awards.

## Representation families

| Question | Established representation | DeltaSieve mapping | Information that must remain visible |
| --- | --- | --- | --- |
| Who sends what, and in which order? | Sequence diagram / time-space diagram | Coordinator, host, original add-in, variant add-in and launcher lifelines | Command ID, send/receive direction, acknowledgment, retry and concurrent work |
| Which event can cause another? | Happens-before graph / logical timeline | Command publication, dispatch, capture, upload, result commit and acknowledgment | Recorded causal links and actor-local order. A measured duration needs a physical-time scale |
| Where is elapsed time spent? | Span waterfall / Gantt swimlanes | Queueing, native geometry, capture, hashing, upload and host comparison intervals | Clock domain, overlap, missing measurements and ownership |
| Is a quiet worker alive, progressing or abandoned? | Aligned event and status tracks / timing diagram | Add-in heartbeat ticks, native process state, verified-progress counter, parent heartbeat age and deadline | Distinct signals, their observation source and the particular timeout they influence |
| What limits fast clients while peers are slow? | Queue/credit occupancy plot with a comparison frontier | Original and variant pending work, retained PNG bytes, upload/decoding and released credits | Explicit tile, pixel and byte budgets plus the point where credits are returned |
| Where do data and control travel? | Directed architecture/dataflow diagram | Small control messages, digest offers, PNG uploads, durable files and native-operation transport | Message type and actual routing through the host/coordinator |
| Which repeated failure patterns occur across attempts? | Event-sequence overview, aligned small multiples or transition matrix | Successful runs, reconnects, stale-parent cleanup, progress timeout and user restart | Cohort denominator, attempt identity, exceptional events and declared aggregation rules |

UML 2.5.1 provides standardized interaction notation, including sequence and
timing diagrams. Its normative specification is a reference for message and
lifeline semantics. The original Dapper paper and the SoCC visualization
perspective below document the role of hierarchical span timelines in tracing.
[UML specification](https://www.omg.org/spec/UML/2.5.1/About-UML),
[Dapper](https://research.google/pubs/dapper-a-large-scale-distributed-systems-tracing-infrastructure/),
[SoCC perspective](https://tjd45.github.io/assets/SeeItToBelieveItSoCC2022.pdf).

## Publications and recognition

The application column contains our design recommendations. Awards in the
evidence column retain their published category. Visualization studies evaluate
the specific tasks, participants and datasets used by their authors. Transfer
to DeltaSieve requires checking the communication and lifecycle semantics.

| Work | Representation and evidence | Application |
| --- | --- | --- |
| **ShiViz / Visualizing Distributed System Executions**, TOSEM 2020 | Interactive time-space diagrams, linked logs and structural comparison of executions. The paper reports a controlled experiment with 39 participants, a class study with 70 students and case studies with two developers. The controlled experiment found significantly more correct comprehension answers with the visualization. Figure 2 on PDF page 8 inspected. [Original paper](https://homes.cs.washington.edu/~mernst/pubs/visualize-distributed-tosem2020.pdf) | Strongest direct reference for command/result exchange and retry understanding. Collapse routine local events while retaining message boundaries and searchable IDs |
| **Ravel / Combing the Communication Hairball**, InfoVis 2014 | Logical timelines reveal message structure, wall-clock delay uses color and clustering reduces process count. Two parallel-code case studies demonstrate utility. Figure 1 inspected. Specific paper award unverified in this review. [Original paper](https://kisaacs.github.io/papers/isaacs_ravelganttchart_tvcg2014.pdf), [author project](https://kisaacs.github.io/) | Align capture rounds and result/acknowledgment boundaries. Keep a physical-time detail view when diagnosing timeout duration |
| **Perfopticon**, EuroVis 2015 | Coordinated execution-plan, communication and per-worker trace views. Developer, scientist and student engagements supply practical evaluation, including debugging bottlenecks. Figure 1 inspected. Specific paper award unverified in this review. [Original paper](https://idl.cs.washington.edu/files/2015-Perfopticon-EuroVis.pdf) | Link a compact communication topology to the currently selected original/variant capture batch and its queue/byte history |
| **Dapper**, Google 2010 | Hierarchical trace spans and production-use experience. The paper explains causal and temporal relationships and the trace-inspection interface. Specific visualization award unverified in this review. [Primary publication](https://research.google/pubs/dapper-a-large-scale-distributed-systems-tracing-infrastructure/) | A measured waterfall can separate Office.js work, upload and comparison latency if corresponding intervals are retained |
| **TraVista / Aggregate-Driven Trace Visualizations for Performance Debugging**, 2020 vision paper | Extends a single-trace Gantt view with metric, temporal and structural aggregate context. The paper presents an initial design. Specific award unverified in this review. [Author PDF](https://stolet.github.io/papers/travista.pdf) | Compare one slow attempt with a declared cohort of comparable operations. A distribution beside an interval helps explain whether its duration is unusual |
| **Temporal Event Sequence Simplification / EventFlow**, VAST 2013 | Verified **Honorable Mention**. Before/after overviews and task-driven simplification reduce event complexity. Figure 1 inspected. [Original paper](https://www.cs.umd.edu/users/ben/Monroe2013Temporal.pdf), [official conference recognition](https://ieeevis.org/year/2013/info/overview-amp-topics/ieee-vast-papers-30-second-videos) | Aggregate repeated normal heartbeats in a long-run overview while preserving counts, observation windows and failures. Retain the full event evidence for inspection |
| **Simplifying Overviews of Temporal Event Sequences**, CHI 2016 Late-Breaking Work | Verified **Best Paper Honorable Mention in the Late-Breaking Work category**, recorded by the authors. Two formative usability studies informed pair selection, ranked category pairs and a category matrix. Figure 1 on PDF page 2 inspected. [Original paper](https://www.cs.umd.edu/users/ben/Mauriello2016Temporal.pdf), [author's award poster](https://www.eecis.udel.edu/~mlm/docs/2016-Mauriello-CHI-Choose2-Poster.pdf) | Begin with a useful signal pair, such as heartbeat and verified progress. An overview matrix can help choose subsequent pairs without drawing all event categories at once |
| **Sequence Braiding**, TVCG 2021 | Layered temporal-event overview. Its comparative user study reports faster task completion overall, with benefits depending on task type. Specific award unverified in this review. Evaluation text reviewed through the primary archived manuscript. [Primary manuscript](https://par.nsf.gov/servlets/purl/10200661) | A later cohort view could reveal common reconnect/cleanup paths. It needs attempt-level event sequences and declared aggregation semantics |
| **Pivot Tracing**, SOSP 2015 | Verified **Best Paper Award** for dynamic causal monitoring and happened-before joins. The recognition concerns the systems contribution. [Institutional publication and award record](https://www.microsoft.com/en-us/research/publication/pivot-tracing-dynamic-causal-monitoring-for-distributed-systems/) | Supports retaining correlation across actors so one operation can be examined through several signals. Its award supplies context for causal monitoring research |

The verified visual-analytics awards here concern the event-sequence research
and its interface contributions. ShiViz supplies the closest empirical evidence
for distributed-execution comprehension. The checked records establish no
separate award for the aesthetic appearance of these protocol figures.

## A paper figure for DeltaSieve

Use three complementary panels with manuscript typography and a restrained
shared palette.

### A. One capture round

Use a sequence diagram with five actors, ordered as original add-in, resident
host, coordinator, variant add-in and native launcher. Keep routing explicit
even when the host serves both add-ins. Show startup/readiness as a compact
prerequisite, then the command, finalized PNG hash offer, optional retained PNG
upload, durable result and acknowledgment. A small inset can show acknowledgment
loss followed by replay of the same result. Command and capture IDs bind the
exchanges. Label the diagram as a protocol illustration.

ShiViz motivates explicit event/causal links and linked detail. Ravel motivates
aligning repeated rounds to expose structure. Put transport retry around the
affected exchange and identify the saved result that survives the retry.
[ShiViz](https://homes.cs.washington.edu/~mernst/pubs/visualize-distributed-tosem2020.pdf),
[Ravel](https://kisaacs.github.io/papers/isaacs_ravelganttchart_tvcg2014.pdf).

### B. A supervision example

Use a shared physical-time axis and separate tracks for add-in heartbeat,
native process alive, verified progress, parent-runner heartbeat and cleanup.
Ticks encode observed heartbeats. A stepped progress counter shows confirmed
work. Mark the particular expiring deadline and the owner that initiates
cleanup. A healthy heartbeat with a flat progress counter is an important
case. An expired parent heartbeat is another case with a different trigger.
The source inventory defines the actual shutdown conditions. The visual
combination is our proposal.

### C. Bounded capture ownership

Use two short stream lanes converging at the comparison frontier. Give each
capture a stable identity across captured, retained, uploaded, compared and
released positions. Plot retained PNG bytes or pending pixels on a separate
quantitative axis. Mark the credit limit and its return point. This explains
why a fast add-in waits for its peer and why memory remains bounded. Color can
encode actor or capture identity, with status conveyed by shape and labels.

Perfopticon supports linking a system-level view to local execution and
communication. TraVista supports contextualizing one interval with aggregate
measurements when comparable samples exist.
[Perfopticon](https://idl.cs.washington.edu/files/2015-Perfopticon-EuroVis.pdf),
[TraVista](https://stolet.github.io/papers/travista.pdf).

## Trace and layout requirements

* A logical sequence can use recorded command, client and receipt identities.
  A physical-time chart needs intervals and a declared clock domain.
* Preserve publication, dispatch, native completion, upload receipt, result
  commit and acknowledgment as distinct observable events where available.
  Existing aggregate timings only support their retained granularity.
* Heartbeat ticks, process probes and verified-progress increments require
  separate tracks and explicit missing-data marks.
* Routine repetition can be collapsed with an event count and interval. Keep
  failures, session retirement, retry boundaries and timeout decisions visible.
* A measured topology edge needs its quantity stated. Message count, payload
  bytes and duration have different units. Buffered-byte occupancy is a level
  over time and needs its own axis.
* Parallel actors need enough space for branch boundaries and acknowledgments.
  Use stable ordering and small multiples when added lanes would obscure paths.

These are design requirements derived from the reviewed literature and the
local communication inventory. Implementation-specific timing and buffer
claims remain tied to retained observations.

## Integrated figure

`papers/deltasieve/figures/communication.tex` implements the three panels as a
native TikZ figure in the method section. Its adjacent PDF is a standalone
vector export generated by `scripts/diagram-design/export.py`. The message
sequence abstracts startup and lower-level upload acknowledgments. Panel B
illustrates parent-runner heartbeat expiry independently of add-in reachability.
Its elapsed times and panel C's occupancy trace are schematic examples.

## Review artifacts

Five original PDFs were downloaded and rendered for figure inspection under
`F:/corpora/communication-design-20261111`. Reviewed pages are ShiViz 3 and 8,
Ravel 1, Perfopticon 1, EventFlow 1 and Simplifying Overviews 2. Original PDFs
remain in that local research directory. Links above identify their primary
publication sources. The remaining references were reviewed through primary
publication pages or manuscript text, as specified in the table.
