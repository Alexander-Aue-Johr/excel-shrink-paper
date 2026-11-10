# AI disclosure and submission checks

Checked 10 November 2026 for the planned ICDE 2027 research-track submission.
The historical repository branch name does not identify the submission year.

## Applicable requirements

ICDE requires an acknowledgment naming the AI systems, affected sections and
scope of assistance. Its examples include data processing, visualization and
experimental workflows. The twelve-page limit excludes references and the AI
declaration. AI cannot be an author. Responsibility remains with the authors.
A model-version inventory, prompt archive or numerical limit on AI assistance
is not explicitly required by the reviewed instructions. No list of forbidden
phrases or automatic acceptance wording was found. [Research CFP](https://icde2027.github.io/cf-research-papers.html),
[submission guidelines](https://icde2027.github.io/submission-guidelines.html).

The two official pages disagree about the one-year resubmission interval.
The CFP measures it from rejection, while the guidelines measure it from the
last submission. The project records a related earlier ICDE rejection. The
author must establish the relevant dates and ask the chairs whether the planned
submission is eligible. Revised versions are expressly covered by the CFP.
The declaration cannot resolve this eligibility question. [Research CFP](https://icde2027.github.io/cf-research-papers.html),
[submission guidelines](https://icde2027.github.io/submission-guidelines.html).

## Published wording examined

- Langlet et al., *Spatiotemporal Sketch Disaggregation*, author-hosted ICDE 2026
  paper, PDF page 13, identifies Grammarly, ChatGPT and Gemini for minor
  sentence-level edits. Its narrow editing-only declaration would understate
  this project's use. [Author PDF](https://zaoxing.github.io/papers/2026/ICDE26_DiSketch.pdf).
- Schmidt et al., *Who Watches the Watchmen?*, Findings ACL 2026, PDF page 20,
  separates writing, code generation, refactoring, model names and human
  verification in an AI Usage Card. This separation informed our prose.
  Its assertion that all outputs were manually verified was not adopted.
  [Published paper](https://aclanthology.org/2026.findings-acl.1145.pdf).

Both manuscripts identify ChatGPT and OpenAI Codex and disclose drafting,
literature assistance, coding, scripts, figure-generation code, evidence analysis
and overnight agent support. The statements about author direction, source review,
manual testing and debugging follow the author's account in this task. They do
not certify exhaustive review of every generated output.

Exact historical model names remain an author check. The official model label
is GPT-6.1 Sol. Light is a reasoning setting in the reviewed documentation.
The label Lite has not been established as an exact model identity. Product names
are used in the declarations to avoid inventing a historical model inventory.
Local agents here describe local automation connected to AI services, without
asserting local model inference. [Official model documentation](https://learn.chatgpt.com/docs/models?surface=cli).

## Interface verification

Inspected SheetShrink source revision
`1b9ee98e3409c3584ffd12a661d0c19f8a5f0788`.

- `excel-shrink-cli/src/capture_host/control_room.rs` registers local HTTP routes,
  including `/api/summary`, `/api/workbooks`, `/api/workers`, group reports,
  evidence access and explicit control requests. Local Host/Origin checks apply.
- `control_room/summary.rs` bounds diagnostic projections and distinguishes
  current attempts, process liveness and verified progress. The normal summary
  limit is eight, with a maximum of fifty.
- `control_room/live_stream.rs` provides WebSocket snapshots and incremental
  status updates. It is separate from resident Office.js worker commands.
- A production-source search for MCP, Model Context Protocol, JSON-RPC and
  `tools/list` or `tools/call` found no MCP implementation. The roadmap proposes
  a future adapter in `docs/planning/sheetshrink-deltasieve-roadmap.md`.

The implemented interface is application-specific HTTP/JSON and WebSocket.
The manuscript makes no MCP implementation claim. MCP defines its own tool
messages and protocol semantics. An HTTP endpoint alone does not establish
conformance. [MCP tool specification](https://modelcontextprotocol.io/specification/2026-07-28/server/tools).

The monitoring paragraph uses the frozen manuscript inventories. The 34,065
Enron entries are content-addressed table candidates. They must not be described
as 34,065 successfully converted or visually validated workbooks. Data.gov and
Destatis contain 9,117 and 1,182 XLSX inputs respectively. None of these counts
establishes completed comparison coverage. No uninterrupted monitoring coverage
or measured reduction in human effort is asserted.

## Author checks before submission

- [ ] Read and approve the final manuscripts, figures and AI declarations. Confirm
  the stated human review and actual AI systems against the session records.
- [ ] Check every cited source, numerical claim and originality claim against its
  primary evidence. Remove unsupported conclusions and fabricated references.
- [ ] Review the consequential AI-assisted rewrite and selection logic and the
  test expectations. Confirm independent controls and manual debugging evidence.
- [ ] Freeze inputs, binary/configuration identities, environment and reporting
  cutoff. Keep failures, unresolved findings and partial coverage explicit.
- [ ] Inspect every visual finding used to support a scientific conclusion.
  Resolve agent-proposed regressions against native captures and comparison
  records. Do not treat agent annotations as independent ground truth.
- [ ] Complete the remaining independent reference, mutation and matched-cost
  experiments stated as unfinished in the papers, or retain bounded claims that
  accurately describe those omissions.
- [ ] Resolve prior-rejection eligibility with the chairs. Verify companion-paper
  overlap, concurrent-review status, correct track, ORCID and CMT conflicts.
- [ ] Check the final body-page budget, unchanged IEEE formatting and accessible
  reproduction artifacts. EAB submissions require all reproduction artifacts.

An editing-only disclosure despite substantive generation, an unsupported claim
that every output was manually verified, or an AI-only correctness assertion
would be inaccurate here. These are factual risks rather than prohibited strings.
A truthful disclosure does not guarantee scientific acceptance.

## Verification in this change

Both manuscripts were rebuilt with TeX Live. The body-end labels occur on page
11, and both PDFs contain twelve pages including the declaration and references.
The final build logs contain no undefined-reference or overfull-box warnings.
The acknowledgment and new method pages were rendered for visual inspection.
Source interfaces were inspected without changing or restarting corpus jobs.
No new native validation campaign or independent scientific audit was performed.
