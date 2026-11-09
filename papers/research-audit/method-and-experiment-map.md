# Method and experiment evidence map

## Revision boundary

Current implementation paths refer to software revision `a7eef1a`. Historical
experiment paths refer to `origin/experiments/visual-validation` at `ce0dc7b`.
The branch and artifact inventories in this directory preserve the wider search
frame. Development controls, completed experiments and ongoing corpus work have
separate evidential roles. This map records retained examples, without claiming
that every regression has an original workbook attached.

## SheetShrink

| Rule family | Specification and implementation | Retained provenance and manuscript use |
| --- | --- | --- |
| Serialization eligibility | `docs/behavior/shrinker.md` S1–S3; `excel-shrink-cli/src/pipeline.rs`, `pipeline_namespace_tests.rs`; `excel-shrink-core/src/xml/namespaces/` | Unprefixed Transitional XML and supported encoding form the eligible frame. Strict, prefixed core structures and unsupported encoding retain exclusion outcomes. Method describes the restricted parser contract. |
| Content and inherited styles | S4; `excel-shrink-core/src/sheet/cells.rs`, `cells_regression_tests.rs`, style analysis | Numeric zero and empty shared-formula followers remain content. Enron `44ad86f1d9f249470fc6`, AUGUST, distinguishes a no-fill override from a tail anchor. Default style can mask inherited fill or border. |
| Font witnesses and row promotion | S5; `rows.rs`, `rows_regression_tests.rs`, `row_view.rs` | Enron `005d034f18d6dc340a2c` exposes a one-pixel row-height change around protected Arial cells. Destatis InformationenzurStatistik A6:H6 supplies a merged-font witness. AFFF Manual Data rows 124–128 require explicit blanks over inherited fill. |
| Tail columns and geometry | S6; `ranges.rs`, `cols.rs`, `drawings.rs` | NLA water-isotope case retains a column-font witness because replacing Calibri 11 with Arial 10 changes 15.75-point rows to 15 points. Hidden columns, borders and merged bounds constrain pruning. |
| Hidden constant tail | S7; `tail_fill_regression_tests.rs`, range/reference analysis | Dense horizontal hidden integer tails ending at XFD require at least 4,096 equal values, with layout and reference guards. Ordinary hidden tables and formula/text/gapped/vertical tails remain outside the admitted rule. |
| Reference safety | `docs/methodology/reference-protection.md`; reference scanner, shared-formula and object-reference modules and regressions | Intersecting known references retain the span. Unresolved dynamic references take the conservative path. Source-workbook links exist for individual regressions, with uneven completeness. |
| Cross-coordinate visual effects | S5/S6 follow-up cases and CLI pipeline regressions | Data.gov `Fig10_New10-29-19` D2:G2 requires the center-across-selection span. `PnP Sector Boilers Dataset`, Fig 4 SCR, K6 masks a blue row fill. Row-level borders require empty-column widths. Linked chart number formats can require an empty series anchor. Added to the method draft. |
| Package bookkeeping and failures | S8–S10; workbook name and calculation-chain handling, pipeline regression tests | Built-in name aliases match name and localSheetId without general formula-equivalence checking. No general table renumbering. Partial processing cannot establish successful output. Excel-hosted sentinel replacement remains prospective. |

## DeltaSieve

| Mechanism | Retained evidence | Interpretation |
| --- | --- | --- |
| Effective styles and dependencies | `docs/behavior/targeted-validation.md` V5–V11, reference and selection tests | Style IDs are package-local. Selection follows supported conditional and chart dependencies. Unsupported formulas, effects and scopes retain declared fallback outcomes. |
| Tall-sheet narrowing | `src/regression/office_capture/tall_sheet_analysis.rs`, `tall_sheet_selection.rs`; investigation `datagov-watch-20261104.md` | Row-local conditional inputs distinguish 404 populated rows from million-row application scopes. The PUF control compares two logical cells through a broader merged native area. Its 305 ms selection/transport and 775 ms comparison timings do not measure a matched full-capture speedup. |
| Native geometry | Same investigation, uniform-height case | Accepting the API's uniformity signal avoids repeated subdivision caused by rounded aggregate heights, 51.5 versus 51.5625. The retained six-capture result is a development control. |
| Unsupported-object cost | `investigations/native-legacy-object-bounds-20261109.md` | Elterngeld and Abwasser remained active beyond five hours because whole-grid fallback produced 104,836 and 185,942 regions. These are censored observations. Zero-height line removal changed 1,042 pixels. A planning improvement still lacks a completed matched cost trial. Added to evaluation. |
| Batching | `experiments/batched-capture-geometry/{README.md,results.json}` | Matched geometries and decoded images under alternative package sizes support a bounded execution observation. Startup excluded, 250 ms polling, sequential package trials and background work limit inference. Added to evaluation. |
| Opening | `experiments/parallel-workbook-open/{README.md,findings-four.json}` | Independent process startup can overlap. Two reversed-order rounds support the reported four-workbook medians. COM command parallelism and corpus throughput remain unmeasured by this experiment. |
| Ownership and recovery | `resident-events.md`, `cleanup-isolation.md`, `execution-and-evidence.md` | Durable receipts, duplicate handling, budgets, confirmed exit and per-group profiles make failure accounting inspectable. Clean-desktop controls establish a specific shared-profile cleanup hazard. |
| Registration accumulation | `investigations/addin-registration-accumulation-20261104.md`, `addin-startup-20261108.md` | Missing registration accumulation is documented, with reversible controls. Restoring registrations also produced a fast startup. Cleanup alone therefore has no established universal causal effect. |
| Transport | `resident_events.rs`, `resident-connection.js`, `live_stream.rs`, `live-status.js`; live-stream experiment | WebSocket commands coexist with HTTP images and durable files. Live status uses a restricted JSON Patch subset and sequence validation. Synthetic status-byte reductions describe that workload. No integrated MCP server was found. |

## Historical alternatives and retained limits

The legacy WPF launcher is preserved in `docs/visual-validation/experiments/`.
Its protocol separates coverage, host/overflow and pixel-budget experiments.
`tools/excel-lab` retains VBA CopyPicture, bitmap/EMF, memory-pair and row-height
adapters, PowerShell COM control, GDI-backed pixel comparison and a Rust native
comparison reference. Clipboard use is an observable design constraint. No
retained trace was found that establishes the recalled clipboard corruption or
reboot requirement. That recollection cannot support a measured paper claim.

Excel-DNA and VSTO experiments measure drawing bounds. Their scope does not
support ranking complete screenshot engines. Rust/.NET host comparisons include
communication and telemetry. A five-pair screening median ratio of 3.34 concerns
those implemented workloads. The 62-case replication completed 48 cases and
failed 14, with repeated drawing-bound outcomes. Four distinct unresolved sheets
explain the coverage gaps. Completed coverage spans 462 of 466 requested sheets.

The repeated 16/32/64-megapixel budget trials favor 64 MP on elapsed time. The
selected 32 MP budget was within 5% of that time on the two measured scopes.
It must not be described as the fastest or as proof that 64 MP is unreliable.
These results can support the design discussion or artifact. They cannot supply
a controlled universal VBA/COM/Office.js ranking.

The Rust-COM shape-bounds adapter pins `excel-api` commit
`8348746310ffab65dc0355b67e25d0040d4d0a37` on `fix-shape-variant-types`.
The retained fix accepts VT_R4 single-precision numeric properties alongside
VT_R8 and integer representations. Shape.Visible's MsoTriState conversion is a
separate unresolved adapter caveat. Opening overlap provides no evidence that
all COM calls execute concurrently. The fixed commit is inspectable at
<https://github.com/Alexander-Aue-Johr/excel-api/commit/8348746310ffab65dc0355b67e25d0040d4d0a37>.

## Historical paper and acquisition reuse

The retained analyzer uses lxml document parsing. Historical cleanup scripts use
regular-expression extraction. The current byte-oriented implementation has a
different rule set and preservation contract. Their timings are historical
observations until exact producer, input and environment identities are
reconciled. Feature-equivalent comparative speedup remains unestablished.

The preparation snapshot `98c4e9d` preserves source, versions, receipts and hashes
for recursive Enron inspection. Its `code/` pipeline and older `paper-reference/`
downloader have different checkpoints and entry points. Reuse the frozen source
selection and content-addressed preparation receipts for methodological
provenance. Integrating both implementations into the shrinker is unnecessary
for manuscript preparation. Old scraper URLs alone cannot identify a frozen
evaluation corpus. Versioned current source profiles supply member hashes.

`plot_retained_wanderungen.py` reads only the retained CSV through Python's csv
module. It never imports workbook loaders or the old benchmark script. Its log
axis preserves visibility of short completed durations alongside the 690.1 s
timeout. The timeout remains a lower-bound marker. The figure excludes cleanup
time and does not combine either historical Rust row into a current pipeline
speedup. The separate more-than-48-hour attempt awaits its report and setup.

The historical figure is an editorial preview outside the manuscript pending
provenance reconciliation. This keeps the current evaluation tables tied to
identified producers and inputs.
