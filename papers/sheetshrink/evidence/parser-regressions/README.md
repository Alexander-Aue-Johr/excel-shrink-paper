# Parser regression evidence

The source census pins SheetShrink revision `1b9ee98` on 11 November 2026.
`commit-messages.json` retains all 1,140 messages reachable from the recorded local
refs. `rust-tests.json` indexes all 1,403 tracked Rust test declarations by path,
line, name and source-file SHA-256. This is a full source census. It includes
ignored tests and declarations outside the parser. It does not count fixture
iterations as separate tests or assert that every test was executed.

## Counts used in the paper

| Source modules | Test declarations |
| --- | ---: |
| Core XML, namespaces, strings, ranges, names and streaming | 23 |
| Diagnostic streaming reader | 8 |
| Core cell, row and tail regression suites | 156 |
| Remaining core unit tests | 120 |

`manuscript-counts.json` identifies the 31 parsing-module correctness tests and
155 cell/row/tail correctness tests used in the prose. Timing benchmarks are
excluded from these manuscript counts. Other modules contain additional parsing
and preservation controls.

The core total is 299. One of the row tests is an ignored manual timing benchmark.
The remaining workspace tests cover validation, orchestration, acquisition and
other components. Module groups describe the source organization. Several tests
exercise both lexical correctness and preservation semantics.

`execution-results.json` records fresh Windows execution receipts with 298 core tests passed with one ignored,
eight diagnostic streaming tests passed, and 39 pipeline tests passed. The core
and CLI commands use `--lib`. No complete corpus campaign was rerun for this edit.

## Findings checked against history and fixtures

- `88c77f9` fixes actual Data.gov content loss from attributed `<v>` elements.
  The supplier-diversity investigation records a removed heading and strings.
  This is the most recent identified lexical content-classification correction
  in the recorded history. Later font and layout fixes address preservation rules.
- `5822800` covers defined-name scope and a self-closing name followed by a paired
  element. The fixture checks that the following element remains intact.
- `c608d4e` requires the terminal worksheet state and prevents publication after
  incomplete parsing. Its integration fixture checks partial-output cleanup and
  preservation of an existing destination. This proves a guarded error path.
- `6caae0f` fixes CDATA recognition in diagnostic XML readers. It concerns the
  validation reader, with separate cleanup passthrough controls.
- `2703f96` compares fused cell boundaries with a reference implementation over
  truncated prefixes and self-closing cells. This records boundary regression coverage.

The review also traced the cell/row/tail suites, reference-scanner, output,
shared-string and namespace tests. Their names and source identities are indexed
without duplicating their source code in this artifact.

## External baseline finding

`sheetreader-recheck.json` retains the targeted baseline/fixed outcomes for
`Opt!O306`. The tested input is an XLSX conversion of
`native_000/3.307869.GDP5KHQBLU2E3JIJKHMQB0USQH05OWZIB.1.xls` from the
[EDRM archive](https://archive.org/download/edrm.enron.email.data.set.v2.xml/edrm-enron-v2_quigley-d_xml.zip).
The submitted [SheetReader correction](https://github.com/polydbms/sheetreader-core/pull/5)
is commit `49c7301f97fe7d272412735547a304c015f00370`, against upstream
`1e7698531199bb7dea29fe63fd61e5525cb1a021`. Its status at this cutoff is open.

## Reproduction and corpus scope

Run `python scripts/paper-review/inventory_parser_regressions.py CHECKOUT OUTPUT`
to regenerate the census. Use the pinned checkout and refs for matching counts.
The extractor recognizes ordinary Rust test attributes. Generated test cases and
non-Rust runners are outside the declaration count.

`full-corpus-distribution.json` is the upstream receipt at revision `ce3722cb`,
from `docs/visual-validation/experiments/full-corpus-distribution-20261111.json`.
Producer `e7b91538` completed the three supplied-XLSX populations with zero
reported processing failures. It records 1,182 Destatis, 8,623 Data.gov and
30,301 prepared Enron outputs, plus 494 Data.gov namespace exclusions.
The earlier `../cleanup-screening.json` describes historical ranking runs,
including recovered outputs and a failed retry. The paper now cites the fresh
release runs. Corpus-wide visual equivalence remains unestablished.

## Namespace follow-up

The prefixed-data experiment at Rust revision `44d73f3` inventories 493 actual
prefixed-data matches among 9,117 Data.gov workbooks. This selection predicate
explains its scope alongside the 494 production namespace exclusions. All 493
normalization/cleanup trials completed. Two compressed outputs became smaller,
by 6.415% and 0.047%. The maximum row coordinate was 694,981. The inventory
therefore supplies no evidence of exceeded Excel row limits in these 493 cases.

The historical manuscript `excel-shrink-paper.tex`, line 296 at revision
`f7f6f8b3`, associates the inspected prefixed files with machine-generated
spreadsheets. This is an exploratory observation. There is no per-file producer
audit establishing the generator of all excluded workbooks. Microsoft also
documents Excel's Strict Open XML output option. The paper retains the measured
savings, scoped observation and production eligibility policy.
