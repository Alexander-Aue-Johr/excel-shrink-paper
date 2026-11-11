# Related-work source review

Checked on November 11, 2026 against original publications, author-hosted
manuscripts, institutional publication records and official product descriptions.

## Integrated sources

- Mack et al., CHI Extended Abstracts 2018, DOI 10.1145/3170427.3174359.
  Forum case study used for performance motivation, with no prevalence estimate
  for redundant OOXML records. https://arxiv.org/abs/1801.03829
- Rahman et al., SIGMOD 2020, DOI 10.1145/3318464.3389782.
  Spreadsheet computation benchmarks and system optimization opportunities.
  https://experts.illinois.edu/en/publications/benchmarking-spreadsheet-systems/
- Bendre et al., ICDE 2018, DOI 10.1109/ICDE.2018.00020.
  DataSpread storage representations and positional access.
  https://arxiv.org/abs/1708.06712
- Bendre et al., SIGMOD 2019, DOI 10.1145/3299869.3319876.
  Anti-Freeze addresses interface responsiveness through asynchronous computation.
  https://people.eecs.berkeley.edu/~adityagp/papers/dataspread-async.pdf
- Budiu et al., PVLDB 2019, DOI 10.14778/3342263.3342279.
  Hillview uses distributed computation and visualization sketches.
  https://arxiv.org/abs/1907.04827
- Tang et al., ICDE 2023. TACO compresses formula graphs and queries the
  compressed representation. https://arxiv.org/abs/2302.05482
- O'Beirne, EuSpRIG 2013. The 2014 arXiv deposit is the same publication.
  The cleanup discussion on PDF page 19 identifies missing information about
  planned and completed removals. https://arxiv.org/pdf/1401.7586

## Scope decisions

The existing SheetReader 2023 journal reference remains the main ingestion
reference. Sigma Workbook and the later SheetReader demonstration were omitted
from the main narrative to focus on cleanup, storage, computation and ingestion.
General JSON-parser and data-repair paragraphs were replaced by closer
spreadsheet-system research. The formula-refactoring and validation connections
were retained.

Official Microsoft descriptions support the stated Check Performance and Inquire
policies. XLStylesTool's developer describes direct OOXML processing independent
of Excel's object model. Pre-native execution alone therefore has practical
precedents. SheetShrink's positioning combines persisted-context removal rules,
compact retained witnesses, execution cost and observed preservation outcomes.

The literature review adds contextual evidence. Performance and preservation
results for SheetShrink remain grounded in the separately retained experiments.

## SIMD implementation and methodological connection

The source inventory covers local revision 765c480 and fetched Origin revision
e738ca3. Core source files, manifests and Cargo.lock agree between these revisions.
Cargo.lock pins memchr 2.8.1, with the default std feature enabled. xml.rs caches
eight memmem::Finder instances through OnceLock. Rows and cells use this helper,
with direct calls in conditional_layout.rs and streaming.rs. The dependency's
x86-64 source supplies AVX2 and SSE2 dispatch, with pattern/input-dependent
substring strategies. Ordinary byte loops remain in attribute and delimiter
readers. This inspection establishes available implementation paths. The fraction
of end-to-end runtime attributable to SIMD requires a controlled ablation.

The simdjson reference is restored as a methodological connection to SIMD-aware
parsing. The Method identifies memchr as the actual search implementation.
The SIMD ablation request in the private SSH handover asks Gotham for matched
builds, verified dispatch control, repeated paired timings, resource measurements,
output equivalence and integration of completed results into Evaluation.
