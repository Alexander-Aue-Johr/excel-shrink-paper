# Preliminary streaming cell-comparison experiment

The receipts contain 63 samples across 21 cases, with three fresh processes
per case. Times are medians; reported process peaks are maxima across trials.
`manifest.json` identifies the inputs, compiler, measurement host and source hashes.
`samples.jsonl` retains phase timers and Windows memory counters.

The production source was revision `35cfd00` of the SheetShrink repository.
The test harness and runner were saved in commit `a8aedc5`. Their measured
copies are retained under `measured-code`. The harness uses the production
scanner, style resolver, record writer and merge without changing their behavior.
The runner expects the SheetShrink repository layout and a release test executable.

The independent oracle checks 6,561 small coordinate/value pairs. The focused
cell-stream, scanner and style-identity suites passed 27 tests. The opt-in
measurement test is run separately by the Python runner.

Normal filesystem caching remains enabled. Input preparation, decompression
and independent ElementTree checks are outside the timed process. Immediate
consumption and collecting `CellChange` records isolate result-retention costs.
They do not model the full pipeline's richer `XmlChange` evidence.

The red paper note marks the required repeat on the already documented
Ryzen 7 7800X3D configuration before finalizing the evaluation.
