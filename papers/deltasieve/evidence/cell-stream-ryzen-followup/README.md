# Ryzen-host cell-stream follow-up

All 63 fresh-process release trials across 21 cases completed and matched their
expected change counts. This lane repeats the preliminary artifact protocol on
an AMD Ryzen 7 7800X3D with sixteen logical processors and 102,066,212,864 physical
bytes. The compiler is rustc 1.96.0 with LLVM 22.1.2. The earlier Intel i7-8700 and
rustc 1.99.1 lane remains separate. Cross-host observations are not pooled.

protocol.json retains exact input, source and executable hashes, repetitions,
normal-cache policy and the completed owned-job guard receipt. samples.jsonl.gz
contains all 63 individual phase/resource observations with a deterministic gzip
checksum in the protocol. summary.json retains medians, merge ranges and all
working-set peaks per case. Input paths are projected to basenames. Source
snapshots under measured-code retain the actual measured bytes without Git text
conversion. Reproduction uses the core repository at revision 366c7e75 and the
included runner, with the compiled release test executable identified by SHA-256.
The complete run used a 4 GiB aggregate Windows-job commit ceiling and 8 GiB
physical-memory and commit-headroom reserves. All descendants exited.

For one million densely changed cells, immediate consumption versus retaining
CellChange records peaks at 5.703 versus 94.973 MiB working set, a 94.0% reduction.
This isolates result retention, excluding richer full-pipeline evidence. One
3.06-million-cell Enron repetition peaks at 42.121 MiB working set, while its
other repetitions peak at 6.105 and 6.078 MiB. The outlier remains in all maxima
and receipts without causal attribution. Working set and private commit have
different meanings. The independent oracle and fixed mutations do not establish
full validation recall. Concurrent native comparisons and retained filesystem
caches qualify the timings. ZIP decompression and preparation precede the timers.

The frozen release executable also passed 27 focused correctness tests, including
6,561 ordered small-grid pairs against an independent coordinate/value oracle.
The three text receipts retain the exact successful suites and the ignored
measurement entry point. The latter is executed separately for each timed sample.

The standalone timing runner does not enforce the outer memory guard. Its exact
MemoryJob and OwnedJob source hashes reference the companion frozen SheetShrink
resource artifact. Apply that guard before launching the runner. The private
outer supervisor hash is retained separately, without publishing local paths.
