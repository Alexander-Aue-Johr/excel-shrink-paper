# Enron opening and saving follow-up

Eighteen bounded openpyxl 3.1.5 trials cover one selected Enron workbook,
six variants and three repetitions. The configured process-tree commit limit
is 4 GiB, with 8 GiB system physical-memory and commit-headroom reserves.
All endpoints have terminal receipts. Original and XLStylesTool reach the
memory limit in all repetitions. The other four variants complete.

Prepared external variants measure the library round trip. SheetShrink measures
fresh cleanup followed by that round trip, from process creation through exit.
These different endpoints do not support an end-to-end ranking of cleaners.
Memory-limit endpoints receive no completed-runtime speedup. Reported sampled
commit peaks are retained exactly, including small overshoots of the configured
limit. Native comparisons ran concurrently and filesystem caches were retained.

The deterministic gzip contains path-free trial receipts. Its checksum is in
protocol.json. Private traceback and command paths are omitted. The eight
measured source hashes reference the frozen sibling bounded-open-save-followup
source snapshot. Summary times use wall_seconds, including process startup.

After the timed series, a separate fresh cleanup matched every uncompressed
member name, byte length and SHA-256 of the retained native SheetShrink candidate.
The 18-member control is recorded in fresh-native-join.json. Timed output archive
hashes are retained, but individual timed member inventories were not retained.
This representative control does not establish per-trial member equivalence or
general cleanup idempotence. Native coverage qualifications remain in the
separate visual audit.
