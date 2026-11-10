# Method scope review

Both main manuscripts target twelve pages, including references and the
SheetShrink opening illustration. The review separates the scientific method
from operational instructions and development chronology.

## Reading sample

Five awarded IEEE/ACM works were selected for complementary methodological
questions. `paper-sample.json` records author-copy URLs, official award sources,
reviewed sections, local PDF hashes and short observations.

| Work | Editorial emphasis |
| --- | --- |
| Photon, SIGMOD | Design alternatives, execution model and memory behavior |
| Incremental Stream Query Placement, ICDE | Component workflow, state transitions and concurrency |
| Reverse Regret Query, ICDE | Problem definition, algorithm obligations and pruning |
| SAGA, SIGMOD | Objective, representation and logical/physical planning |
| Detecting Logic Bugs of Join Optimizations, SIGMOD | Oracle, generation, search and controlled evidence |

This small purposive sample informs the edits. It does not establish a universal
conference style. The algorithm, supported domain, necessary conditions and
measured consequences take precedence over lists of code locations or retries.

## Complete file inventory

`files.csv` lists every tracked path in the production tree, the separate capture
experiment tree and all other non-ancestor origin heads. Tracked paper scripts
are included as historical methods and research tooling. Ancestor branches are
listed in `summary.json`; their older code is not a separate current method.
Source files have content hashes, roles, families, line counts and lexical symbol
counts. `symbols.json.gz` holds the navigable name/line index. The extension list
and exclusions are explicit. Binary artifacts remain listed by Git identity.

The inventory covers 719 source files in the 928-file production tree. The wider
review contains 1,563 distinct source blobs across 15 tree snapshots. File
versions repeated across branches are not distinct implementations. Every source
family has an editorial decision in `decisions.csv`, and `checks.json` reconciles
the inventory counts. Tests, generated views and third-party libraries are
identified separately from production mechanisms.

This is a complete file and lexical inventory with family-level editorial
assessment. It is not a line-by-line correctness audit. Untracked temporary
scripts, installed dependencies and arbitrary older commits are outside its
scope. The production revision and all branch heads are pinned in the ledger.
No branch was merged and no native experiment was rerun for this editorial pass.

## Main-text changes

SheetShrink now explains the byte-state boundaries and retained buffers, compact
witnesses, row-aligned chunk work, workbook admission, precompression and writer
ownership. It identifies the CRC/length shortcut separately from exact identity.
The Wanderungen presentation case remains explicit. Fifty detailed examples
remain in the atlas, with three representative figures in the main method.

DeltaSieve now explains the fixed-size cell stream, style-component identity,
inverse conditional-reference mapping, all-state image qualification, independent
state controls, attribution and cache admission. Its code figures explain those
mechanisms. Transport schemas, UI behavior and long retry narratives no longer
occupy their own method subsections.

Measured experiments retain input population, endpoints, repetitions, failures
and scope. Raw case names, pinned revisions and detailed attempt evidence remain
in existing manifests and source-linked reports. The main text retains important
counterexamples where they establish a preservation condition, an oracle boundary
or a limitation. No final corpus-recall or universal speedup claim was added.

## Reproduce the ledger

```powershell
python scripts/paper-review/inventory.py --repository F:/repos/excel-shrink-rust --output NEW_DIRECTORY
```

The command reads current Git refs and writes only the selected output directory.
To reproduce this frozen ledger, use the revisions in `summary.json`. The symbol
index is lexical and intentionally does not claim a complete language AST.
