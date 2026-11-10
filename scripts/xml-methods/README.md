# XML method figures

[Review atlas](../../papers/sheetshrink/figures/xml-methods/xml-cleanup-atlas.pdf)
contains 50 column-width code comparisons, native controls, schematic failures
and a source index. Nine selected figures are included in the manuscript.

## Rebuild

Install `reportlab`, `Pygments`, `Pillow` and `pypdf` in a local Python environment.
Then run from the repository root.

```powershell
python scripts/xml-methods/render.py
python scripts/xml-methods/verify.py
```

The LaTeX build consumes the generated PDFs and needs no Python or shell escape.
PDF text is selectable. SVG text and its font files are embedded for editing.
The bundled DejaVu and Tinos fonts retain their licenses in `fonts/`.

## Read the examples

The editable inventory is
`papers/sheetshrink/evidence/xml-methods/cases.json`. Each case pins a source
file, test or implementation anchor, line, commit and file hash to the reviewed
Rust revision. Style indices are local. XML wrappers, irrelevant attributes,
cached values and repeated cells are omitted. Some result comments summarize a
policy and are not literal serialized output.

Red backgrounds and minus signs mark removals. Green backgrounds and plus signs
mark additions. Folded color bands connect each edit to a thin marker at its exact gap.
Removals use the right gutter and additions the left. Bands and markers share
their fill. A same-color center stroke enforces the 1.2-point minimum width.
A blue equals sign marks a retained structure. The code panels
are 252 points wide with 8.2-point type. They should be included at column width.

## Native evidence

`native-font/` in the evidence directory contains three constructed XLSX packages,
three installed-Excel captures and a receipt. Only the stated B2/C2 records differ.
The control demonstrates the font-witness mechanism separately from the named
corpus regression. To repeat it with installed Excel, run the following command.
It opens read-only fixtures in one owned Excel instance and replaces its PNGs
and receipt. Run the renderer and verification again afterward.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/xml-methods/capture-font-control.ps1
```

`enron-fill/` retains the historical full-rebuild captures and their provenance.
They were reused unchanged. Their original workbooks remain in the local corpus
case bundle identified by that provenance. No new whole-workbook validation is
claimed. The four guard-effect drawings are labeled schematic illustrations.

The source review used `1b9ee98` and a passing core suite with 298 tests and one
ignored manual timing test. `verification.json` records native image comparisons,
input hashes and PDF checks. The alias-name example explicitly retains the known
limitation that the implementation does not compare the formula bodies.

The visual references were Animated Vega-Lite Figure 2 and Bluefish Figure 2.
The figures here use newly drawn geometry and separately retained Excel images.

- https://vis.mit.edu/pubs/animated-vega-lite.pdf
- https://vis.csail.mit.edu/pubs/bluefish.pdf

## Vector Excel ranges

The figures now use `native-vector/` exports from Excel's print renderer with
printed gridlines, 100-percent scale and zero page margins. These differ from
the retained screen captures above. The source workbooks were opened read-only,
fully rebuilt and left unsaved. The Enron inputs were the retained rebuild/Save
controls. Their existing workbook and image evidence remains unchanged.

The cropped PDF ranges retain native text and paths. `native_insets.py` replaces
the declared preview-image draws with those PDF ranges before publishing the
figure. The native control figures and atlas contain no rasterized cell pictures. SVGs retain
600-dpi raster previews of those same ranges alongside editable vector code.

To repeat the exports, provide the retained Enron case directory and run.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/xml-methods/export-native-vectors.ps1
python scripts/xml-methods/prepare_native_vectors.py
python scripts/xml-methods/render.py
python scripts/xml-methods/verify.py
```

Use `-CorpusCaseDirectory` to override the case-bundle location. The cropper uses
the recorded range geometry. It retains the export receipt, verifies both
positive native pairs, and renders previews through Poppler at 600 dpi.
Tinos is a Times-compatible serif family from `https://github.com/googlefonts/tinos`.

## Coverage audit

`coverage.json` groups all 50 examples and accounts for all ten declared cleanup
causes, including disabled and inactive paths. Diagnostic fixtures are labeled
separately from the Default policy. The audit covers the principal rule families,
not every combination of Excel features. The core suite passed 298 tests with
one ignored manual timing test. Seven selected deferred-cleanup tests also passed.

## Wanderungen presentation case

`wanderungen-source.json` records the source SHA-256 and all 21 L:XFD zero runs.
To check the source again without opening Excel, run the following command.

```powershell
python scripts/xml-methods/inspect-wanderungen.py ORIGINAL.xlsx NEW_RECEIPT.json
```

The inspector verifies the exact column sequence and row identities. It makes no
new visual-validation or cleanup-performance claim. The 11 focused hidden-tail
regressions passed in the source revision recorded by the figure inventory.
