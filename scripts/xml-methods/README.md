# XML method figures

[Review atlas](../../papers/sheetshrink/figures/xml-methods/xml-cleanup-atlas.pdf)
contains 28 column-width code comparisons, native controls, schematic failures
and a source index. Three selected figures are included in the manuscript.

## Rebuild

Install `reportlab`, `Pygments`, `Pillow` and `pypdf` in a local Python environment.
Then run from the repository root.

```powershell
python scripts/xml-methods/render.py
python scripts/xml-methods/verify.py
```

The LaTeX build consumes the generated PDFs and needs no Python or shell escape.
PDF text is selectable. SVG text and its font files are embedded for editing.
The bundled DejaVu and Bitstream Vera fonts retain their licenses in `fonts/`.

## Read the examples

The editable inventory is
`papers/sheetshrink/evidence/xml-methods/cases.json`. Each case pins a source
file, test or implementation anchor, line, commit and file hash to the reviewed
Rust revision. Style indices are local. XML wrappers, irrelevant attributes,
cached values and repeated cells are omitted. Some result comments summarize a
policy and are not literal serialized output.

Red backgrounds and minus signs mark removals. Green backgrounds and plus signs
mark additions. Folded color bands connect each edit to a thin marker at its exact gap.
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

The source review used `e6ae437` and a passing core suite with 298 tests and one
ignored manual timing test. `verification.json` records native image comparisons,
input hashes and PDF checks. The alias-name example explicitly retains the known
limitation that the implementation does not compare the formula bodies.

The visual references were Animated Vega-Lite Figure 2 and Bluefish Figure 2.
The figures here use newly drawn geometry and separately retained Excel images.

- https://vis.mit.edu/pubs/animated-vega-lite.pdf
- https://vis.csail.mit.edu/pubs/bluefish.pdf
