# Kostenstruktur row-height investigation

The first completed targeted NXPowerLite comparison contains 264 selected regions and 74 regions with observed differences. On worksheet `10.2 T1`, native rebuilt row 36 measures 15 points in the original and 12 points in the NXPowerLite output. The observed A1:AC41 images differ near the bottom text and in their retained height.

The source and NXPowerLite package both retain the literal row attribute `ht="14.4"` without `customHeight`. Source cells Q36:AB36 are empty style-228 cells with 11-point fonts. NXPowerLite removes those twelve cells, leaving the populated 9-point cells and P36. SheetShrink retains Q36:AB36. The package observations identify a plausible automatic-height witness mechanism. The native observation establishes the measured height difference for this attempt. Exclusive causal attribution requires the independent repeat and same-input control scheduled with the frozen A1:AC41 plan.

The selected range is a diagnostic follow-up selected after observing the finding. It is not an outcome-blind evaluation sample or a full-workbook reference.

## Independent repeat

All four second-attempt reference experiments completed in fresh native sessions. The original image SHA-256 was identical across all four experiments, `4045f4232b34414df39eb6160385eb48320c7d5c1ca3a01cbcb99665529d7452`.

| Comparison subject | Image dimensions | Changed pixels within the common image rectangle |
| --- | --- | --- |
| Original against original | 2485 x 800 / 2485 x 800 | 0 |
| NXPowerLite | 2485 x 800 / 2485 x 795 | 47,383 |
| Check Performance | 2485 x 800 / 2485 x 796 | 34,762 |
| SheetShrink | 2485 x 800 / 2485 x 799 | 0 |

The common-rectangle diagnostic compares exact RGB pixels at matching origins, without resampling or image registration. The SheetShrink difference is a one-pixel white bottom extent, so the explicit reference command's dimension-sensitive inequality does not imply changed visible content in this region. NXPowerLite and Check Performance also differ inside the common rectangle, beginning at pixel y=278. Their observations cannot be explained solely by the omitted white bottom extent. The original/native and optimized/native working copies are retained in each reference directory.

The first reference attempt stopped because the orchestration shell treated a normal stderr run-log announcement as a terminating NativeCommandError. The corrected wrapper retained stderr and continued to check process exit codes. These interrupted first attempts do not establish a rendering result.

`original.png`, `nxpowerlite.png` and `finding.json` retain the first finding. The row-attribute snapshots were extracted directly from ZIP/XML and do not open the workbook in a dataframe library. Original corpus workbooks remain unchanged.
