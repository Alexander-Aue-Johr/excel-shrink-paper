# Six-file Destatis comparator benchmark

The six frozen candidates in `../results/provisional-panel.json` were processed on 2026-11-09. Each comparator received separate SHA-256-verified original copies. SheetShrink outputs are from the pinned a7eef1a build. Inquire is deferred to another machine.

## Comparator settings

- NXPowerLite Desktop 11.0.1 x64 trial. Balanced compression (`Ausgewogene Komprimierung`), optimized copies in the original working directory. Excel options observed before processing: JPEG quality 7, resize to 1280 x 720, allow JPEG conversion, remove cropped image areas, flatten embedded objects, reduce PNG color depth, and remove excess cell formatting enabled. The built-in profile was retained. Its outputs were copied without modification into `nxpowerlite/outputs`, with the original filenames restored for pairing. These settings permit image changes and must accompany any interpretation of visual differences.
- Excel Check Performance, Excel 16.0.20527.20150, de-DE, one owned private-desktop worker. The pinned production `regression performance-prepare` command recorded positive native completion for all six inputs. Manual calculation, macros disabled by the worker. All source, ordinary-save, before-optimization and after-optimization artifacts and receipts are retained. Attempt 1 failed during helper deployment before processing because the capture assets were not explicitly bound. Attempt 2 explicitly bound the pinned source capture directory and completed all six in 278.502 seconds. This elapsed time includes orchestration and is not a controlled cross-tool performance measurement.
- XLStylesTool UWP 3.0.1.0, all six scanned and processed individually. `Force all Styles to Normal`, `Unhide All`, and external named-range `Remove All` were disabled for every file. The first two were initially enabled and were explicitly cleared before the first processing action. The application retained those choices for subsequent files. In-place outputs were copied to `xlstyles/outputs`. Invalid named ranges and unused styles were removed by the application.

## XLStylesTool observations

| Workbook | Styles before / after | Defined names before / after | Invalid names before / after |
| --- | --- | --- | --- |
| Stromsteuerstatistik | 15 / 11 | 98 / 0 | 98 / 0 |
| Wanderungen | 30 / 25 | 176 / 135 | 41 / 0 |
| Kostenstruktur Aerzte | 28 / 28 | 302 / 300 | 2 / 0 |
| Zeitverwendung | 48 / 12 | 2293 / 162 | 2131 / 0 |
| Verkehrsunfaelle | 9 / 9 | 143 / 108 | 35 / 0 |
| Kernhaushalt Gemeinden | 15 / 13 | 643 / 113 | 530 / 0 |

These counts were transcribed from the native scan and update screens. Package inspection independently verifies retained outputs.

## Visual evaluation

Native DeltaSieve comparisons use the same pinned a7eef1a source, one workbook pair at a time, 300-second operation and idle-progress limits, ordinary-save controls and paired rebuilt states. Equal images are retained. Exact launch configurations, raw native evidence and failures remain in each tool directory. A failed or incomplete capture is unresolved. Pixel inequality requires attribution before it is classified as an optimization regression.

No original Wanderungen workbook was opened with Python or Pandas. ZIP/XML inspection is permitted and separate from native Excel rendering.
