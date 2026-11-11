# Current capture and report defaults

Checked on 11 November 2026 after fetching and fast-forwarding
`codex/datagov-validation-watch` to `7e3dd27`.

| Step | Current behavior | Source in `excel-shrink-cli/src/` |
| --- | --- | --- |
| Transport | Control Room defaults to `live-images`. The resident path selects pipelined image capture unless hash-first or serialized execution is explicitly selected | `capture_host/control_room.rs`, `regression/office_capture/resident_capture_live.rs` |
| Native evidence | Image capture returns finalized PNGs and validated image receipts. Hashes also serve identity and integrity checks | `regression/office_capture/resident_capture_pipeline.rs`, `capture_host/resident.rs` |
| Retention | `CaptureOptions::default()` and new study configurations set `retain_equal_images` to `true` | `validation.rs`, `validation_study/configuration.rs` |
| Discard gate | Discard of verified-equal images requires a successful comparison and an explicit false retention setting | `validation/run.rs`, around line 595 |
| Packaging | Copies report-addressed images, replaces references and publishes an integrity manifest for retained bytes | `validation_report/evidence.rs::ImageStore::retain_report_images` |
| Equal-image encoding | Best PNG compression and adaptive filters. A smaller encoding is accepted after exact decoded RGBA equality checking | `validation_report/evidence.rs::ImageStore::retain_image`, around line 172 |
| Protected images | Differing or unverified references protect their source encoding. Explicit `preserve_image_encoding` also preserves source encoding | `validation_report/evidence.rs::classify_image_references`, `retain_report_images` |

Report evidence is stored under `evidence/<sha256>.png`. Matching retained bytes
can share a content-addressed file. The manifest records the final file identity.
Recompression preserves pixel dimensions and decoded RGBA pixels. Encoded PNG
bytes and their checksum can change.

Existing checks include
`equal_png_recompression_preserves_pixels_and_protects_differences`,
`an_image_shared_with_a_difference_is_protected` and
`capture_transport_defaults_to_live_images_and_rejects_unknown_modes`.
These tests were inspected as source for this audit.

## Figure correction

The first version of Figure 12 selected the supported hash-first alternative
as its main example. The default image transport was already present in the
previously reviewed snapshot. Panel A now follows image upload, image receipt,
result commit and report retention, including verified lossless recompression.

## Internal compatibility note

The older archive-study retention entry point has its own preview policy.
`regression/archive/study/prepared/targeted.rs` defaults `equal_preview_width`
to 512 before calling `retention::apply_with_policy`. That path can downscale
confirmed-equal regions after full-resolution comparison, while protecting
differences and their context. A width of zero keeps full resolution.
This distinction is recorded here for source review. The manuscript figure
describes the current default publication path.
