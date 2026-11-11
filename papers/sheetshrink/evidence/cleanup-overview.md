# Figure 2 cleanup inventory

Reviewed against Rust revision `1b9ee98`, the behavior specification and the
source/test inventory on 11 November 2026. Figure 2 groups active transformations
into nine schematic examples. Each example has its own eligibility conditions.

| Figure example | Implemented transformation and source | Representative test anchors |
| --- | --- | --- |
| Empty cells and redundant formatting | `sheet/cells.rs` removes unnecessary visible/hidden empty records, including redundant inherited fill/border and irrelevant protection-only records. `sheet/output.rs::remove_protection_only_cells_outside_bounds` applies exterior protection policy. Explicit no-fill, borderless and unstyled masks remain when needed. | `removes_empty_cell_when_default_fill_override_has_nothing_to_clear`; `keeps_empty_cell_when_default_fill_override_clears_visible_row_fill`; `keeps_empty_cell_when_default_border_override_clears_visible_row_border`; `protected_sheet_keeps_empty_apply_protection_cell_inside_visible_bounds` |
| Repeated font and layout witnesses | `sheet/rows.rs::remove_unneeded_style_representatives` retains required representatives and removes redundant occurrences. Font-record identity, local inheritance and later anchors matter. | `keeps_one_empty_no_empty_cell_visible_style_cell_per_style_when_row_has_no_custom_height`; `merged_font_representatives_are_local_to_each_merge`; `font_override_is_preserved_when_content_follows` |
| Styles promoted into a row | `sheet/rows.rs` reuses an existing row-compatible style and removes repeated empty cell records. Dense, bounded-gap, sparse and uniform-fill routes keep their individual guards and necessary exceptions. | `dense_full_width_row_compacts_identical_empty_styles_into_row_header`; `dense_row_promotion_requires_an_explicit_fixed_height`; `full_width_fill_promotion_requires_the_same_explicit_fill_on_content_cells` |
| Formatting-only tails | `sheet/output.rs` prunes cell/range and row tails after retained content/layout anchors, including redundant uniform fills. Hidden trailing rows can be collapsed while their final boundary remains. Deferred removals trigger row cleanup again. | `enron_nov_styleless_representatives_do_not_anchor_a_content_free_sheet`; `enron_august_no_fill_overrides_do_not_anchor_an_empty_sheet`; `column_exceptions_and_uncovered_columns_preserve_the_white_tail`; `trailing_hidden_tail_keeps_only_last_hidden_row`; `phase3_trims_trailing_row_after_deferred_candidate_cleanup` |
| Column-definition tails | `sheet/cols.rs::merge_extraneous_columns_with_styles` coalesces compatible default/equal-fill exterior definitions and retains needed font identities. Width/border, gaps, overlaps and hidden-column guards apply. | `collapses_visible_equal_fill_column_tail_with_no_borders`; `equal_fill_does_not_paint_an_uncovered_gap`; `keeps_one_default_font_column_before_a_different_font_tail`; `visible_row_borders_preserve_empty_column_widths` |
| Hidden constant tails | `sheet/hidden_content.rs` and workbook-deferred cleanup admit dense horizontal suffixes of at least 4,096 equal signed integers ending at XFD. References, layout and protection can prevent removal. Required cell formatting can survive payload removal. | `hidden_tail_still_removes_wanderungen_full_width_zero_tail` in `pipeline_hidden_tail_tests.rs`; `pipeline_retains_hidden_chart_and_drawing_inputs`; `phase3_row_cleanup_preserves_stripped_whitespace_cell_with_visible_fill` exercises retained formatting after a separate deferred payload removal |
| Workbook metadata | `workbook.rs` removes recognized built-in aliases when their `_xlnm.` counterpart has the same scope. `pipeline.rs::clean_calc_chain` removes entries for absent cells in processed sheets with resolved identities. The alias predicate currently checks names/scope and lacks formula-body equivalence checking. | `aliases_on_different_sheets_are_not_duplicates_when_scope_precedes_name`; `a_self_closing_name_does_not_consume_a_following_element`; `calc_chain_removes_missing_cells_for_processed_sheet_only` |
| Compact syntax and row spans | `sheet/cells.rs::keep_empty_cell_as_self_closing`, `sheet/rows.rs::render_empty_row` and `rewrite_row_spans_upper_bound` compact eligible empty representations and update existing span upper bounds after removal. | `self_closing_row_stays_borrowed_and_still_records_visible_style`; `spans_use_only_current_row_cells_after_recleaning`; `spans_exclude_pruned_representatives_but_include_retained_visible_cells` |
| Optional whitespace | `sheet/cells.rs` and deferred pipeline cleanup remove eligible unreferenced whitespace payloads under explicit switches, retaining needed formatting. Default disables this policy. Maximum and diagnostic profiles can alter the representative policy as well. | `phase3_strips_unprotected_x4_whitespace_candidate_cell_content`; `phase3_row_cleanup_preserves_stripped_whitespace_cell_with_visible_fill`; `hidden_column_content_preserves_automatic_row_height` |

Source paths above are relative to `excel-shrink-core/src`, except CLI pipeline
paths. Full test paths are indexed in `parser-regressions/rust-tests.json`.
The 50-case XML atlas supplies detailed positive/negative examples for these
mechanisms. Figure 2 summarizes families, with inherited-format masks, font
witnesses, merged spans, borders, drawings and protected dependencies represented
by their guards.

`CleanupCause` in `models.rs` is a telemetry taxonomy. Some changes share a cause,
and the existence of a cause does not establish an active deletion path.
`remove_irrelevant_attributes_for_empty_rows` returns its input unchanged.
The early self-closing-cell preprocessor is disabled by default and feeds the
same empty-cell family when enabled. The old hidden-row removal path is disabled.
Styles, shared strings and font tables retain their indices. Formula evaluation,
image/object removal and global style-table compaction are outside this inventory.