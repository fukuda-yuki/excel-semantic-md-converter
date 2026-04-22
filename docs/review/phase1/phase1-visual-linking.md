# phase1-visual-linking review note

## Review Scope

- Milestone: `phase1-visual-linking`
- Scope: block 共通フィールド拡張、visual linking 層、`inspect` の post-linking block 出力、visual linking fixtures/tests、関連 phase docs。
- Out of scope: `manifest.json` writer、Markdown writer、Excel COM rendering、LLM integration。
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-visual-linking.md`
- `src/excel_semantic_md/models.py`
- `src/excel_semantic_md/excel/__init__.py`
- `src/excel_semantic_md/excel/visual_linker.py`
- `src/excel_semantic_md/cli/main.py`
- `tests/test_models.py`
- `tests/test_ooxml_visual_reader.py`
- `tests/test_visual_linker.py`
- `tests/test_workbook_reader.py`
- `tests/fixtures/visuals/table-image-visual.xlsx`
- `tests/fixtures/visuals/table-shape-visual.xlsx`

## Subagents

- spec compliance and functional correctness reviewer: not started
  - Reason: current session policy allows `spawn_agent` only when the user explicitly asks for sub-agents or delegation.
- tests, edge cases, and regression risk reviewer: not started
  - Reason: current session policy allows `spawn_agent` only when the user explicitly asks for sub-agents or delegation.

## Raw Findings Summary

- Self-review only:
  - Accepted during implementation: `inspect` block JSON now needs stable common fields for all block kinds, so `visual_id` / `related_block_id` were added at the base model rather than only visual-origin blocks.
  - Accepted during implementation: `absoluteAnchor` and other non-cell-addressable anchors still need standalone block preservation, so a synthetic trailing anchor plus warning was introduced and documented.
  - Accepted during implementation: final `related_block_id` must point at post-sort/post-renumber block IDs, so linking is resolved before sorting and rewritten after final ID assignment.
  - `python -m pytest tests/test_models.py tests/test_visual_linker.py tests/test_ooxml_visual_reader.py tests/test_workbook_reader.py tests/test_block_detector.py tests/test_cli.py` passed with 54 tests.
- Pending externalized review:
  - spec compliance / functional correctness review has not been run by a subagent.
  - tests / edge cases / regression review has not been run by a subagent.

## MainAgent Validity Judgment

- `inspect` now reflects the intended processing order: workbook reading -> block detection -> visual metadata -> visual linking.
- linked/unlinked shape/image/chart visuals are preserved as blocks with explicit `visual_id` / `related_block_id`, while unsupported `unknown` visuals remain in `visuals`.
- The current implementation intentionally prepares manifest-compatible block schema without claiming that `manifest.json` writing itself is complete.
- Milestone review requirements are not fully satisfied yet because the required subagent review could not be started under the current session policy.

## Response Plan

- Keep the implementation, fixtures, and tests in place.
- If the user explicitly authorizes subagent/delegated review, run the two required reviewers and update this note before declaring the milestone complete.
- Until then, treat this note as a partial review record and keep the review gap explicit.

## Applied Fixes

- Added `visual_id` / `related_block_id` to the shared block schema and preserved round-trip compatibility for existing block types.
- Added `link_visuals()` and exported it from `excel.__init__`.
- Implemented anchor rect normalization, heading-scope linking, nearest-block fallback, synthetic standalone anchors, final block re-sorting, and block ID re-assignment.
- Updated `inspect` to emit post-linking `blocks` while keeping raw `visuals` output unchanged.
- Added synthetic `table + image` and `table + text shape` OOXML fixtures plus visual linking tests.
- Updated phase docs to reflect the new public contract and the manifest-writer boundary.

## Residual Risks

- `absoluteAnchor` fallback uses a synthetic trailing anchor because the OOXML payload does not provide a cell range; this preserves the block but does not represent true screen position.
- `related_block_id` is one-to-one from each visual-origin block to one target block; reverse indices or per-section link tables are intentionally deferred.
- `manifest.json` writer remains unimplemented, so persistence of the new link fields is only schema-ready today.
- The review workflow still lacks the required subagent pass, so regression / spec gaps may remain undiscovered.

## Pending Items

- Obtain explicit authorization for subagent review, then rerun and update this note.
- Implement `manifest.json` writer so `visual_id` / `related_block_id` are persisted in output artifacts, not only `inspect`.
