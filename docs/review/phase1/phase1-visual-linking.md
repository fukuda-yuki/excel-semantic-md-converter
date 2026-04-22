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

- spec compliance and functional correctness reviewer: completed
  - Agent: `Maxwell` (`019db56b-d710-7a00-92c9-c2bad30151fe`)
  - Result: 2 findings reported.
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Avicenna` (`019db56b-d755-79f3-adf1-39f9aa05848e`)
  - Result: 3 findings reported.

## Raw Findings Summary

- `Maxwell`:
  - High: heading scope 判定が「その heading block から次の heading block 手前まで」を満たさず、section 後半の visual を誤って scope 外扱いにする。
  - Medium: synthetic anchor 採番が cell block の最終行しか見ないため、より下に実 anchor visual がある sheet で synthetic block が末尾側に置かれない。
- `Avicenna`:
  - P2: synthetic anchor ordering が lower real visual を考慮せず不安定。
  - P2: `link_visuals()` が入力 `block_model` の block object を破壊的に再採番している。
  - P3: visual block の `Rect.a1` が `from` / `to` から再計算されず、stale/noncanonical な `a1` をそのまま通す。
- MainAgent raw consolidation:
  - synthetic anchor ordering の指摘は 2 本の subagent で重複している。
  - `python -m pytest tests/test_models.py tests/test_visual_linker.py tests/test_ooxml_visual_reader.py tests/test_workbook_reader.py tests/test_block_detector.py tests/test_cli.py` は、review 開始前時点で 54 件成功だった。

## MainAgent Validity Judgment

- `inspect` now reflects the intended processing order: workbook reading -> block detection -> visual metadata -> visual linking.
- linked/unlinked shape/image/chart visuals are preserved as blocks with explicit `visual_id` / `related_block_id`, while unsupported `unknown` visuals remain in `visuals`.
- The current implementation intentionally prepares manifest-compatible block schema without claiming that `manifest.json` writing itself is complete.
- Accepted:
  - heading scope bug
  - synthetic anchor ordering bug
  - input `block_model` mutation bug
  - `Rect.a1` normalization gap
- Rejected:
  - none
- Accepted findings were fixed locally and revalidated with the full test suite.
- Milestone review requirements are satisfied for this milestone because the required subagent review was executed and the accepted findings were addressed.

## Response Plan

- Completed: fixed heading scope so section coverage is row-range based rather than insertion-index based.
- Completed: fixed synthetic anchor allocation so fallback anchors are placed after both cell blocks and addressable visual anchors on the sheet.
- Completed: made `link_visuals()` non-mutating with respect to its input `WorkbookModel`.
- Completed: normalized linked visual block `Rect.a1` from numeric bounds and extended tests for all accepted findings.

## Applied Fixes

- Added `visual_id` / `related_block_id` to the shared block schema and preserved round-trip compatibility for existing block types.
- Added `link_visuals()` and exported it from `excel.__init__`.
- Implemented anchor rect normalization, heading-scope linking, nearest-block fallback, synthetic standalone anchors, final block re-sorting, and block ID re-assignment.
- Updated `inspect` to emit post-linking `blocks` while keeping raw `visuals` output unchanged.
- Added synthetic `table + image` and `table + text shape` OOXML fixtures plus visual linking tests.
- Updated phase docs to reflect the new public contract and the manifest-writer boundary.
- Reworked heading scope matching to use row-range coverage up to the next heading instead of insertion-index boundaries.
- Reworked synthetic anchor allocation to place fallback rows after all addressable visual anchors on the sheet.
- Cloned input blocks inside `link_visuals()` so the source `WorkbookModel` remains unchanged.
- Normalized linked visual block `Rect.a1` from computed numeric bounds and added regression tests for late-section heading scope, mixed synthetic/real ordering, non-mutating behavior, and `a1` normalization.
- Re-ran `python -m pytest` and confirmed 58 tests passed after the fixes.

## Residual Risks

- `absoluteAnchor` fallback uses a synthetic trailing anchor because the OOXML payload does not provide a cell range; this preserves the block but does not represent true screen position.
- `related_block_id` is one-to-one from each visual-origin block to one target block; reverse indices or per-section link tables are intentionally deferred.
- `manifest.json` writer remains unimplemented, so persistence of the new link fields is only schema-ready today.

## Pending Items

- Implement `manifest.json` writer so `visual_id` / `related_block_id` are persisted in output artifacts, not only `inspect`.
