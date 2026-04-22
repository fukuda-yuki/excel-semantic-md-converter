# phase1-excel-com-rendering review note

## Review Scope

- Milestone: `phase1-excel-com-rendering`
- Scope: render planner、Excel COM wrapper、`render` CLI、render tests、関連 phase docs。
- Out of scope: `convert` 全体フロー、LLM integration、Markdown writer、`manifest.json` writer。
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-excel-com-rendering.md`
- `src/excel_semantic_md/cli/main.py`
- `src/excel_semantic_md/render/__init__.py`
- `src/excel_semantic_md/render/types.py`
- `src/excel_semantic_md/render/planner.py`
- `src/excel_semantic_md/render/excel_com_renderer.py`
- `tests/test_cli.py`
- `tests/test_render.py`

## Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Harvey` (`019db580-7227-78a0-9839-22af5767fb37`)
  - Result: 3 findings reported.
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Goodall` (`019db580-725d-7793-9c77-5275f8eda9c7`)
  - Result: 2 findings reported.

## Raw Findings Summary

- `Harvey`:
  - P1: `.xlsm` を macro-disabled で開いていない。
  - P2: image block の主成果物が元画像ではなく screenshot になっている。
  - P2: cell-based block の Range 画像を常に `markdown` 扱いしている。
- `Goodall`:
  - P1: `Workbooks.Open()` 失敗時に専用 Excel session が cleanup されない。
  - P2: `ExcelSession` の cleanup 経路を直接検証するテストが存在しない。
- Local pre-review validation: `python -m pytest tests/test_render.py tests/test_cli.py tests/test_visual_linker.py tests/test_ooxml_visual_reader.py tests/test_workbook_reader.py tests/test_block_detector.py tests/test_models.py` は 65 件成功。

## MainAgent Validity Judgment

- Initial implementation completed and the required subagent review has been collected.
- Accepted:
  - `.xlsm` macro-disabled mismatch
  - `Workbooks.Open()` failure cleanup leak
  - missing direct cleanup-path tests
  - image primary/original asset role inversion
  - cell-based range role mismatch
- Rejected:
  - none
- Accepted findings were fixed locally and revalidated with the full test suite.
- Milestone review requirements are satisfied for this milestone because the required subagent review was executed and the accepted findings were addressed.

## Response Plan

- Completed: ran the two required review subagents in parallel.
- Completed: fixed the accepted cleanup, macro-disabled, and planner-role issues.
- Completed: reran the full pytest suite and updated this note with the final post-fix state.

## Applied Fixes

- Added `render` planner models and Excel COM renderer for the live-confirmation-only `render` command.
- Implemented JSON output for `render` with `temp_dir`, artifact metadata, warnings, and failures.
- Ensured `.xlsm` rendering sessions force macro-disabled automation behavior before opening the workbook.
- Made `ExcelSession` cleanup idempotent and safe even when `Workbooks.Open()` fails after `DispatchEx()`.
- Restored application automation security during cleanup and surfaced workbook/application cleanup failures as warnings.
- Reworked planner roles so cell range captures are `render_artifact`, image original copies are the primary `markdown` artifact, and screenshot copies remain confirmation artifacts.
- Added direct `ExcelSession` tests for open-failure cleanup, macro-disabled behavior, and cleanup warning propagation.
- Re-ran `python -m pytest` and confirmed 67 tests passed after the fixes.

## Residual Risks

- `Range.CopyPicture` / `Shape.CopyPicture` / `Chart.Export` の実機安定性は未確認であり、live confirmation が必要。
- shape / image / chart matching は anchor と補助ヒントに依存するため、同一レイアウトの重複 object では曖昧性 failure になりうる。

## Pending Items

- Live confirmation on a real Excel environment remains pending and should capture the emitted JSON plus generated temp artifacts.
