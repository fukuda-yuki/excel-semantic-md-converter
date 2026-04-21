# phase1-block-detection review note

## Review Scope

- Milestone: `phase1-block-detection`
- Scope: cell / merged-range ベースの block 検出、`inspect` への block JSON 追加、関連 phase docs、synthetic workbook tests。
- Out of scope: OOXML visual metadata、shape / image / chart linking、Excel COM rendering、LLM integration、Markdown / manifest writer。
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-block-detection.md`
- `src/excel_semantic_md/cli/main.py`
- `src/excel_semantic_md/excel/__init__.py`
- `src/excel_semantic_md/excel/block_detector.py`
- `tests/test_block_detector.py`
- `tests/test_workbook_reader.py`

## Subagents

- spec compliance and functional correctness reviewer: not started
  - Reason: current session policy allows `spawn_agent` only when the user explicitly asks for sub-agents or delegation.
- tests, edge cases, and regression risk reviewer: not started
  - Reason: current session policy allows `spawn_agent` only when the user explicitly asks for sub-agents or delegation.

## Raw Findings Summary

- Self-review only:
  - Accepted and fixed during implementation: merged caption paragraph was incorrectly receiving both `table_caption_candidate` and `mixed_sparse_region`.
  - No additional automated failures remained after `python -m pytest`.
- Pending externalized review:
  - spec compliance / functional correctness review has not been run by a subagent.
  - tests / edge cases / regression review has not been run by a subagent.

## MainAgent Validity Judgment

- `spec.md` conflict on caption handling was resolved by updating the spec to the user-confirmed `paragraph + warning` rule before implementation.
- `inspect` now matches the milestone contract: workbook reading JSON is preserved and `blocks` are appended per sheet.
- The current heuristic set is intentionally conservative and only covers cell-based `heading` / `paragraph` / `table`.
- Milestone review requirements are not fully satisfied yet because the required subagent review could not be started under the current session policy.

## Response Plan

- Keep the implementation and tests in place.
- If the user explicitly authorizes subagent/delegated review, run the two required reviewers and update this note with their findings before declaring the milestone complete.
- If the user does not want subagents, treat this note as a partial review record and keep the remaining gap explicit.

## Applied Fixes

- Added `detect_blocks(read_result) -> WorkbookModel` and exported it from `excel.__init__`.
- Implemented used-range estimation, recursive empty-row / empty-column splitting, stable block ordering, conservative `table` / `heading` / `paragraph` detection, and warning emission.
- Extended `inspect` to append `blocks` while keeping existing workbook reading JSON fields.
- Updated phase docs for the user-confirmed caption handling and milestone behavior.
- Added synthetic tests for table detection, header column detection, heading extraction, merged caption handling, sparse paragraph fallback, hidden row / column exclusion, stable block IDs, and `inspect` JSON output.
- `python -m pytest` passed with 39 tests.

## Residual Risks

- Header detection remains heuristic and text-only; some string-only tables may still be classified conservatively as `paragraph`.
- The review workflow still lacks the required subagent pass, so regression / spec gaps may remain undiscovered.
- Visual metadata and block-to-asset linking are intentionally out of scope for this milestone.

## Pending Items

- Obtain explicit authorization for subagent review, then rerun and update this note.
