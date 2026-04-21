# phase1-cli review note

## Review Scope

- Milestone: `phase1-cli`
- Scope: CLI 基盤のみ。Workbook 読み取り、block 検出、Excel COM rendering、LLM、manifest / Markdown 生成は対象外。
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `src/excel_semantic_md/cli/main.py`
- `tests/test_cli.py`
- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`

## Subagents

- spec compliance and functional correctness reviewer: completed
- tests, edge cases, and regression risk reviewer: completed

## Raw Findings Summary

- Spec compliance reviewer:
  - Review note still recorded required reviews as pending. Blocks workflow completion until updated.
  - `setup --out` cleanup behavior was not asserted by tests.
  - CLI runtime implementation aligned with the CLI foundation scope.
- Tests / edge cases reviewer:
  - `setup --out` parent traversal could loop forever when no existing parent is reachable.
  - `setup --out` side effects were not asserted.
  - Non-blocking test gaps remain for some shared validator cases such as existing file at `--out`, input directory, uppercase `.XLSX`, and `.xls` rejection.
  - Skill launcher diagnostics depend on running from the repository root; not blocking for this milestone.

## MainAgent Validity Judgment

- Valid and accepted: `setup --out` parent traversal needed a root guard. A user-supplied path must not hang setup diagnostics.
- Valid and accepted: `setup --out` cleanup behavior should be asserted because `docs/phase1/spec.md` now explicitly documents it.
- Valid and accepted: review note had to be updated before completion.
- Deferred: broader shared-validator tests are useful but not required to complete `phase1-cli`; existing tests cover the committed behavior and common errors.
- Deferred: skill launcher lookup from non-repository working directories is a diagnostic-quality improvement, not a spec blocker for this milestone.

## Response Plan

- Apply the accepted `setup --out` root guard fix.
- Add cleanup assertions for created output directories and existing output directories.
- Record review results and final verification.

## Applied Fixes

- Initial implementation completed before review.
- Added a parent traversal guard in `setup --out` output directory probing.
- Added tests that assert `setup --out` removes probe files, removes directories created only for probing when possible, and preserves existing output directories.
- Added a regression test for missing-parent traversal returning a diagnostic instead of looping.
- `python -m pytest` passed with 20 tests before review note creation.
- `python -m pytest` passed with 22 tests after review fixes.
- Final `python -m pytest` passed with 22 tests.

## Residual Risks

- `setup` can only report Excel COM and Copilot availability heuristically; final confirmation remains live confirmation.
- `convert`, `inspect`, and `render` validate CLI input but intentionally stop before downstream processing.

## Pending Items

- None for `phase1-cli`.
