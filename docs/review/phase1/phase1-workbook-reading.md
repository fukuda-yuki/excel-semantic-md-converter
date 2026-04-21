# phase1-workbook-reading review note

## Review Scope

- Milestone: `phase1-workbook-reading`
- Scope: Workbook 読み取り層、`inspect` 接続、synthetic workbook fixture tests、関連 phase docs。
- Out of scope: block 検出、OOXML visual metadata、visual linking、Excel COM rendering、LLM integration、Markdown / manifest writer。
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `pyproject.toml`
- `src/excel_semantic_md/cli/main.py`
- `src/excel_semantic_md/excel/__init__.py`
- `src/excel_semantic_md/excel/workbook_reader.py`
- `tests/test_cli.py`
- `tests/test_workbook_reader.py`
- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`

## Subagents

- spec compliance and functional correctness reviewer: completed
- tests, edge cases, and regression risk reviewer: completed
- security reviewer: completed

## Raw Findings Summary

- Spec / functional reviewer findings:
  - Accepted: empty-string formula cache was incorrectly treated as cache-missing and could fail an otherwise valid sheet.
  - Accepted: workbook was opened with `read_only=False`, which did not satisfy the read-only contract.
  - Deferred as residual risk: automated `.xlsm` coverage still does not exercise a real VBA stream; macro-disabled behavior remains live confirmation territory.
- Tests / regression reviewer findings:
  - Accepted: malformed OOXML errors could leak traceback behavior and leave the file locked because metadata parsing happened after `load_workbook()`.
  - Rejected with spec update: `inspect` returning workbook-reading JSON at this milestone is intentional per the user-approved plan; `spec.md` was updated to record this interim command contract.
  - Rejected with spec/knowledge update: conservative number/date formatting is intentional for this milestone; `spec.md` and `knowledge.md` now record that full Excel display-value reproduction is not the target here.
  - Rejected as aligned with assumptions: filter-hidden handling is based on saved hidden-row metadata for this milestone.
  - Deferred as residual risk: synthetic formula-cache tests still do not cover every non-numeric cached type.
- Security reviewer findings:
  - Accepted: read-only contract issue matched the spec/functional finding above.
  - Accepted: malformed OOXML exception handling needed to be normalized into CLI errors.
  - No additional formula/comment/hyperlink leakage defects were identified after code review.

## MainAgent Validity Judgment

- Accepted and fixed: empty-string formula cache handling.
- Accepted and fixed: read-only open mode for workbook reading.
- Accepted and fixed: malformed OOXML cleanup / CLI error normalization.
- Rejected with explicit spec alignment: milestone-level `inspect` JSON shape and conservative display formatting are intentional for `phase1-workbook-reading`.
- Deferred: real-VBA `.xlsm` automation coverage remains a live-confirmation concern, not a blocker for this milestone.

## Response Plan

- Apply accepted fixes from subagent review.
- Update `spec.md` / `knowledge.md` where the user-approved milestone plan intentionally narrows or defers behavior.
- Re-run automated tests and update this note with final validation and residual risks.

## Applied Fixes

- Added `openpyxl>=3.1,<4`.
- Added Workbook reading result dataclasses and visible cell extraction.
- Connected `inspect --input` to Workbook reading JSON output.
- Added synthetic workbook tests for visible cells, hidden sheet/row/column/filter-hidden rows, `.xlsm` non-modification, formula cached values, formula cache missing failures, merged cells, text normalization, comments, hyperlinks, and inspect JSON output.
- Updated phase docs to record formula cache missing behavior and task completion.
- `python -m pytest` passed with 30 tests before subagent review.
- Added explicit date display normalization coverage in `tests/test_workbook_reader.py`.
- Switched workbook reading to `read_only=True` and moved hidden/merged/formula-cache inspection to OOXML metadata parsing.
- Fixed empty-string formula cache handling so a valid cached empty display value does not fail the sheet.
- Added malformed OOXML inspect coverage and verified the input file is not left locked after CLI failure.
- Updated `spec.md` to record the milestone-level `inspect` JSON contract and conservative display-formatting scope.
- Final `python -m pytest` passed with 32 tests after subagent-review fixes.

## Residual Risks

- Filter-hidden row handling relies on saved workbook hidden-row metadata, which matches the current milestone assumption but is not a full semantic interpretation of Excel filter state.
- Automated tests still do not exercise a workbook with a real VBA project stream; `.xlsm` macro-disabled behavior remains live confirmation.
- Synthetic formula-cache tests do not yet cover every non-numeric cached result type.

## Pending Items

- None for `phase1-workbook-reading`.
