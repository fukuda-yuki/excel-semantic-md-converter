# phase1-llm-integration review note

## Review Scope

- Milestone: `phase1-llm-integration`
- Scope: `llm/` 層のモデル・prompt/input builder・attachment builder・response parser・Copilot SDK adapter、関連 tests、phase docs 更新
- Out of scope: `convert` 全体フロー、`result.md` writer、`manifest.json` writer、orchestrator、実 SDK / vision attachment の live confirmation
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-llm-integration.md`
- `pyproject.toml`
- `src/excel_semantic_md/llm/__init__.py`
- `src/excel_semantic_md/llm/adapter.py`
- `src/excel_semantic_md/llm/builders.py`
- `src/excel_semantic_md/llm/models.py`
- `src/excel_semantic_md/llm/parser.py`
- `src/excel_semantic_md/llm/prompt.py`
- `tests/test_llm.py`

## Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Hilbert` (`019db593-34aa-7a93-b21d-3e7308ee969d`)
  - Result: 2 findings reported.
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Hooke` (`019db593-3556-7a32-bfb1-1f76912c3c71`)
  - Result: 3 findings reported.
- security reviewer: completed
  - Agent: `Gauss` (`019db593-35a6-7512-bc28-d891c477717d`)
  - Result: 4 findings reported.

## Raw Findings Summary

- `Hilbert`:
  - High: `client.stop()` failure can leak a raw exception and violate the sheet-level failed-and-continue contract.
  - Medium: attachment ordering did not use related-block proximity and therefore did not match the spec's "importance + proximity" requirement.
- `Hooke`:
  - P1: `client.stop()` failure path lacked recovery and regression coverage.
  - P2: task validation no longer showed the remaining live confirmation items for Copilot SDK local CLI behavior and vision attachment behavior.
  - P3: `Copilot SDK adapter は mock でテストできるようにする。` was duplicated as both unchecked and checked.
- `Gauss`:
  - High: the adapter auto-approved permission requests, which weakens the prompt-injection boundary.
  - Medium: attachment selection and prompt payload exposed more path/context than necessary.
  - Medium: making `github-copilot-sdk` a required dependency may broaden install-time attack surface.
  - Medium: `client.stop()` failure leaked outside the `LlmRunResult` contract.
- Local validation before and after fixes:
  - `python -m pytest` passed with 78 tests after the fixes.

## MainAgent Validity Judgment

- Initial implementation completed and the required subagent review was collected.
- Accepted:
  - `client.stop()` exception propagation bug
  - missing cleanup regression test
  - attachment ordering missing related-block proximity
  - auto-approved permission request boundary
  - prompt payload exposing absolute local paths
  - task validation missing live confirmation items
  - duplicate task checklist line
- Rejected:
  - making `github-copilot-sdk` optional in packaging for this milestone
    - Reason: this milestone adds a runtime LLM adapter rather than a separate opt-in install flow, and the source-of-truth docs do not define extras-based installation. The dependency remains isolated by import boundary, and install-flow redesign is out of scope here.
  - treating "no explicit `--max-images-per-sheet` default" itself as a spec bug
    - Reason: `spec.md` defines selection behavior once a cap applies, but does not define a required default cap. The implementation was still tightened by proximity-based ordering and by removing absolute paths from prompt JSON.
- Accepted findings were fixed locally and revalidated with the full test suite.
- Milestone review requirements are satisfied for this milestone because the required subagent review was executed and the accepted findings were addressed.

## Response Plan

- Completed: replace permissive session permission handling with the SDK's default deny-by-default behavior.
- Completed: ensure SDK cleanup failure is converted into a sheet-level failed result instead of a raw exception.
- Completed: rank attachment candidates by importance and related-block proximity, then add regression coverage.
- Completed: redact absolute local paths from LLM input JSON while preserving absolute file paths for SDK attachments.
- Completed: restore accurate `task.md` state for mock-test coverage and pending live confirmation work.

## Applied Fixes

- Added `LlmRunOptions`, `LlmAttachment`, `LlmInput`, `LlmResponse`, and `LlmRunResult`.
- Added prompt construction that keeps instructions in Python and treats workbook text as data.
- Added sheet-scoped input JSON building and attachment ranking from `RenderSheetResult`.
- Added JSON parsing/validation for plain JSON and fenced JSON responses, with a single retry on validation failure.
- Added a GitHub Copilot SDK adapter with runtime-only SDK imports so non-LLM imports do not require `copilot`.
- Removed unconditional permission auto-approval from session creation.
- Wrapped SDK shutdown so cleanup failures return `LlmRunResult(status="failed")` instead of leaking exceptions.
- Changed prompt-facing asset metadata to use file names rather than absolute local paths.
- Added LLM tests for prompt contract, attachment selection, retry behavior, cleanup failure handling, and import-boundary isolation.
- Updated `task.md` and `knowledge.md` to reflect the implemented milestone and the remaining live confirmation work.

## Residual Risks

- `--vision-model` pass-through is implemented, but actual SDK compatibility remains unverified and still requires live confirmation.
- Real Copilot SDK / local CLI behavior and image attachment behavior remain outside automated test coverage.
- `convert` still does not orchestrate render -> LLM -> output generation; this milestone only prepares the `llm/` boundary.

## Pending Items

- Run live confirmation for Copilot SDK local CLI behavior.
- Run live confirmation for vision attachment behavior.
- Wire the new `llm/` boundary into a later `convert` / output-generation milestone.
