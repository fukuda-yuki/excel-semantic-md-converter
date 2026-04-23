# phase1-requirements-implementation-review review note

## Review Scope

- Milestone: `phase1-requirements-implementation-review`
- Scope: `docs/phase1/spec.md` を Source of Truth とした Phase 1 仕様・実装・テスト・skill の横断レビュー
- Reviewed areas:
  - `docs/phase1/spec.md`
  - `docs/phase1/task.md`
  - `docs/phase1/knowledge.md`
  - `src/excel_semantic_md/**`
  - `tests/**`
  - `skills/excel-semantic-markdown/**`
- Out of scope: コード修正、仕様修正、タスクチェックリスト更新、live confirmation の実施
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `src/excel_semantic_md/render/excel_com_renderer.py`
- `src/excel_semantic_md/output/writers.py`
- `src/excel_semantic_md/app/convert_pipeline.py`
- `src/excel_semantic_md/llm/adapter.py`
- `src/excel_semantic_md/llm/builders.py`
- `src/excel_semantic_md/llm/models.py`
- `src/excel_semantic_md/llm/__init__.py`
- `tests/test_render.py`
- `tests/test_ooxml_visual_reader.py`
- `tests/test_output.py`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-requirements-implementation-review.md`

## Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Pasteur` (`019dba5f-23a9-7ad2-b13c-ee5d4d84fccf`)
  - Launch status: success after initial spawn option retry
  - Result: no certain spec mismatch reported
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Godel` (`019dba5f-23f0-73e2-8858-c95de00ec507`)
  - Launch status: success after initial spawn option retry
  - Result: 2 findings and 2 residual risks reported
- architecture and over-implementation reviewer: completed
  - Agent: `Kant` (`019dba5f-241c-7e91-9ed1-b9a6d35c6f56`)
  - Launch status: success after initial spawn option retry
  - Result: 2 findings reported
- reliability, cleanup, and security reviewer: completed
  - Agent: `Peirce` (`019dba5f-246c-7382-bdb7-38c0a0a2293c`)
  - Launch status: success after initial spawn option retry
  - Result: 3 findings and 3 residual risks reported

## Raw Findings Summary

- `Pasteur`
  - No certain mismatch between `docs/phase1/spec.md` and implementation was found in the reviewed scope.
  - Residual risk: Excel COM and Copilot SDK live confirmation was not executed.
- `Godel`
  - P2: `render` CLI success path is not tested against the stdout JSON contract, non-generation of `result.md` / `manifest.json`, and LLM non-invocation.
  - P2: drawing part parse failure is not covered as a warning-and-continue path.
  - Residual risk: Copilot SDK local CLI behavior and vision attachment behavior remain live-confirmation pending.
  - Residual risk: `setup` and real Excel COM / Copilot environment behavior remain live-confirmation dependent.
- `Kant`
  - P2: LLM input / attachment construction is duplicated between `convert_pipeline` debug preparation and `GitHubCopilotSdkAdapter`, so debug JSON can drift from the actual request.
  - P3: workbook read -> block detection -> visual metadata -> linking flow is duplicated across `convert`, `inspect`, and `render` paths.
- `Peirce`
  - P1: Excel COM rendering suppresses failure to set `AutomationSecurity = 3`, so `.xlsm` macro-disabled handling can fail open.
  - P2: exception strings stored in failure details can include absolute paths, while sanitization only redacts path-like values under selected keys.
  - P2: managed output replacement deletes old outputs before moving staging outputs, so a move failure can leave partial or missing outputs.
  - Residual risk: `.xlsm` macro-disabled behavior still needs a real VBA-containing workbook live confirmation.
  - Residual risk: Copilot SDK attachment path behavior remains external-service dependent.
  - Residual risk: Excel COM cleanup needs confirmation on a machine with existing user Excel processes.

## MainAgent Validity Judgment

Accepted findings:

1. P1: `.xlsm` macro-disabled contract can fail open when `AutomationSecurity` cannot be set.
   - Evidence: `src/excel_semantic_md/render/excel_com_renderer.py:53-57` catches and ignores exceptions while setting `self.app.AutomationSecurity = 3`, then opens the workbook at lines 58-65.
   - Spec reference: `docs/phase1/spec.md` sections 3.2, 7.1, 11.3.
   - Judgment: Valid. Phase 1 explicitly says `.xlsm` is read-only and macro-disabled. If macro disabling cannot be enforced, continuing is not a safe default.
   - Direction for next phase: fail closed for macro-capable rendering when automation security cannot be forced, or otherwise prove a safe macro-disabled path.

2. P2: `manifest.json` / debug payload can leak absolute paths embedded in exception strings.
   - Evidence: `src/excel_semantic_md/output/writers.py:407-417` redacts only values keyed as `path`, `workbook`, or `temp_dir`, or strings containing known temp prefixes. Errors are stored as strings in `src/excel_semantic_md/render/excel_com_renderer.py:176-182` and `src/excel_semantic_md/llm/adapter.py:41-49`, `83-90`, `178-200`.
   - Spec reference: `docs/phase1/spec.md` sections 9.4, 9.5, 11.3.
   - Judgment: Valid. Debug output is opt-in and may contain workbook data, but default `manifest.json` should not accidentally expose full local paths through error text.
   - Direction for next phase: scrub path-like substrings in all failure detail strings or replace raw exception text with bounded error summaries.

3. P2: `render` success-path CLI contract is under-tested.
   - Evidence: `tests/test_render.py` includes failure-oriented CLI coverage, but no success-path assertion that stdout contains the full artifact JSON contract, that `result.md` / `manifest.json` are not generated, and that LLM is not called.
   - Spec reference: `docs/phase1/spec.md` section 2.4 and section 11.3.
   - Judgment: Valid as a test gap. The implementation appears consistent, but the externally supported live-confirmation command has an insufficient regression guard.
   - Direction for next phase: add a monkeypatched success-path CLI test for `render`.

4. P2: drawing XML parse failure warning-and-continue behavior is under-tested.
   - Evidence: existing OOXML tests cover missing drawing targets and unknown shape handling, but no fixture verifies a malformed drawing part is retained as a sheet/visual warning rather than promoted to workbook failure.
   - Spec reference: `docs/phase1/spec.md` sections 6.2 and 12.1.
   - Judgment: Valid as a test gap. The spec explicitly distinguishes drawing-part parse failures from workbook/worksheet primary XML corruption.
   - Direction for next phase: add a corrupted drawing XML fixture and assert `read_visual_metadata` / `inspect` continue with warnings.

5. P2: LLM request/debug construction is duplicated and can drift.
   - Evidence: `src/excel_semantic_md/app/convert_pipeline.py:186-203` builds attachments and `llm_input_payload` for debug/output, while `src/excel_semantic_md/llm/adapter.py:19-33` rebuilds attachments, LLM input, and prompt for the actual call.
   - Spec reference: `docs/phase1/spec.md` sections 8.1, 8.4, 8.7, 9.4.
   - Judgment: Valid architecture risk. This is not a current proven behavior bug, but it weakens the ability to audit what was actually sent to the LLM.
   - Direction for next phase: create one canonical LLM request object or have the adapter return the actual request payload used for debug output.

6. P2: managed output replacement is not rollback-safe.
   - Evidence: `src/excel_semantic_md/output/writers.py:357-372` deletes existing managed outputs before moving staging contents into place.
   - Spec reference: `docs/phase1/spec.md` sections 9 and 11.
   - Judgment: Valid reliability risk. The spec does not explicitly require atomic replacement, but partial output is harmful for a CLI whose output artifacts are the user-facing contract.
   - Direction for next phase: replace outputs through a safer swap strategy or preserve old outputs until all new managed outputs are ready to move.

7. P3: common workbook pre-processing flow is duplicated across CLI/app paths.
   - Evidence: `src/excel_semantic_md/cli/main.py:276-329` and `src/excel_semantic_md/app/convert_pipeline.py:25-28` independently assemble read -> detect -> visual metadata -> link.
   - Spec reference: `docs/phase1/spec.md` sections 2.2, 2.3, 2.4.
   - Judgment: Valid but lower priority. This is not yet over-implementation, but duplicated orchestration increases the chance of future `inspect` / `render` / `convert` divergence.
   - Direction for next phase: consider a small shared pre-processing use case if touching this area for another fix.

Rejected or downgraded findings:

- `Pasteur` reported no spec mismatch. This is accepted as that reviewer result, but it does not invalidate the security/test/architecture findings above because those findings are either safety failures, regression gaps, or maintainability risks rather than broad functional contract mismatches.
- Copilot SDK local CLI behavior, vision attachment behavior, `.xlsm` live macro-disabled behavior, and Excel COM cleanup with existing user processes are not actionable code findings in this review because they require live confirmation. They are recorded as residual risks and pending evidence.
- `setup` real environment behavior is also a residual live-confirmation risk. The current implementation and existing validation notes cover diagnostic behavior in the available environment, but not every real machine state.

## Response Plan

- Implement accepted findings 1-6 because they affect safety, public output reliability, or regression coverage.
- Do not implement accepted finding 7 in this pass. The common pre-processing flow duplication is a lower-priority maintainability concern, not a current spec violation, and refactoring it now would broaden scope without direct user-facing benefit.
- Keep live-confirmation-only risks as residual risks rather than code findings.

## Applied Fixes

- `.xlsm` rendering now fails closed when Excel COM cannot set `AutomationSecurity = 3`; workbook opening is skipped in that case.
- `manifest.json` warning / failure details now redact local absolute path substrings even when they are embedded in generic exception strings.
- Managed output replacement now backs up existing managed outputs before publishing staging outputs and restores the old outputs if publishing fails.
- LLM request construction now uses a single `build_llm_request()` path for attachments, LLM input, and prompt; `convert` stores debug input from the same request object passed to the adapter.
- Added regression tests for `render` success stdout contract, non-generation of `result.md` / `manifest.json`, LLM non-invocation, malformed drawing XML warning-and-continue behavior, path redaction, output rollback, and prepared LLM request reuse.
- Post-fix review follow-up also broadened path redaction to common file paths with spaces, removed unnecessary absolute workbook/artifact paths from render failure details, asserted direct redaction keys, asserted LLM request attachments, and added a multi-sheet malformed drawing regression test.

## Validation

- `python -m pytest tests/test_render.py tests/test_ooxml_visual_reader.py tests/test_output.py tests/test_llm.py -q`
  - Result: passed (`45 passed`)
- `python -m pytest -q`
  - Result: passed (`95 passed`)

## Post-fix Subagent Review

- Initial launch attempt: failed because full-history forked agents cannot override agent type/model/reasoning options in this environment.
- spec compliance and functional correctness reviewer: completed
  - Agent: `Hilbert` (`019dba70-ff4e-7fb3-ada6-4449cbf8fcbd`)
  - Launch status: success after retry without fork context
  - Raw finding: P2 path redaction was incomplete for Windows/UNC paths containing spaces and for generic POSIX absolute paths.
  - MainAgent judgment: Valid. The accepted security fix should handle common local path forms, not only no-space paths.
  - Applied response: Added file-path regexes that handle common extensions with spaces and widened POSIX fallback redaction.
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Singer` (`019dba70-ff9c-7871-8814-af4710d38fc0`)
  - Launch status: success after retry without fork context
  - Raw findings:
    - P2 direct `path` / `workbook` / `temp_dir` redaction branches were not explicitly tested.
    - P2 prepared LLM request test did not assert attachments, ordering, limit, or absolute path normalization.
    - P2 malformed drawing XML test did not prove a following sheet still continues.
  - MainAgent judgment: Valid as regression coverage gaps. They do not change the product specification, but they make the accepted fixes harder to regress.
  - Applied response: Added direct-key redaction assertions, LLM attachment assertions, and a two-sheet malformed drawing fixture generated in test.
- reliability and security reviewer: completed
  - Agent: `Bernoulli` (`019dba70-ffc0-7581-aab3-fd442660b991`)
  - Launch status: success after retry without fork context
  - Raw findings:
    - P2 backup directory can remain if restore itself fails.
    - P2 render cleanup warnings include absolute workbook path.
    - P2 render failure details can include absolute artifact path.
  - MainAgent judgment:
    - Backup directory finding: rejected. If restoration itself fails, deleting the backup would risk losing the user's previous managed outputs. Leaving the backup is safer and preserves recovery data.
    - Render cleanup warning path: valid. The absolute workbook path is unnecessary in warning details.
    - Render failure artifact path: valid. Failure details do not need absolute artifact paths because `temp_dir` / successful artifact paths are already part of the render contract.
  - Applied response: Changed render cleanup warning workbook detail to workbook file name and render failure `path` details to artifact file names.

## Residual Risks

- Excel COM rendering, `Range.CopyPicture`, `Shape.CopyPicture`, and `Chart.Export` remain live-confirmation dependent.
- `.xlsm` macro-disabled behavior requires a real macro-enabled workbook confirmation.
- Copilot SDK local CLI behavior and vision attachment behavior remain pending.
- Attachment file path handling in the actual Copilot SDK/provider boundary is not proven by mock tests.
- Existing user Excel process isolation requires real-machine confirmation.
- Common workbook pre-processing flow remains duplicated across CLI/app paths by design for this pass.

## Pending Items

- Add or update live confirmation evidence for:
  - `.xlsm` macro-disabled behavior
  - Excel COM cleanup with existing user Excel processes
  - Copilot SDK local CLI behavior
  - vision attachment behavior
