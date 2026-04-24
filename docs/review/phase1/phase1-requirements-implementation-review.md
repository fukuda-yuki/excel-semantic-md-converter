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

## 2026-04-24 Re-review

### Review Scope

- Milestone: `phase1-requirements-implementation-review`
- Scope: current Phase 1 spec/implementation re-review requested by the user; no fixes in this pass
- Reviewed areas:
  - `docs/phase1/spec.md`
  - `docs/phase1/task.md`
  - `docs/phase1/knowledge.md`
  - `src/excel_semantic_md/**`
  - `tests/**`
- Out of scope: code fixes, live confirmation, commit/publish work
- Source of Truth: `docs/phase1/spec.md`

### Changed Files

- `docs/review/phase1/phase1-requirements-implementation-review.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`

### Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Volta` (`019dbb97-752b-7923-a6d6-e998cc42998a`)
  - Launch status: success
  - Result: 3 findings reported
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Carver` (`019dbb97-757b-7aa3-9786-7d0b8aa8c444`)
  - Launch status: success
  - Result: 6 findings reported
- reliability and security reviewer: completed
  - Agent: `Arendt` (`019dbb97-75ae-7b52-927d-0831767d5e7b`)
  - Launch status: success
  - Result: 4 findings and 3 residual risks reported
- architecture and over-implementation reviewer: completed
  - Agent: `Plato` (`019dbb97-75e8-7fa1-ab56-d18f0fc2360b`)
  - Launch status: success
  - Result: 4 findings and 1 deferred concern reported

### Raw Findings Summary

- `Volta`
  - P1: `convert` always renders cell-based blocks through Excel COM, so simple table/paragraph workbooks fail without COM despite Range images being supplemental.
  - P2: number formatting ignores most `number_format` cases, so display values can diverge from Excel.
  - P3: empty-sheet short-circuit may diverge from the `1 sheet = 1 session` wording.
- `Carver`
  - P2: `absoluteAnchor` is only covered via hand-built `VisualElement`, not real OOXML fixtures.
  - P2: several warning-only drawing relationship branches are untested.
  - P2: chart degradation warning branches are untested.
  - P2: attachment fallback for Copilot SDK call shape is untested.
  - P3: visible-only filtering is not covered with a more realistic filtered workbook fixture.
  - P3: display-value formatting branches are only partially covered.
- `Arendt`
  - P1: OOXML image targets are not validated as image content before copy/attachment, so a crafted workbook can expose non-image parts.
  - P2: unexpected COM/OSError exceptions abort the whole sheet render instead of degrading per artifact.
  - P2: `convert` always screenshots every cell-based block.
  - P2: omitting the image cap sends every rendered artifact to Copilot.
  - Residual risks: live confirmation is still needed for `.xlsm` macro-disabled behavior, COM cleanup with existing Excel processes, and real Copilot attachment handling.
- `Plato`
  - P1: `render` planning is shared into `convert`, creating an unnecessary hard Excel COM dependency for cell-based conversion.
  - P2: workbook pre-processing flow is duplicated across `convert` / `inspect` / `render`.
  - P2: debug-only payloads are materialized even when `--save-debug-json` is off.
  - P3: warning/failure types are fragmented across layers.

### MainAgent Validity Judgment

Accepted findings:

1. P1: `convert` hard-depends on Excel COM even for simple cell-based sheets, and `--max-images-per-sheet 0` does not remove that dependency.
   - Evidence:
     - `src/excel_semantic_md/render/planner.py:26-36` always adds `range_copy_picture` for `SourceKind.CELLS`.
     - `src/excel_semantic_md/app/convert_pipeline.py:166-205` always runs render before LLM whenever `plan_items` exist.
     - Local validation:
       - `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-review-convert`
       - `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-review-convert-zero --max-images-per-sheet 0`
       - In both cases, `result.md` reports `render: Excel COM rendering requires pywin32 modules (...)`.
   - Spec reference: `docs/phase1/spec.md` sections 7.2, 8.1, 8.3.
   - Judgment: Valid. Phase 1 says structured blocks are the primary source and Range images are supplemental. The current pipeline broadens the dependency surface beyond that contract.
   - Direction for next phase: split `convert` rendering needs from `render` live-confirmation needs; do not require cell screenshots when they are not needed, and honor `--max-images-per-sheet 0`.

2. P1: OOXML image targets are publishable/attachable without validating that the target part is actually an image.
   - Evidence:
     - `src/excel_semantic_md/excel/ooxml_visual_reader.py:431-477` records `content_type` for image candidates.
     - `src/excel_semantic_md/render/planner.py:59-79` accepts any non-null `target_part`.
     - `src/excel_semantic_md/render/excel_com_renderer.py:268-282` blindly copies the package part bytes.
     - `src/excel_semantic_md/llm/builders.py:26-32` may forward the resulting artifact to Copilot when the cap is unset.
   - Spec reference: `docs/phase1/spec.md` sections 3.2, 7.4, 8.3.
   - Judgment: Valid. This crosses the stated boundary that Phase 1 does not extract macro content and only treats workbook images as image assets.
   - Direction for next phase: require an image MIME/content-type allowlist before copying or attaching OOXML image targets.

3. P2: `render` CLI can crash on unexpected planning/render exceptions instead of returning JSON failures.
   - Evidence:
     - `src/excel_semantic_md/cli/main.py:322-337` has no normalization around `build_render_plan()` or `render_with_excel_com()`.
     - Local validation via an inline monkeypatch script showed `RuntimeError: plan boom` escaping from `cli_main.main(["render", ...])`.
   - Spec reference: `docs/phase1/spec.md` sections 2.4, 11.3.
   - Judgment: Valid. `render` is an externally supported diagnostic command whose contract is JSON with `warnings` / `failures`; uncaught exceptions break that contract.
   - Direction for next phase: normalize unexpected exceptions in `render` the same way `convert` normalizes sheet-level failures.

4. P2: per-artifact unexpected COM/OSError failures abort the rest of sheet rendering and degrade to a generic sheet-level failure.
   - Evidence:
     - `src/excel_semantic_md/render/excel_com_renderer.py:160-179` only catches `RenderTaskError` inside the artifact loop.
     - `_copy_package_part()` (`268-282`) and `_copy_object_to_png()` (`285-300`) can raise ordinary exceptions that escape to the outer `except`.
   - Spec reference: `docs/phase1/spec.md` sections 11.2, 11.3.
   - Judgment: Valid. The spec asks the tool to return failures in JSON as far as possible; current behavior stops at the first unwrapped exception and loses the rest of the artifact-level picture.
   - Direction for next phase: wrap ordinary render-time exceptions per item and continue the loop where safe.

5. P2: when `--max-images-per-sheet` is omitted, all rendered artifacts are sent to Copilot by default.
   - Evidence:
     - `src/excel_semantic_md/llm/builders.py:26-32` returns the full ranked list when the cap is `None`.
     - Combined with `src/excel_semantic_md/render/planner.py:26-36`, this includes every cell-range screenshot by default.
   - Spec reference: `docs/phase1/spec.md` section 8.3.
   - Judgment: Valid. The spec explicitly says not to send all images indiscriminately; the current default does exactly that for complex sheets.
   - Direction for next phase: add a default relevance gate and/or a conservative default cap, and keep Range screenshots opt-in or ambiguity-driven.

6. P2: display-value formatting does not satisfy the spec's `number_format`-aware contract for many common numeric formats.
   - Evidence:
     - `src/excel_semantic_md/excel/workbook_reader.py:379-386` only handles percent formats and integer-valued floats specially.
     - Currency, grouping, fixed decimals, and other display formats fall through to `str(value)`.
     - Current tests (`tests/test_workbook_reader.py:190-224`) only cover one percentage and one date-path example.
   - Spec reference: `docs/phase1/spec.md` section 3.5.
   - Judgment: Valid. The spec intentionally allows a conservative rendering of Excel display values, but it still requires number-format-aware stringification.
   - Direction for next phase: extend `_format_number()` for the supported conservative subset and add regression coverage for currency/grouping/decimal cases.

7. P2: key warning-only degradation paths are under-tested.
   - Evidence:
     - `src/excel_semantic_md/excel/ooxml_visual_reader.py:224-286` includes `sheet_drawing_relationships_missing`, `sheet_drawing_relationship_id_missing`, and `drawing_part_parse_failed`.
     - `src/excel_semantic_md/excel/ooxml_visual_reader.py:518-557` includes `chart_relationship_id_missing`, `chart_target_missing`, `chart_part_missing`, and `chart_part_parse_failed`.
     - Existing tests in `tests/test_ooxml_visual_reader.py` cover successful image/chart parsing, `drawing_part_missing`, and one malformed drawing case, but not the broader warning-only matrix.
     - `src/excel_semantic_md/llm/adapter.py:124-130` has an untested attachment-call fallback branch.
   - Spec reference: `docs/phase1/spec.md` sections 6.2, 8.3, 12.1.
   - Judgment: Valid as a regression-gap finding. These paths encode explicit Phase 1 behavior and should be protected.
   - Direction for next phase: add fixtures/tests for warning-only continuation paths and the attachment-call fallback.

Accepted but lower-priority finding:

8. P3: pre-processing flow is duplicated across `convert`, `inspect`, and `render`, and their warning aggregation already differs.
   - Evidence:
     - `src/excel_semantic_md/app/convert_pipeline.py:25-28`
     - `src/excel_semantic_md/cli/main.py:279-290`
     - `src/excel_semantic_md/cli/main.py:302-327`
   - Spec reference: `docs/phase1/spec.md` sections 2.2, 2.3, 2.4.
   - Judgment: Valid, but lower priority. This is a maintainability risk rather than the most urgent product defect.
   - Direction for next phase: only refactor this when touching the shared pre-processing boundary for a concrete fix.

Rejected or downgraded findings:

- Empty-sheet short-circuit (`src/excel_semantic_md/app/convert_pipeline.py:144-165`, `src/excel_semantic_md/output/writers.py:137-143`):
  - Judgment: Rejected as a bug finding. `docs/phase1/spec.md` does not require an empty visible sheet to call the external LLM, and the current behavior avoids an unnecessary provider call. This is also already recorded in `docs/phase1/knowledge.md`.
- Debug-only payloads are materialized even when `--save-debug-json` is off:
  - Judgment: Downgraded. This is a real over-implementation/performance concern, but there is not enough evidence here that it is currently harmful enough to outrank the accepted correctness and safety findings.
- Warning/failure type fragmentation across layers:
  - Judgment: Downgraded. The inconsistency is real, but this pass surfaced more direct contract failures and security issues.
- `absoluteAnchor` real-OOXML coverage gap and realistic filter-fixture gap:
  - Judgment: Downgraded into the broader test-gap finding. Both are valid coverage concerns, but they are narrower instances of the accepted regression-gap category.

### Response Plan

- Do not change code in this pass.
- Address next-phase fixes in this order:
  1. Narrow `convert` rendering/attachment behavior so cell-only conversion does not require unconditional screenshots or unbounded attachments.
  2. Validate OOXML image targets before copying/publishing/attaching them.
  3. Harden `render` failure normalization at CLI and per-artifact levels.
  4. Extend display-value formatting for the supported conservative `number_format` subset.
  5. Add regression tests for the accepted warning-only and attachment-fallback paths.
- Keep pre-processing deduplication as a follow-on refactor, not a prerequisite for the higher-priority fixes.

### Applied Fixes

- None in this pass. The user explicitly requested findings only; fixes are deferred to the next phase.

### Validation

- `python -m pytest -q`
  - Result: passed (`95 passed`)
- `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-review-convert`
  - Result: exit code `0`, but `result.md` reports `render: Excel COM rendering requires pywin32 modules (...)`
- `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-review-convert-zero --max-images-per-sheet 0`
  - Result: same sheet failure as above, confirming that `--max-images-per-sheet 0` does not bypass render dependency
- Inline monkeypatch validation against `cli_main.main(["render", ...])`
  - Result: unexpected `RuntimeError("plan boom")` escapes instead of producing JSON failure output

### Residual Risks

- `.xlsm` macro-disabled behavior still needs live confirmation with a real VBA-containing workbook.
- Excel COM cleanup with existing user Excel processes remains live-confirmation dependent.
- Actual Copilot SDK/provider handling of attachment count, file types, and non-image files remains an external-boundary risk.
- The accepted test gaps mean some warning-only continuation behavior is still less protected than the spec warrants.

### Pending Items

- In the next phase, implement the accepted findings before adding more feature scope.
- After those fixes, collect live confirmation evidence for:
  - `.xlsm` macro-disabled behavior
  - Excel COM cleanup with existing user Excel processes
  - Copilot SDK local CLI behavior
  - vision attachment behavior

## 2026-04-24 Re-review Fix Batch Implementation

### Review Scope

- Milestone: `phase1-requirements-implementation-review`
- Scope: accepted 2026-04-24 re-review findings for convert/render behavior, OOXML image safety, number formatting, and regression coverage
- Reviewed areas:
  - `src/excel_semantic_md/app/convert_pipeline.py`
  - `src/excel_semantic_md/render/planner.py`
  - `src/excel_semantic_md/render/excel_com_renderer.py`
  - `src/excel_semantic_md/cli/main.py`
  - `src/excel_semantic_md/llm/builders.py`
  - `src/excel_semantic_md/excel/workbook_reader.py`
  - `tests/test_output.py`
  - `tests/test_render.py`
  - `tests/test_ooxml_visual_reader.py`
  - `tests/test_llm.py`
  - `tests/test_workbook_reader.py`
  - `docs/phase1/spec.md`
  - `docs/phase1/task.md`
  - `docs/phase1/knowledge.md`
- Out of scope: live confirmation, pre-processing dedup refactor, commit/publish work
- Source of Truth: `docs/phase1/spec.md`

### Changed Files

- `src/excel_semantic_md/app/convert_pipeline.py`
- `src/excel_semantic_md/render/planner.py`
- `src/excel_semantic_md/render/excel_com_renderer.py`
- `src/excel_semantic_md/cli/main.py`
- `src/excel_semantic_md/llm/builders.py`
- `src/excel_semantic_md/excel/workbook_reader.py`
- `tests/test_output.py`
- `tests/test_render.py`
- `tests/test_ooxml_visual_reader.py`
- `tests/test_llm.py`
- `tests/test_workbook_reader.py`
- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-requirements-implementation-review.md`

### Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Euler` (`019dbf03-d166-7063-95c3-ad7759bfe7b8`)
  - Launch status: success
  - Result: 3 findings reported
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Feynman` (`019dbf03-d1ab-71a3-8948-159405179dd3`)
  - Launch status: success
  - Result: 3 findings reported
- reliability and security reviewer: completed
  - Agent: `Arendt` (`019dbf03-d1d8-7393-b06f-4b92588cad93`)
  - Launch status: success
  - Result: 3 findings reported
- architecture and responsibility-boundary reviewer: completed
  - Agent: `Dirac` (`019dbf03-d20d-7aa0-a759-89acad10a6e0`)
  - Launch status: success
  - Result: 3 findings reported

### Raw Findings Summary

- `Euler`
  - P1: scaling comma number formats such as `#,##0,` were still mis-rendered instead of falling back conservatively.
  - P2: dropping all cell-based render items means `convert --save-render-artifacts` would become a no-op for cell-only sheets.
  - P2: default attachment cap of 3 is not valid unless the spec is updated.
- `Feynman`
  - P1: cell-only convert path disables explicit `--save-render-artifacts` behavior.
  - P2: hidden default attachment cap should be reflected in the spec/contract.
  - P2: multi-section number-format rules remain incomplete.
- `Arendt`
  - P1: content-type-only checking was not a sufficient OOXML image allowlist.
  - P2: dropping all cell-based render items can remove layout evidence in convert.
  - P2: partial output files can remain in `temp_dir` after per-artifact failures.
- `Dirac`
  - P1: orchestration-level filtering of cell-based render items changes convert behavior significantly.
  - P2: attachment default policy lives in builder-level logic.
  - P2: scaling comma should fallback rather than render as grouped output.

### MainAgent Validity Judgment

Accepted findings:

1. P1/P2: explicit `--save-render-artifacts` for cell-only sheets must still preserve render artifacts.
   - Judgment: Valid. The user-requested plan removed default convert dependence on cell screenshots, but it did not justify making an explicit opt-in flag inert.
   - Response: `convert_pipeline` now keeps cell-based render items when `save_render_artifacts=True`, while the default path still skips them.

2. P1/P2: scaling comma formats must not be rendered as grouped output.
   - Judgment: Valid. This violated the agreed conservative formatting approach.
   - Response: `_format_number()` now falls back to raw stringification when the integer pattern ends with a scaling comma, and regression coverage was added.

3. P1: OOXML image allowlist must be stronger than `content_type.startswith("image/")`.
   - Judgment: Valid. A content-type-only check left room for spoofed package parts.
   - Response: planner-side allowlisting now also requires `xl/media/` placement plus a known safe extension/content-type mapping before planning `ooxml_image_copy`.

4. P2: per-artifact failures should not leave partial output files behind.
   - Judgment: Valid. Failure normalization is more trustworthy if partial files are removed before control returns.
   - Response: `_render_plan_item()` now removes the reserved output path on any exception, and regression coverage was added.

Rejected or downgraded findings:

- Default cell-only convert path lacking range screenshots: rejected as a bug. This is the intended behavior of the approved fix batch. The agreed plan explicitly removed default convert dependence on cell-based screenshots; only explicit `--save-render-artifacts` keeps them now.
- Default attachment cap of 3 being unspecified: rejected after spec update. `docs/phase1/spec.md` was updated in this pass to make the default explicit.
- Multi-section number-format behavior such as zero-empty sections: rejected for this batch. The approved plan explicitly scoped `_format_number()` to a conservative first-section subset rather than full Excel format fidelity.
- Builder-layer placement of the default attachment policy: downgraded. This is a maintainability concern, but not a current product-contract failure after the spec update and current single-caller design.

### Applied Fixes

- `convert` no longer requires Excel COM for cell-only/table-only/paragraph-only sheets by default, but explicit `--save-render-artifacts` still preserves cell-based render artifacts.
- `--max-images-per-sheet` default behavior is now a documented max of 3 major visuals, with range screenshots excluded from the default attachment set.
- OOXML original-image copy now requires a trusted `xl/media/` path and a known extension/content-type mapping before publish/attach planning.
- `render` CLI unexpected planning/render exceptions are normalized to JSON failures.
- `render_with_excel_com()` now normalizes ordinary per-artifact exceptions and removes partial output files on failure.
- Conservative number formatting now covers currency/grouping/fixed decimals and falls back for scaling-comma formats.
- Added regression tests for cell-only convert without COM, explicit render-artifact opt-in, attachment defaults, attachment fallback call shape, chart/drawing warning-only paths, trusted/untrusted OOXML image planning, partial-file cleanup, and number-format fallback.
- Updated `docs/phase1/spec.md`, `docs/phase1/task.md`, and `docs/phase1/knowledge.md` to reflect the finalized behavior.

### Validation

- `python -m pytest tests/test_llm.py tests/test_render.py tests/test_ooxml_visual_reader.py tests/test_workbook_reader.py tests/test_output.py -q`
  - Result: passed (`69 passed`)
- `python -m pytest -q`
  - Result: passed (`109 passed`)

### Residual Risks

- `.xlsm` macro-disabled behavior still needs live confirmation with a real VBA-containing workbook.
- Excel COM cleanup with existing user Excel processes remains live-confirmation dependent.
- Actual Copilot SDK/provider behavior for attachment handling remains an external-boundary risk.
- Full Excel multi-section number-format fidelity remains out of scope for this batch by explicit plan choice.

### Pending Items

- Collect live confirmation evidence for:
  - `.xlsm` macro-disabled behavior
  - Excel COM cleanup with existing user Excel processes
  - Copilot SDK local CLI behavior
  - vision attachment behavior
- Keep pre-processing deduplication across `convert` / `inspect` / `render` as a later refactor, not part of this fix batch.

## 2026-04-24 Requirements/Implementation Re-review After Fix Batch

### Review Scope

- Milestone: `phase1-requirements-implementation-review`
- Scope: 要件定義をもとにした current Phase 1 仕様・実装・テストの再レビュー
- Reviewed areas:
  - `docs/phase1/spec.md`
  - `docs/phase1/task.md`
  - `docs/phase1/knowledge.md`
  - `src/excel_semantic_md/**`
  - `tests/**`
  - `skills/excel-semantic-markdown/**`
- Out of scope: code fixes, spec fixes, task checklist updates, live confirmation
- Source of Truth: `docs/phase1/spec.md`

### Changed Files

- `docs/review/phase1/phase1-requirements-implementation-review.md`

### Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Halley` (`019dbf7f-adf5-7a63-b361-82db8d7aeb63`)
  - Launch status: success after initial fork option retry
  - Result: 3 findings reported
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Carson` (`019dbf7f-ae37-7ea0-a4ef-ab0e42bf4b53`)
  - Launch status: success after initial fork option retry
  - Result: 3 findings reported
- architecture and over-implementation reviewer: completed
  - Agent: `Jason` (`019dbf7f-ae6b-7c91-a1c0-3baafd3a51b6`)
  - Launch status: success after initial fork option retry
  - Result: 2 findings reported
- reliability and security reviewer: completed
  - Agent: `Poincare` (`019dbf7f-aebb-7923-9f1b-03bab039589c`)
  - Launch status: success after initial fork option retry
  - Result: 3 findings and live-confirmation residual risks reported

### Raw Findings Summary

- `Halley`
  - P1: text/image visual blocks still make normal `convert` depend on Excel COM through `shape_copy_picture`.
  - P2: filter visibility is implemented as hidden-row filtering only; actual autoFilter criteria are not evaluated.
  - P2: visual metadata omits spec-defined common fields when values are `None`.
- `Carson`
  - P2: table-only `convert` no-COM regression path is not directly tested.
  - P2: OOXML image missing-target / missing-part warning-and-skip paths are under-tested.
  - P3: Copilot attachment fallback API shape is fixed by automated tests despite live confirmation still pending.
- `Jason`
  - P3: empty sheet conversion synthesizes LLM success without an actual Copilot session, making the LLM boundary unclear.
  - P3: render-plan policy is split between planner and convert orchestration, making future divergence likely.
- `Poincare`
  - P2: LLM attachment payload uses absolute local paths at the Copilot SDK boundary.
  - P3: CLI diagnostics and parser errors print local paths.
  - P3: successful output replacement silently ignores backup cleanup failures.
  - Residual risks: `.xlsm` macro-disabled behavior, Excel COM cleanup with existing user processes, and real Copilot attachment behavior still require live confirmation.

### MainAgent Validity Judgment

Accepted findings:

1. P1: normal `convert` still depends on Excel COM for text-shape and image sheets.
   - Evidence:
     - `src/excel_semantic_md/render/planner.py` emits `shape_copy_picture` for shape blocks and both `ooxml_image_copy` and `shape_copy_picture` for image blocks.
     - `src/excel_semantic_md/app/convert_pipeline.py` filters only `SourceKind.CELLS` from the normal convert plan, so shape/image items remain.
     - Local reproduction:
       - `python -m excel_semantic_md.cli.main convert --input tests\fixtures\visuals\table-shape-visual.xlsx --out .tmp-review-shape`
       - `python -m excel_semantic_md.cli.main convert --input tests\fixtures\visuals\table-image-visual.xlsx --out .tmp-review-image`
       - Both commands exited `0` in non-strict mode, but `result.md` reported sheet failure: `render: Excel COM rendering requires pywin32 modules (...)`.
   - Spec reference: `docs/phase1/spec.md` sections 7.3, 7.4, 8.1, 14.
   - Judgment: Valid. The previous fix removed unconditional cell-range screenshots from normal convert, but the same over-dependency remains for common Phase 1 target workbooks containing text shapes or images. Text-in-shape should prefer extracted text, and OOXML image assets should not require COM when a trusted image part exists.
   - Direction for next phase: split normal convert visual asset needs from render live-confirmation needs. For shape/image, require COM only when the visual appearance itself is needed and cannot be satisfied from extracted text or trusted OOXML original assets.

2. P2: filtered-row visibility is underspecified for implementation and likely incomplete.
   - Evidence: `src/excel_semantic_md/excel/workbook_reader.py` excludes rows from XML only when row metadata has `hidden="1"`; it does not evaluate `autoFilter` criteria.
   - Spec reference: `docs/phase1/spec.md` section 3.3 requires rows hidden by filters to be skipped.
   - Judgment: Valid, but this is partly a specification-design issue. Reconstructing Excel's currently visible filtered view from OOXML without Excel may be unreliable; requiring full filter evaluation may over-expand Phase 1.
   - Direction for next phase: either implement a clearly bounded saved-hidden-row interpretation and update `spec.md`, or explicitly add Excel/live-confirmation dependence for true filter-view reproduction.

3. P2: visual metadata schema omits spec-defined fields when values are `None`.
   - Evidence: `src/excel_semantic_md/excel/ooxml_visual_reader.py` omits `relationship_id`, `target_part`, `source_part`, `extension`, `content_type`, and some chart-series fields from `to_dict()` when absent.
   - Spec reference: `docs/phase1/spec.md` section 6.1 says visual elements have common fields and chart series fields.
   - Judgment: Valid as a schema-contract ambiguity. The implementation is reasonable for compact JSON, but `inspect` is a supported structural output; consumers need to know whether absent and `null` are equivalent.
   - Direction for next phase: either emit explicit `null` for spec-defined fields or narrow the spec to state that optional fields may be omitted when not available.

4. P2: table-only no-COM convert regression coverage is missing.
   - Evidence: current no-COM convert regression coverage exercises `no-visuals.xlsx` paragraph/cell-only behavior, but does not directly lock table-only behavior after the fix batch.
   - Spec reference: `docs/phase1/spec.md` section 14 says cell-only / table-only / paragraph-only sheets with empty render items proceed to LLM without Excel COM.
   - Judgment: Valid as a test gap.
   - Direction for next phase: add a table-only convert test that proves rendering is skipped and LLM receives table block input.

5. P2: OOXML image missing-target / missing-part warning-and-skip paths are under-tested.
   - Evidence: tests cover non-image content type and untrusted image part, but not `image_target_missing` / `image_part_missing` flowing through render planning as skip-with-warning.
   - Spec reference: `docs/phase1/spec.md` section 14 requires non-image content type, missing target, and missing part to be warning-and-skip.
   - Judgment: Valid as a regression gap.
   - Direction for next phase: add OOXML fixtures or synthetic visual tests for missing relationship target and missing package part.

6. P3: empty-sheet conversion records synthetic LLM success.
   - Evidence: `src/excel_semantic_md/app/convert_pipeline.py` creates `LlmRunResult(status="succeeded", attempts=1, raw={"generated_by": "empty_sheet_short_circuit"})` without calling Copilot.
   - Spec reference: `docs/phase1/spec.md` section 8.1 says `1 sheet = 1 Copilot SDK session` is the basic rule; section 14 only clarifies render skipping, not LLM success synthesis.
   - Judgment: Valid as a spec/manifest clarity issue, not a behavioral blocker. Avoiding an external LLM call for an empty visible sheet is pragmatic, but `attempts=1` and `status=succeeded` can be mistaken for a real provider result.
   - Direction for next phase: represent this as `llm.status = skipped` with an explicit empty-sheet reason, or update `spec.md` to define synthetic empty-sheet success.

7. P3: render-plan policy is split between planner and convert orchestration.
   - Evidence: `build_render_plan()` plans cell-range items universally, while `convert_pipeline.py` removes cell-sourced items for normal convert and `render` CLI keeps them.
   - Spec reference: `docs/phase1/spec.md` sections 2.2, 2.4, 7.2, and 14.
   - Judgment: Valid maintainability risk. Current behavior matches the recent clarification, but the policy is distributed and easy to regress.
   - Direction for next phase: expose planner mode or a convert-specific planning helper so normal convert, `render`, and debug render-plan payloads do not drift.

8. P3: successful output replacement can silently leave old backup data if backup cleanup fails.
   - Evidence: `src/excel_semantic_md/output/writers.py` removes `.excel-semantic-md-backup-*` with `ignore_errors=True` after successful publication.
   - Spec reference: `docs/phase1/spec.md` sections 9.4, 9.5, 11.3, and security guidance on generated artifacts.
   - Judgment: Valid low-priority reliability/security risk. The backup protects rollback, but after success a cleanup failure may leave old generated artifacts or debug data without any user-visible warning.
   - Direction for next phase: surface backup cleanup failure as a warning or choose a documented retention/cleanup policy.

Rejected or downgraded findings:

- LLM attachment absolute paths at the Copilot SDK boundary:
  - Judgment: Downgraded to residual external-boundary risk. The LLM input JSON intentionally stores only file names, while the SDK attachment payload likely needs local absolute paths to read files. The real provider behavior remains pending live confirmation, so this is not accepted as a code finding in isolation.
- CLI diagnostics and parser errors printing local paths:
  - Judgment: Rejected as a product defect for this pass. `setup` is a local diagnostic command, and parser errors benefit from showing the path the user supplied. The `logs/` and `manifest.json` path-minimization requirements do not automatically prohibit interactive CLI diagnostics. If this tool is later run in centralized log collection, the spec should be revisited.
- Copilot fallback API shape test:
  - Judgment: Downgraded. The test may overfit an unconfirmed SDK shape, but it currently protects the adapter's intended fallback behavior. Treat as part of the broader pending live-confirmation risk rather than a standalone defect.

### Response Plan

- Do not change code in this pass.
- Address next-phase fixes in this order:
  1. Remove normal `convert`'s unconditional COM dependency for text-shape and trusted OOXML image workbooks.
  2. Decide and document the Phase 1 filter visibility contract before expanding implementation.
  3. Resolve visual metadata optional/null schema behavior.
  4. Add missing regression tests for table-only no-COM convert and image missing-target/missing-part warning paths.
  5. Clarify empty-sheet LLM status and consolidate render planning policy.
  6. Surface or document output backup cleanup failures.

### Applied Fixes

- None in this pass. The user explicitly requested findings only; fixes are deferred to the next phase.

### Validation

- `python -m pytest -q`
  - Result: passed (`109 passed`)
- `python -m excel_semantic_md.cli.main convert --input tests\fixtures\visuals\table-shape-visual.xlsx --out .tmp-review-shape`
  - Result: exit code `0` in non-strict mode, but `result.md` reports render failure because `pywin32` is unavailable.
- `python -m excel_semantic_md.cli.main convert --input tests\fixtures\visuals\table-image-visual.xlsx --out .tmp-review-image`
  - Result: exit code `0` in non-strict mode, but `result.md` reports render failure because `pywin32` is unavailable.
- Temporary `.tmp-review-shape` and `.tmp-review-image` outputs were removed after inspection.

### Residual Risks

- `.xlsm` macro-disabled behavior still needs live confirmation with a real VBA-containing workbook.
- Excel COM cleanup with existing user Excel processes remains live-confirmation dependent.
- Actual Copilot SDK/provider behavior for attachment handling remains an external-boundary risk.
- Full Excel filter-view reproduction may be impractical without Excel; this should be resolved as a spec decision before implementation.

### Pending Items

- Fixes are intentionally deferred to the next phase.
- After fixes, collect live confirmation evidence for:
  - `.xlsm` macro-disabled behavior
  - Excel COM cleanup with existing user Excel processes
  - Copilot SDK local CLI behavior
  - vision attachment behavior

## 2026-04-25 Requirements/Implementation Re-review Fix Implementation

### Review Scope

- Milestone: `phase1-requirements-implementation-review`
- Scope: accepted findings from `2026-04-24 Requirements/Implementation Re-review After Fix Batch`
- Reviewed areas:
  - `docs/phase1/spec.md`
  - `docs/phase1/task.md`
  - `docs/phase1/knowledge.md`
  - `src/excel_semantic_md/app/convert_pipeline.py`
  - `src/excel_semantic_md/render/excel_com_renderer.py`
  - `src/excel_semantic_md/excel/ooxml_visual_reader.py`
  - `src/excel_semantic_md/output/writers.py`
  - `tests/test_output.py`
  - `tests/test_ooxml_visual_reader.py`
- Source of Truth: `docs/phase1/spec.md`

### Changed Files

- `docs/phase1/spec.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/phase1-requirements-implementation-review.md`
- `src/excel_semantic_md/app/convert_pipeline.py`
- `src/excel_semantic_md/render/excel_com_renderer.py`
- `src/excel_semantic_md/excel/ooxml_visual_reader.py`
- `src/excel_semantic_md/output/writers.py`
- `tests/test_output.py`
- `tests/test_ooxml_visual_reader.py`

### Subagents

- spec compliance and functional correctness reviewer:
  - Agent: `Peirce` (`019dbf8e-16c7-72f3-b10b-4ee7b07cb6be`)
  - Launch status: failed due usage limit.
  - Result: unavailable.
- tests, edge cases, and regression risk reviewer:
  - Agent: `Zeno` (`019dbf8e-16fa-73b0-9aff-6fbc1e154e81`)
  - Launch status: failed due usage limit.
  - Result: unavailable.
- reliability and security reviewer:
  - Agent: `Hegel` (`019dbf8e-44a5-70d2-b8f9-39b9cf58be4f`)
  - Launch status: failed due usage limit.
  - Result: unavailable.

User follow-up `continue` was received after the subagent failure notifications. MainAgent proceeded with an explicit alternative self-review because the required subagent review results could not be obtained in this environment.

### Raw Findings Summary

- No subagent findings were available because all post-fix review agents failed to start or complete due usage limits.
- MainAgent self-review focused on the current diff against `docs/phase1/spec.md`, the accepted findings, and regression coverage.

### MainAgent Validity Judgment

Accepted findings addressed:

1. P1: normal `convert` still depended on Excel COM for text-shape and image sheets.
   - Response:
     - `convert_pipeline` now filters default convert render items through `_is_default_convert_render_item()`.
     - Cell-based `range_copy_picture` and `shape_copy_picture` items are excluded from normal convert unless `--save-render-artifacts` is explicitly set.
     - `render_with_excel_com()` now processes `ooxml_image_copy` outside the Excel COM session and skips COM diagnostics when no COM-backed render items remain.
   - Result: text shape sheets can proceed using extracted shape text without COM; trusted OOXML image assets can be copied and attached without COM.

2. P2: filtered-row visibility was underspecified and likely incomplete.
   - Response:
     - `docs/phase1/spec.md` now limits Phase 1 workbook-reading visibility to OOXML-saved hidden row / column state.
     - Full `autoFilter` condition re-evaluation is explicitly not a Phase 1 required implementation and is left to live confirmation or a later phase.
   - Result: the spec no longer implies a broad Excel filter engine in Phase 1.

3. P2: visual metadata omitted spec-defined fields when values were `None`.
   - Response:
     - `VisualSource.to_dict()`, `AssetCandidate.to_dict()`, and `ChartSeries.to_dict()` now emit defined fields with `null` when unavailable.
     - `docs/phase1/spec.md` now states these defined fields are emitted with `null` or empty arrays.
   - Result: `inspect` visual metadata has a more stable schema.

4. P2: table-only no-COM convert regression coverage was missing.
   - Response: added `test_run_convert_pipeline_skips_render_for_table_only_sheet()`.

5. P2: OOXML image missing-target / missing-part warning-and-skip paths were under-tested.
   - Response: added tests for `image_target_missing` and `image_part_missing` flowing into render planning as original-image skip warnings.

6. P3: empty-sheet conversion recorded synthetic LLM success.
   - Response:
     - empty visible sheets now set `markdown=""` without creating a synthetic `LlmRunResult`.
     - `docs/phase1/spec.md` states empty visible sheets skip provider LLM and surface as `llm` skipped.
   - Result: provider success is no longer confused with empty-sheet short-circuit behavior.

7. P3: render-plan policy was split between planner and convert orchestration.
   - Response: partially addressed by centralizing normal-convert filtering in `_is_default_convert_render_item()`. A fuller planner-mode refactor remains intentionally deferred to avoid broadening this fix batch.

8. P3: successful output replacement could silently leave old backup data if backup cleanup failed.
   - Response:
     - successful backup cleanup failure now emits a `RuntimeWarning` instead of being ignored.
     - regression coverage was added.

Rejected or deferred items:

- Full render planner mode refactor remains deferred. Current changes are narrowly scoped and covered by tests.
- Copilot SDK attachment boundary and CLI diagnostic path output remain residual risks / product-scope decisions from the previous review.

### Applied Fixes

- Removed normal `convert` COM requirement for text-shape sheets.
- Enabled trusted OOXML image original asset copy / attach without opening Excel COM.
- Kept chart rendering on the Excel COM path because `Chart.Export` remains the Phase 1 primary chart image path.
- Clarified filter visibility scope in `docs/phase1/spec.md`.
- Made visual metadata `source`, `asset_candidate`, and chart series fields stable with explicit `null` values.
- Changed empty visible sheets from synthetic LLM success to provider-skipped successful sheet output.
- Surfaced managed-output backup cleanup failures as `RuntimeWarning`.
- Added regression tests for table-only no-COM convert, text-shape no-COM convert, trusted OOXML image copy without COM, image missing-target / missing-part skip paths, and backup cleanup warnings.
- Updated `docs/phase1/task.md` and `docs/phase1/knowledge.md`.

### Validation

- `python -m pytest tests/test_output.py tests/test_ooxml_visual_reader.py tests/test_render.py -q`
  - Result: passed (`49 passed`)
- `python -m pytest -q`
  - Result: passed (`115 passed`)

### Residual Risks

- `.xlsm` macro-disabled behavior still needs live confirmation with a real VBA-containing workbook.
- Excel COM cleanup with existing user Excel processes remains live-confirmation dependent.
- Actual Copilot SDK/provider behavior for attachment handling remains an external-boundary risk.
- Full Excel filter-view reproduction remains out of Phase 1 scope unless the spec is changed again.
- Post-fix subagent review could not be completed because all spawned agents hit usage limits; this review note records the failure and the MainAgent alternative review.

### Pending Items

- Collect live confirmation evidence for:
  - `.xlsm` macro-disabled behavior
  - Excel COM cleanup with existing user Excel processes
  - Copilot SDK local CLI behavior
  - vision attachment behavior
- If agent capacity becomes available later, run a follow-up subagent review over this fix batch.

## 2026-04-25 Post-fix Subagent Review Rerun

### Review Scope

- Milestone: `phase1-requirements-implementation-review`
- Scope: post-fix subagent review rerun requested by the user after previous usage-limit failures.
- Reviewed areas:
  - `docs/phase1/spec.md`
  - current git diff
  - `src/excel_semantic_md/**`
  - `tests/**`
- Source of Truth: `docs/phase1/spec.md`

### Changed Files

- `src/excel_semantic_md/output/writers.py`
- `tests/test_output.py`
- `docs/review/phase1/phase1-requirements-implementation-review.md`

### Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Pascal` (`019dc00d-3ab9-7442-a161-e95db11e8777`)
  - Launch status: success
  - Result: no findings
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Goodall` (`019dc00d-3a7c-7273-a665-d8bc15571ba2`)
  - Launch status: success
  - Result: no findings
- reliability and security reviewer: completed
  - Agent: `Curie` (`019dc00d-3ae7-7571-8195-515aca8e3e62`)
  - Launch status: success
  - Result: 1 finding reported

### Raw Findings Summary

- `Pascal`
  - No findings.
  - Ran `python -m pytest tests/test_output.py tests/test_ooxml_visual_reader.py`; result `35 passed`.
- `Goodall`
  - No findings.
  - Ran selected relevant tests; result `6 passed`.
- `Curie`
  - P3: backup cleanup `RuntimeWarning` includes the absolute backup directory path, which can expose the user's output parent path in stderr or collected logs.

### MainAgent Validity Judgment

Accepted findings:

1. P3: backup cleanup warning leaked a local absolute path.
   - Evidence: `src/excel_semantic_md/output/writers.py` warning text included `backup_dir`, an absolute temp directory path under the output parent.
   - Spec reference: `docs/phase1/spec.md` sections 9.5 and security guidance around generated artifacts / logs.
   - Judgment: Valid. The previous change intentionally surfaced cleanup failure, but the warning does not need the absolute path to be actionable.
   - Response: Warning text now includes only `backup_dir.name`, preserving the backup marker without exposing the output parent path.

Rejected or downgraded findings:

- None from this rerun.

### Applied Fixes

- Changed managed-output backup cleanup `RuntimeWarning` to include only the backup directory basename.
- Added regression assertion that the warning text does not include the pytest temp/output parent path.

### Validation

- `python -m pytest tests/test_output.py tests/test_ooxml_visual_reader.py -q`
  - Result: passed (`35 passed`)
- `python -m pytest -q`
  - Result: passed (`115 passed`)

### Residual Risks

- `.xlsm` macro-disabled behavior still needs live confirmation with a real VBA-containing workbook.
- Excel COM cleanup with existing user Excel processes remains live-confirmation dependent.
- Actual Copilot SDK/provider behavior for attachment handling remains an external-boundary risk.

### Pending Items

- Collect live confirmation evidence for:
  - `.xlsm` macro-disabled behavior
  - Excel COM cleanup with existing user Excel processes
  - Copilot SDK local CLI behavior
  - vision attachment behavior
