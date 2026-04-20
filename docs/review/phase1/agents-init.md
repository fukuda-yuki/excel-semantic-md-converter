# agents-init review

## 2026-04-21

### Scope

- Initialize repository agent instructions after user clarification.
- Align AGENTS.md with the rule that `docs/phase1/spec.md` is the absolute Source of Truth.
- Review whether phase documents support that workflow.

### Changed files under review

- `AGENTS.md`
- `docs/phase1/plan.md`
- `docs/phase1/task.md`
- `docs/phase1/knowledge.md`
- `docs/review/phase1/agents-init.md`

### Subagents used

- `019dac5a-72f6-7db3-806a-e1088ca9cbf7`: Source of Truth, spec compliance, functional correctness, and `docs/phase1/*.md` usage review. Launch status: success.
- `019dac5a-731d-7883-95f6-4df0ddff6d72`: tests, edge cases, regression risk, review workflow, review note, `task.md`, and `knowledge.md` process review. Launch status: success.
- `019dac5f-57fb-7860-a5c1-dbb46c0752dc`: final Source of Truth and phase-document consistency review after fixes. Launch status: success.
- `019dac5f-584c-7ae2-a0f9-5dbd5276ace6`: final process audit review after fixes. Launch status: success.

### Raw findings summary

1. AGENTS.md still contained spec-like implementation constraints copied from README, even though `docs/phase1/spec.md` is defined as the absolute Source of Truth.
2. AGENTS.md listed README-derived CLI options as implementation guidance.
3. `docs/phase1/plan.md` still stated that README is the source of truth.
4. Mandatory subagent review lacked handling for subagent launch failure or unavailable subagents.
5. Review note creation was conditional, leaving no audit trail for zero findings, rejected findings, or failed review attempts.
6. Review note milestone selection was undefined outside commit creation.
7. `knowledge.md` could become a hidden specification store unless confirmed specifications are promoted to `spec.md`.
8. `task.md` update rules were defined only at completion, not at task start or when a checklist item is missing.
9. After fixes, `docs/phase1/plan.md` still read like an authoritative requirements document.
10. The review note changed-file list did not include all files changed by the task.
11. The review note did not record subagent roles and launch status.
12. `knowledge.md` still implied only valid findings should be recorded, conflicting with the all-results review audit rule.

### MainAgent validation

1. Valid. AGENTS.md must not act as a second product specification.
2. Valid. README-derived CLI details should live in `docs/phase1/spec.md` or remain non-authoritative background.
3. Valid. `docs/phase1/plan.md` line 4 conflicts with the clarified Source of Truth rule.
4. Valid. A mandatory review gate needs a failure policy.
5. Valid. Review evidence must be recorded even when findings are absent or rejected.
6. Valid. Review note path depends on milestone and must be known before review note creation.
7. Valid. `knowledge.md` is for shared context, not authoritative specification.
8. Valid. `task.md` should be checked at task start and updated at completion.
9. Valid. `plan.md` needs an explicit non-authoritative draft disclaimer.
10. Valid. Review evidence must list all reviewed and changed files.
11. Valid. Review evidence must show required reviewer roles and launch status.
12. Valid. `knowledge.md` must align with AGENTS.md and require recording zero findings, rejected findings, and launch failures.

### Response plan

- Remove README-derived product specification details from AGENTS.md and make AGENTS.md an operational guide.
- Add explicit subagent failure handling and audit trail requirements.
- Add milestone confirmation rules for review notes.
- Add rules for promoting confirmed specifications into `docs/phase1/spec.md`.
- Add task lifecycle rules for `docs/phase1/task.md`.
- Update `docs/phase1/plan.md` so it no longer says README is the source of truth.
- Record the user clarification in `docs/phase1/knowledge.md`.
- Update `docs/phase1/task.md` for this AGENTS initialization task.
- Make `docs/phase1/plan.md` explicitly non-authoritative until promoted into `docs/phase1/spec.md`.
- Update this review note with all changed files, subagent roles, launch status, and final review findings.
- Align `docs/phase1/knowledge.md` with the all-results review audit rule.

### Fixes applied

- Updated `AGENTS.md` to remove README-derived product specification details and keep it focused on operational rules.
- Added review workflow rules for milestone confirmation, mandatory review note creation even with no findings, subagent launch failure handling, and completion blocking.
- Added rules that confirmed specifications must be reflected in `docs/phase1/spec.md`, while `docs/phase1/knowledge.md` remains supporting context.
- Updated `docs/phase1/plan.md` so it no longer names README as the source of truth and marks README-derived details as requiring promotion into `docs/phase1/spec.md`.
- Added an explicit `docs/phase1/plan.md` disclaimer that the document is a draft of unpromoted candidates, not implementation specification.
- Updated `docs/phase1/knowledge.md` with the user-confirmed operating rules.
- Updated `docs/phase1/knowledge.md` to require recording zero findings, rejected findings, and subagent launch failures in review notes.
- Updated `docs/phase1/task.md` with the completed agent instruction initialization tasks.
- Updated this review note to include all changed files, required reviewer role mapping, launch status, and final review findings.

### Remaining risks

- `docs/phase1/spec.md` is still sparse relative to the detailed project plan. Implementation should not proceed until required behavior is specified there or the user explicitly directs how to promote plan/README content.

### Deferred items or open questions

- Decide whether to migrate detailed requirements currently present in README and `docs/phase1/plan.md` into `docs/phase1/spec.md`.
