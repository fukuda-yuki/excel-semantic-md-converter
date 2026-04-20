# Repository Guidelines

**output the final response in Japanese**

## Source of Truth
- This repository is docs-first.
- `docs/phase1/spec.md` is the highest-priority source of truth for product behavior and implementation constraints.
- Use `docs/phase1/knowledge.md` for verified observations from real logs.
- Use `docs/phase1/task.md` as the implementation checklist.
- If this file conflicts with `docs/phase1/spec.md`, follow `docs/phase1/spec.md` and explicitly mention the discrepancy in the final response.

## Project Structure



## Build, Test, and Development Commands
Run commands from the repository root.

- `dotnet build vs-log-watch.slnx`
- `dotnet test vs-log-watch.slnx`
- `dotnet run --project src/CopilotLogWatcher.Service`

Use the SDK version pinned by `global.json`.

## Coding Constraints
- Target `.NET 10`.
- Keep nullable reference types and implicit usings enabled.
- Use 4-space indentation.
- Keep block splitting, parsing, aggregation, queueing, persistence, and diagnostics concerns separated.
- Prefer clear, explicit code over clever abstractions.
- Do not broaden Phase 1 scope unless the user explicitly changes the scope.
- In particular, do not introduce UI, search API, notification features, VS Code log parsing, Copilot CLI log parsing, inline completion, `.vs\...\copilot-chat\...\sessions` parsing, or `ActivityLog.xml` as a primary source.

## Testing Guidelines
- Place tests in `tests/CopilotLogWatcher.Tests`.
- Cover block splitting, parsing, turn aggregation, fingerprints, SQLite queue behavior, retention cleanup, shared file reads, and incomplete block carry-over.
- Do not commit raw Visual Studio Copilot logs as fixtures.
- Use minimal sanitized or synthetic samples that preserve only the shape required for the test.

## Validation and Live Confirmation
Prefer automated validation first, but do not claim end-to-end success when the behavior depends on a live environment that you cannot verify directly.

You must ask the user to perform a live confirmation step when any of the following is true:
- the change depends on actual Visual Studio or GitHub Copilot behavior
- the change depends on fresh log generation under user profile log directories
- the change depends on `FileSystemWatcher` timing, file locks, shared reads, truncation, rotation, or partial writes
- the change depends on Windows service installation, service account behavior, or machine-specific permissions
- the change depends on actual PostgreSQL connectivity or machine-specific configuration

When live confirmation is required:
- do not pretend the scenario was fully verified
- give the user concrete reproduction or confirmation steps
- state exactly what evidence is needed, such as generated log blocks, database rows, diagnostic log lines, or reproduced error messages

## Mandatory Review Workflow
After every source-code change, perform a review before finalizing the task.

Review process:
1. Delegate review to multiple subagents in parallel.
2. At minimum, run:
   - one subagent for spec compliance and functional correctness
   - one subagent for tests, edge cases, and regression risk
3. For broad or risky changes, add another subagent for architecture, reliability, or security review.
4. The MainAgent must collect all findings, remove duplicates, validate each finding, and reject invalid findings with reasons.
5. The MainAgent must not apply fixes blindly. Each fix must be justified against the code, tests, and `docs/phase1/spec.md`.
6. Before finalizing the task, confirm that `docs/phase1/task.md` reflects the current implementation and validation status.

If the environment cannot launch subagents, explicitly state that limitation and perform a structured self-review using the same categories.

## Review Output
When review finds issues that require fixes, documentation, or follow-up, write a review note under `docs/review/phase1/`.

File rule:
- create or update one Markdown file per milestone
- path: `docs/review/phase1/<milestone>.md`

The review note must include:
- scope of the reviewed change
- changed files
- subagents used
- raw findings summary
- MainAgent validation result for each finding
- fixes applied
- remaining risks
- deferred items or open questions

Do not overwrite useful prior review history silently. Append a dated section when updating an existing milestone review file.

## Git Operations and Commit Rules
- Agents may create local commits in this repository.
- Agents must not push branches or tags.
- Agents must not open, update, merge, or auto-merge pull requests.
- Agents must not rewrite remote history.
- If the requested task does not justify a commit, do not create one.

Create a commit only after implementation, validation, and review are complete, unless the user explicitly asks for a checkpoint commit.

If the milestone is unclear, do not guess. Ask the user before creating a commit.

Commit messages must start with the milestone name, then follow Conventional Commits format.

Required format:
- `<milestone>: <type>(<scope>): <subject>`
- `<milestone>: <type>: <subject>` when scope is unnecessary

Examples:
- `phase1-m1: feat(parser): split multiline Copilot response blocks`
- `phase1-m1: fix(aggregator): avoid binding ambiguous raw responses`
- `phase1-m2: test(queue): add retention cleanup coverage`
- `phase1-m2: docs(review): record validated subagent findings`

Use valid Conventional Commits types such as:
- `feat`
- `fix`
- `refactor`
- `test`
- `docs`
- `chore`

Keep the subject specific and imperative. Do not use vague messages such as `update`, `fix issues`, or `work in progress`.

## Security and Configuration
- Never commit secrets, local database files, runtime logs, or connection strings.
- Keep PostgreSQL connection strings out of tracked files.
- Follow `.gitignore` and do not weaken log or database exclusions unless the user explicitly requests it.

## Agent-Specific Instructions
- Use installed Codex skills proactively when they match the task.
- When changing source code, review must happen before the final response.
- Ask the user for live confirmation instead of making unsupported claims about real-world runtime behavior.