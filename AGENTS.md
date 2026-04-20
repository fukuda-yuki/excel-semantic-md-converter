# Repository Guidelines

**最終応答は日本語で書くこと。**

## Source of Truth
- このリポジトリのプロダクト仕様・実装制約・優先順位の絶対的な Source of Truth は `docs/phase1/spec.md`。
- `docs/phase1/spec.md` と他の文書・実装メモ・このファイルが矛盾する場合は、必ず `docs/phase1/spec.md` を優先し、最終応答で矛盾点を明示する。
- `README.md` はプロジェクト背景・全体像・初期構想の参考資料として扱う。仕様判断の根拠にしてよいのは、`docs/phase1/spec.md` に反映済みの内容だけ。
- `docs/phase1/plan.md` は実装計画、分割方針、着手順序を共有するために使う。
- `docs/phase1/task.md` は実装チェックリストと完了状況を管理するために使う。
- `docs/phase1/knowledge.md` はエージェント間で共有すべき前提、調査結果、判断理由、ユーザー確認済み事項を記録するために使う。
- `docs/phase1/*.md` の内容が互いに矛盾する場合は、優先順位を `spec.md`、`plan.md`、`task.md`、`knowledge.md` の順に扱い、必要ならユーザーへ確認する。

## Clarification Rules
- 疑問がある場合は立ち止まってユーザーへ確認する。推測で進めない。
- ユーザーへの迎合は不要。技術的・仕様的に不明確な点、矛盾、リスクは明確に指摘する。
- `docs/phase1/spec.md` に仕様が存在しない場合、README や既存実装から勝手に仕様を補完しない。軽微で可逆な作業を除き、ユーザーへ確認する。
- ユーザー確認で得た、以後のエージェント間共有が必要な情報は、必ず `docs/phase1/knowledge.md` に記録する。

## Scope Discipline
- AGENTS.md はエージェントの作業手順を定める文書であり、プロダクト仕様を定義する場所ではない。
- アーキテクチャ、CLI、入出力、対象ファイル、非採用方針、受け入れ条件などの実装判断は `docs/phase1/spec.md` に従う。
- README や `docs/phase1/plan.md` に具体的な案があっても、`docs/phase1/spec.md` に昇格されていない限り、実装仕様として扱わない。
- `docs/phase1/spec.md` に必要な仕様が不足している場合は、実装へ進む前にユーザーへ確認する。
- 確認によって仕様が確定した場合は `docs/phase1/spec.md` に反映し、補足的な背景や判断理由を `docs/phase1/knowledge.md` に記録する。

## Build, Test, and Development Commands
Run commands from the repository root.

- Use commands defined by `docs/phase1/spec.md`, project configuration files, or existing repository scripts.
- If no command is defined, inspect the project files and ask the user before inventing a new required workflow.
- Do not carry over commands from another project.
- Exercise the externally supported interface defined in `docs/phase1/spec.md` whenever behavior must be reproducible.

## Testing Guidelines
- Derive test scope from `docs/phase1/spec.md` and `docs/phase1/task.md`.
- Use small, sanitized or synthetic fixtures. Do not commit confidential user data or generated runtime artifacts.
- Cover the behavior changed by the task, plus adjacent edge cases and regression risks identified during review.
- Isolate external tool, network, local machine, and live-service boundaries so ordinary automated tests remain deterministic.
- When behavior cannot be verified automatically, document the required live confirmation steps.

## Validation and Live Confirmation
Prefer automated validation first, but do not claim end-to-end success when behavior depends on the local machine or an external interactive tool.

Ask the user for live confirmation when the change depends on:

- software, credentials, account state, or permissions on the user's machine
- behavior of external tools or services that cannot be reproduced in automated tests
- file locks, local filesystem timing, generated artifacts, or machine-specific paths
- any live integration explicitly called out by `docs/phase1/spec.md`

When live confirmation is required, give concrete reproduction steps and state the evidence needed, such as generated `result.md`, `manifest.json`, asset files, debug JSON, hook logs, CLI output, or reproduced error messages.

## Review Workflow
実装を完了させる前に、必ず複数のサブエージェントを並列で立ち上げてレビューさせる。

Minimum required subagents:

- spec compliance and functional correctness reviewer
- tests, edge cases, and regression risk reviewer

For broad or risky changes, add one or more reviewers for:

- architecture and separation of responsibilities
- reliability and cleanup, especially Excel COM process handling
- security, especially `.xlsm`, file paths, generated artifacts, and credentials

MainAgent responsibilities:

- レビュー開始前に milestone を確認する。milestone が不明な場合は、`docs/review/phase1/<milestone>.md` を作成する前にユーザーへ確認する。
- MainAgent はレビュー中、取りまとめに徹する。
- サブエージェントの指摘事項を重複排除し、各指摘の妥当性を `docs/phase1/spec.md`、関連テスト、変更内容に照らして検証する。
- 妥当でない指摘は、理由を明記して棄却する。
- レビュー結果は、指摘の有無や採否に関わらず、修正前にまず `docs/review/phase1/<milestone>.md` に記録する。
- review note には、レビュー範囲、変更ファイル、使用したサブエージェント、subagent の起動成否、raw findings summary、MainAgent の妥当性判断、対応方針、適用した修正、残リスク、保留事項を含める。
- 指摘事項へ対応した場合は、対応後に同じ review note を更新する。
- 修正は盲目的に適用しない。必ずコード、テスト、`docs/phase1/spec.md` に照らして必要性を説明できるものだけ対応する。
- サブエージェントを起動できない、または必要数のレビュー結果を取得できない場合は、実装完了を宣言しない。ユーザーへ状況を報告し、代替レビューへ進む明示許可を得る。

Completion requirements:

- 作業開始時に `docs/phase1/task.md` を確認し、現在の作業に該当する項目があるか確認する。該当項目がなく、タスク化が必要な場合は追加するかユーザーへ確認する。
- 実装完了を宣言する前に、`docs/phase1/task.md` を現在の実装・検証状況に合わせて更新する。
- レビュー指摘に対応した場合は、`docs/review/phase1/<milestone>.md` も更新する。
- エージェント間で共有すべき新しい前提、判断、調査結果、確認済み事項が生じた場合は、必ず `docs/phase1/knowledge.md` に記録する。
- 仕様として確定した事項は `docs/phase1/spec.md` に反映する。`docs/phase1/knowledge.md` だけに仕様を退避しない。
- `docs/phase1/task.md`、`docs/phase1/knowledge.md`、`docs/phase1/spec.md`、review note を必要に応じて更新できない理由がある場合は、完了宣言せず、理由と必要な確認事項をユーザーへ提示する。

## Security and Configuration
- Never commit secrets, credentials, local connection strings, confidential user files, generated logs, debug dumps with user data, or rendered/generated artifacts from private inputs.
- Keep local outputs out of commits unless they are intentional sanitized fixtures.
- Do not weaken `.gitignore` protections for runtime outputs or private data.
- Treat executable or macro-capable input formats as potentially unsafe when they are in scope. Do not execute embedded code unless `docs/phase1/spec.md` explicitly requires it and the user confirms the risk.
- Prefer read-only input access where possible.

## Git Operations and Commit Rules
- Agents may create local commits only when the requested task justifies it and validation/review is complete.
- Do not push branches or tags.
- Do not open, update, merge, or auto-merge pull requests.
- Do not rewrite remote history.
- If the milestone is unclear, ask before creating a commit.

Commit messages must start with a milestone name and then use Conventional Commits:

- `<milestone>: <type>(<scope>): <subject>`
- `<milestone>: <type>: <subject>` when scope is unnecessary

Examples:

- `phase1-skeleton: feat(cli): add convert command skeleton`
- `phase1-extraction: test(excel): cover merged cell block detection`
- `phase1-skill: docs(skill): add minimal Copilot launcher guidance`

Use clear subjects. Avoid vague messages such as `update`, `fix issues`, or `work in progress`.

## Agent-Specific Instructions
- Use installed Codex skills proactively when they match the task.
- Keep `docs/phase1/spec.md` driven scope narrow unless the user explicitly changes the project direction.
- If `docs/phase1/spec.md` and an implementation request conflict, call out the conflict and ask or proceed only if the user's latest instruction clearly overrides it.
- Preserve Japanese product terminology where the phase documents already use it.
