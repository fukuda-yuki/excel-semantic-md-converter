# Phase 1 Knowledge

エージェント間で前提知識、判断理由、確認済み事項がぶれないようにするための共有メモ。

## 2026-04-21: agent operation rules confirmed by user

- 疑問がある場合は立ち止まってユーザーへ確認する。推測で進めない。
- ユーザーへの迎合は不要。技術的・仕様的な矛盾やリスクは明確に指摘する。
- Source of Truth は絶対に `docs/phase1/spec.md`。
- `docs/phase1/plan.md`、`docs/phase1/task.md`、`docs/phase1/knowledge.md` は共通運用文書として扱うが、仕様判断では `docs/phase1/spec.md` を優先する。
- 実装完了前に複数サブエージェントを並列で立ち上げてレビューさせる。
- MainAgent はサブエージェント指摘の取りまとめと妥当性検証に徹する。
- レビュー結果は、指摘ゼロ、棄却、起動失敗を含めて `docs/review/phase1/<milestone>.md` へ記録する。妥当な指摘へ対応する場合は、対応方針を書いてから修正する。
- 実装完了宣言時には `docs/phase1/task.md` を更新する。
- レビュー指摘に対応した場合は `docs/review/phase1/<milestone>.md` も更新する。
- エージェント間で共有すべき事項は `docs/phase1/knowledge.md` に記録する。
- 仕様として確定した事項は `docs/phase1/spec.md` に反映する。`knowledge.md` は仕様の代替ではなく、背景・判断理由・確認履歴の共有先である。
