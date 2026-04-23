# phase1-validation-review review note

## Review Scope

- Milestone: `phase1-validation-review`
- Scope: `docs/phase1/task.md` の 11〜15、validation 記録、`convert` の sheet-level failure 正規化、`manifest.json` の LLM metadata、関連テストとドキュメント更新
- Out of scope: `pywin32` 導入、Copilot 認証変更、Copilot SDK local CLI behavior / vision attachment behavior の live confirmation 完了
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `docs/phase1/knowledge.md`
- `docs/phase1/task.md`
- `docs/review/phase1/phase1-validation-review.md`
- `src/excel_semantic_md/app/convert_pipeline.py`
- `src/excel_semantic_md/llm/adapter.py`
- `src/excel_semantic_md/llm/models.py`
- `src/excel_semantic_md/output/writers.py`
- `tests/test_llm.py`
- `tests/test_output.py`

## Subagents

- pre-fix spec compliance and functional correctness reviewer: completed
  - Agent: `Singer` (`019db756-dd8b-7d73-a888-4925866ad8c2`)
  - Launch status: success
  - Result: 4 findings reported
- pre-fix tests, edge cases, and regression risk reviewer: completed
  - Agent: `Maxwell` (`019db756-ddd8-7180-8738-fbdc3c0a89b7`)
  - Launch status: success
  - Result: 4 findings reported
- post-fix spec compliance and functional correctness reviewer: completed
  - Agent: `Euler` (`019db761-fc73-75a0-aa7b-3f6130eafc4e`)
  - Launch status: success
  - Result: 3 findings reported
- post-fix tests, edge cases, and regression risk reviewer: completed
  - Agent: `Meitner` (`019db761-fcc3-7b32-bdfb-789b81cc0bce`)
  - Launch status: success
  - Result: 2 findings reported
- post-fix reliability and cleanup reviewer: completed
  - Agent: `Linnaeus` (`019db761-fce7-7740-bf0d-ab97a032ad08`)
  - Launch status: success
  - Result: 3 findings reported

## Raw Findings Summary

- `Singer`
  - P1: task 14 の `setup` / Copilot SDK local CLI / vision attachment / `--strict` validation 証跡が不足している
  - P1: `phase1-validation-review` の review note が未作成
  - P2: 一部の未捕捉例外では `manifest.json` へ中間状態を残せない
  - P2: spec 9.2 の `used model` が `manifest.json` に出ていない
- `Maxwell`
  - P1: `setup` の検証はモック中心で、実環境の診断分岐の確認が不足している
  - P1: vision attachment は優先度付けロジックしか確認されておらず、実 SDK / local CLI 境界の live confirmation が未実施
  - P2: `--strict` はモック化テスト中心で、実 CLI 証跡が不足している
  - P3: 既存 review note の形式不備ではなく、未検証項目が残っている点が問題
- `Euler`
  - P1: render / LLM 例外経路で `manifest.json` の stage status が `skipped` になり得る
  - P2: 既に正しい `assets/...` path を含む Markdown が basename 置換で二重化し得る
  - P3: `knowledge.md` の pytest 件数が stale
- `Meitner`
  - P2: task 14 の `--save-debug-json` / `--save-render-artifacts` validation 完了判定が CLI/integration 実測ではない
  - P3: render plan 例外の修正が output writer 契約まで検証されていない
- `Linnaeus`
  - P1: fatal abort 経路では partial result cleanup warning が観測不能
  - P2: review note の post-fix reviewer 状態が pending のまま
  - P3: `knowledge.md` の pytest 件数が stale

## MainAgent Validity Judgment

- Accepted:
  - `phase1-validation-review` review note 未作成
  - `setup` 実行確認不足
  - `--strict` 実機証跡不足
  - Copilot SDK local CLI behavior / vision attachment behavior が task 14 で未完了のまま残っていることの明示不足
  - `manifest.json` の `used model` 未出力
  - 一部の未捕捉例外で sheet failure に正規化されず、可能な範囲の中間出力まで到達できない経路
  - render / LLM 例外経路の stage status が `skipped` になり得る問題
  - asset path の basename 二重置換リスク
  - render plan 例外が output writer まで落ちることのテスト不足
  - post-fix reviewer 状態と pytest 件数の stale 記録
- Rejected:
  - `phase1-output-generation` review note の形式不備
    - Reason: 既存 note は required sections を満たしており、不備は validation scope が未完了のまま残っている点にある
  - `--save-debug-json` / `--save-render-artifacts` validation 完了判定を未完了へ戻す提案
    - Reason: これらは既存 output milestone で writer と option contract を検証済みで、今回の milestone は未実施だった `setup` / `--strict` 実測と stale task 整理を対象にしている。追加の live confirmation blocker ではない
  - fatal abort 経路の cleanup warning を今回の blocker とする提案
    - Reason: 今回の修正で sheet pipeline 内例外は可能な範囲で sheet failure に正規化される。workbook-level fatal abort の cleanup warning 観測性は残リスクとして記録するが、11〜15 の完了 blocker にはしない

## Response Plan

- `convert` の sheet pipeline 内で未捕捉例外を `FailureInfo` に正規化し、可能な限り他 sheet 継続と output writer 到達を維持する
- Copilot SDK session から取得できる場合だけ `used_model` を `LlmRunResult` と `manifest.json` に保持する
- `docs/phase1/task.md` と `docs/phase1/knowledge.md` を 2026-04-23 の validation 実測へ更新する
- 修正後に post-fix subagent review を 3 本並列実行し、妥当な指摘だけを反映する

## Applied Fixes

- `LlmRunResult` に `used_model` を追加し、Copilot SDK session の current model が取得できる場合だけ保持するようにした。
- `manifest.json` の sheet-level `llm` payload に `used_model` を含めるようにした。
- `convert` の sheet pipeline 内で render plan を含む未捕捉例外を `FailureInfo` に正規化し、可能な限り他 sheet 継続と output writer 到達を維持するようにした。
- render plan 例外正規化と `used_model` 出力の回帰テストを追加した。
- render / LLM 例外経路でも `manifest.json` の stage status が `failed` になるようにした。
- 既に公開 asset path を含む Markdown を二重置換しないようにした。
- render plan 例外が `result.md` / `manifest.json` に反映されることをテストに追加した。
- `docs/phase1/task.md` で stale な未完了項目を整理し、`setup` 実行確認、`--strict` 実行確認、manifest への visual linking 反映を完了へ更新した。
- `docs/phase1/knowledge.md` に 2026-04-23 の validation 実測と現環境制約を追記した。

## Validation

- `python -m pytest tests/test_llm.py tests/test_output.py -q`
  - Result: passed (`19 passed`)
- `python -m pytest -q`
  - Result: passed (`86 passed`)
- `python -m excel_semantic_md.cli.main setup --out .tmp-validation-setup`
  - Result: exit code `0`
- `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-validation-convert`
  - Result: exit code `0`, `result.md` / `manifest.json` generated, render failure was recorded because `pywin32` is unavailable in this environment
- `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-validation-strict --strict`
  - Result: exit code `1`, `result.md` / `manifest.json` generated, failed sheet was retained in outputs
- `C:\Users\mwam0\AppData\Roaming\Python\Python314\Scripts\excel-semantic-md.exe setup --out .tmp-validation-script-setup`
  - Result: exit code `0`
- `C:\Users\mwam0\AppData\Roaming\Python\Python314\Scripts\excel-semantic-md.exe convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-validation-script-convert`
  - Result: exit code `0`, `result.md` / `manifest.json` generated
- `C:\Users\mwam0\AppData\Roaming\Python\Python314\Scripts\excel-semantic-md.exe convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-validation-script-strict --strict`
  - Result: exit code `1`, `result.md` / `manifest.json` generated

## Residual Risks

- `pywin32` 未導入のため、現環境では Excel COM rendering は live confirmation 完了扱いにできない
- Copilot SDK local CLI behavior と vision attachment behavior は認証状態と実モデル応答に依存し、この turn では pending のまま残す
- `excel-semantic-md.exe` は user scripts 配下に生成されるが、この shell の `PATH` には含まれていないため、素の `excel-semantic-md` コマンドは環境調整なしでは解決されない
- workbook-level fatal abort 経路では partial cleanup warnings を user-facing output に反映できない場合がある

## Pending Items

- Copilot SDK local CLI behavior live confirmation
- vision attachment behavior live confirmation
