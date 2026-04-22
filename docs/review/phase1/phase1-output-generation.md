# phase1-output-generation review note

## Review Scope

- Milestone: `phase1-output-generation`
- Scope: `convert` orchestration, output aggregation DTO, `result.md` / `manifest.json` / `assets/` / `debug/` writers, related tests, phase docs updates
- Out of scope: Excel COM live confirmation quality, Copilot SDK live confirmation, `setup` command validation
- Source of Truth: `docs/phase1/spec.md`

## Changed Files

- `docs/phase1/knowledge.md`
- `docs/phase1/task.md`
- `docs/review/phase1/phase1-output-generation.md`
- `src/excel_semantic_md/app/__init__.py`
- `src/excel_semantic_md/app/convert_pipeline.py`
- `src/excel_semantic_md/cli/main.py`
- `src/excel_semantic_md/excel/block_detector.py`
- `src/excel_semantic_md/output/__init__.py`
- `src/excel_semantic_md/output/models.py`
- `src/excel_semantic_md/output/writers.py`
- `tests/test_cli.py`
- `tests/test_output.py`

## Subagents

- spec compliance and functional correctness reviewer: completed
  - Agent: `Hegel` (`019db743-f895-7b21-ba9f-e8de32801bae`)
  - Launch status: success
  - Result: 4 findings reported.
- tests, edge cases, and regression risk reviewer: completed
  - Agent: `Pasteur` (`019db744-0c99-7441-9295-c134181a957c`)
  - Launch status: success
  - Result: 2 findings reported.
- architecture and separation reviewer: completed
  - Agent: `Ohm` (`019db744-20a7-7d73-9844-466f1988e5fd`)
  - Launch status: success
  - Result: 3 findings reported.
- reliability and cleanup reviewer: completed
  - Agent: `Dirac` (`019db744-34a4-7093-b19b-7dff62a22097`)
  - Launch status: success
  - Result: 5 findings reported.

## Raw Findings Summary

- `Pasteur`
  - P1: render / LLM 例外が sheet failure に正規化されず `convert` 全体を abort し得る。
  - P1: `LlmRunResult(status="failed")` でも `failure` が無いと success 扱いになる。
- `Hegel`
  - P1: LLM 本文が basename のまま画像参照すると、最終公開 path へ書き換わらず壊れた参照が残る。
  - P1: 同じ `--out` の再実行で `debug/` と旧 artifact が残留する。
  - P2: `block_detection.json` が debug 出力に無い。
  - P2: 通常 manifest に内部 temp path / workbook path が漏れる。
- `Dirac`
  - P1: sheet 途中の予期しない例外で temp dir cleanup に到達できない経路がある。
  - P1: manifest に temp path / local workbook path が流れる。
  - P1: 最終出力ディレクトリへ直書きなので部分的な成果物が残る。
  - P2: `assets/` / `debug/` を掃除しないため旧生成物が残る。
  - P2: temp dir cleanup 失敗を完全に握りつぶしている。
- `Ohm`
  - P1: app 層が output DTO に逆依存している。
  - P1: output DTO が runtime 状態と公開後状態を同時に所有している。
  - P2: debug snapshot と live object を二重管理している。

## MainAgent Validity Judgment

- Accepted and fixed:
  - render / LLM 例外の sheet-level failure 正規化
  - `LlmRunResult(status="failed")` without `failure` の failure 補完
  - basename asset 参照の最終公開 path への書き換え
  - `block_detection.json` の追加
  - manifest から temp path / local workbook path を隠すサニタイズ
  - staging directory を使った出力書き込みと managed outputs の入れ替え
  - `debug/` / `assets/` の再実行時残留防止
  - pipeline 途中例外時の partial temp dir cleanup
  - temp dir cleanup 失敗の warning 化
- Rejected:
  - app 層から output DTO への依存そのものをこの milestone の blocker とみなす指摘
    - Reason: 今回のユーザー承認済み計画が「`output/` に内部 DTO と writer 群を追加する」前提であり、責務境界の改善余地はあるが、現時点では correctness を崩す欠陥より設計負債に近い。
  - runtime state と published state の DTO 分離を今すぐ必須とする指摘
    - Reason: 現構造でも manifest/result/debug の契約は満たせており、再設計コストに対して今回の milestone 完了条件を越える。
  - debug snapshot と live object の単一 canonical model 化を必須とする指摘
    - Reason: debug 出力は opt-in の観測用であり、今回修正後は `block_detection` / post-link / render plan / LLM 入出力の各段階を個別に固定できる。直近の correctness blocker ではない。

## Response Plan

- Completed: `convert` の render / LLM 例外を sheet-level `FailureInfo` に落とし、strict/non-strict 契約を維持する。
- Completed: failed-status without `failure` を generic `FailureInfo` へ正規化する。
- Completed: `result.md` で basename の画像参照を公開 asset path へ差し替える。
- Completed: `debug/` に `block_detection.json` を追加する。
- Completed: manifest の warning / failure details をサニタイズして temp path / local workbook path を出さない。
- Completed: staging directory 書き込みで managed outputs を置き換え、旧 `assets/` / `debug/` を残さないようにする。
- Completed: partial pipeline failure 時も既に作られた render temp dir を cleanup する。
- Completed: cleanup failure を warning として観測できるようにする。

## Applied Fixes

- `detect_blocks()` が workbook reading の warnings / failures を `SheetModel` に引き継ぐよう更新した。
- `convert_pipeline` を追加し、read -> detect -> visual -> link -> render -> llm -> output の最小オーケストレーションを実装した。
- `ConvertResult` / `ConvertSheetResult` / `PublishedAsset` を追加し、writer が必要な集約状態を保持できるようにした。
- `write_convert_outputs()` を追加し、`result.md` / `manifest.json` / `assets/` / `debug/` を生成するようにした。
- basename 参照から公開 asset path への置換、manifest details のサニタイズ、staging replace、managed outputs cleanup を追加した。
- `block_detection.json` を debug 出力に追加した。
- render temp dir cleanup を `convert` 完了時と pipeline 途中例外時の両方で試みるようにした。
- `tests/test_output.py` を追加し、writer 契約、strict/non-strict、render exception、LLM failed-status without details を固定した。
- `tests/test_cli.py` を更新し、`convert` が pipeline を呼ぶことを確認するようにした。
- `docs/phase1/task.md` と `docs/phase1/knowledge.md` を今回の実装状態へ更新した。

## Validation

- `python -m pytest -q`
  - Result: passed (`83 passed`)
- `python -m pip install -e .`
  - Result: passed
- `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out <temp dir>`
  - Result: exit code `0`, `result.md` / `manifest.json` generated

## Residual Risks

- Excel COM / Copilot SDK の live confirmation は未実施で、実環境での render quality や vision attachment の有効性は残課題。
- staging replace は managed outputs だけを置換するため、同じ `--out` 配下の unmanaged files は意図的に残す。
- app/output 境界は今回の計画に沿って実装したが、別出力先や API 応答へ再利用する段階では DTO 再整理の余地がある。

## Pending Items

- `excel-semantic-md setup` の validation は未実施。
- Copilot SDK local CLI behavior の live confirmation は未実施。
- vision attachment behavior の live confirmation は未実施。
