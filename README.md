# Excel Semantic Markdown Converter

Excel workbook の内容を LLM に再解釈させ、意味の通る Markdown を生成するローカル CLI ツール。

単なる Excel to Markdown 変換ではなく、表、図形、画像、グラフ、近接配置による補足関係を読み取り、文書として使いやすい Markdown に再構成する。

## Source of Truth

Phase 1 のプロダクト仕様、実装方針、優先順位の正は `docs/phase1/spec.md` である。
この README はプロジェクト背景、全体像、初期構想の参考資料として扱う。

Phase 1 の詳細は次に分ける。

- `docs/phase1/plan.md`: Phase 1 の要件と判断
- `docs/phase1/spec.md`: Phase 1 の詳細仕様
- `docs/phase1/task.md`: 実装タスク
- `docs/phase1/knowledge.md`: 補助知識、外部制約、検証メモ

README と `docs/phase1/spec.md` が矛盾する場合は `docs/phase1/spec.md` を優先する。

## 方針

- 本体は Python CLI とする。
- 自然言語 UX は GitHub Copilot skill で提供する。
- LLM 解釈は Python ツール内部の GitHub Copilot SDK local CLI が担当する。
- Excel の見た目取得は、利用者の Windows 端末上で Excel COM を使って画像化する。
- LLM 単位は `1 sheet = 1 Copilot SDK session` を基本とする。
- Excel 構造抽出と block 化を主情報とし、画像をもとにした LLM 分析は補足情報として扱う。
- skill は薄い起動ラッパーに留める。
- prompt 構築、LLM 応答契約、変換ロジック、ビジネスロジックは Python ツール側に置く。
- 中間表現は Copilot SDK と skill から独立させる。

## 対象

入力:

- `.xlsx`
- `.xlsm`

`.xlsm` は読み取り専用・マクロ無効前提で扱う。

出力:

- `result.md`
- `manifest.json`
- `assets/`

必要に応じて、オプション指定時のみ `debug/` や `logs/` を出力する。

## Phase 1 スコープ

Phase 1 で扱う workbook 型:

- 表中心 workbook
- 表 + 注記図形 workbook
- 表 + 画像 workbook
- 表 + グラフ workbook
- 複数 sheet workbook

Phase 1 で扱う要素:

- セル本文
- 結合セル
- 表
- 見出し候補
- 段落候補
- テキスト入り図形
- 画像
- グラフ
- 近接配置による補足関係

Phase 1 で扱わない要素:

- `.xls`
- SmartArt の完全解釈
- 複雑な OLE 埋め込みの完全解釈
- cell comment
- note
- hyperlink
- `resume` / session persistence
- Excel の完全な見た目再現
- server-side / service 実行

## 読み取り方針

- 数式セルは表示値を優先する。
- hidden sheet / hidden row / hidden column / filter で非表示の行は扱わない。
- Markdown で自然に再現できるものは Markdown として出す。
- Markdown で再現しにくい画像、グラフ、意味を持つ図形は画像として貼る。
- 画像からの LLM 分析は、セル値、OOXML メタデータ、block 検出結果を補うために使う。
- 紐付かない図形、画像、グラフは破棄せず独立セクションとして扱う。
- 不確実な解釈は `result.md` に注記として出す。

## CLI

Phase 1 で安定させる CLI:

```bash
excel-semantic-md setup
excel-semantic-md convert --input "C:\work\sample.xlsx" --out "C:\out"
excel-semantic-md inspect --input "C:\work\sample.xlsx"
excel-semantic-md render --input "C:\work\sample.xlsx" --sheet "要件一覧"
```

`setup` はローカル実行に必要な前提を確認し、不足している設定や導線を案内する。
自動的な外部インストール、認証情報の保存、ユーザー workbook の変更は行わない。

Phase 1 の主なオプション:

- `--model`
- `--vision-model`
- `--max-images-per-sheet`
- `--save-debug-json`
- `--save-render-artifacts`
- `--strict`

`--model` / `--vision-model` 未指定時は Copilot CLI 側の既定に任せる。

## アーキテクチャ

```text
User
 └─ GitHub Copilot
     └─ skill
         └─ launcher script
             └─ Python CLI
                 ├─ excel: workbook / OOXML / block extraction
                 ├─ render: Excel COM rendering
                 ├─ llm: Copilot SDK local CLI integration
                 └─ output: Markdown / manifest / assets
```

推奨ディレクトリ:

```text
src/excel_semantic_md/
  app/
  excel/
  render/
  llm/
  output/
  cli/

skills/excel-semantic-markdown/
  SKILL.md
  run_excel_semantic_md.ps1
  examples.md
```

## LLM 方針

- workbook 全体を 1 prompt にしない。
- row 単位で LLM に投げない。
- sheet 単位で LLM に解釈させる。
- 画像 attachment は関連する近傍画像だけに限定する。
- 画像をもとにした LLM 分析は、構造抽出結果を補完するための補足情報として扱う。
- Excel 内テキストは prompt instruction ではなく data として扱う。
- LLM 応答が壊れた場合は 1 回だけ再試行する。
- 再試行しても失敗する場合は該当 sheet を failed として記録し、他 sheet の処理を継続する。

## 受け入れ条件

- 表だけの sheet を意味の通る Markdown table にできる。
- テキスト入り図形を注記として Markdown に反映できる。
- グラフを画像付きセクションとして説明できる。
- 画像を attachment として LLM に渡し、補足情報として本文へ反映できる。
- 複数 sheet をまとめて `result.md` にできる。
- `manifest.json` で block と asset の対応を追える。
- CLI だけで再現可能である。
- skill から呼んでも、実行結果は CLI 実行と同じになる。
- `setup` でローカル前提条件を確認できる。

## 開発

`pyproject.toml` が追加されたら、リポジトリルートで次を使う。

```bash
python -m pip install -e .
```

テストコマンドが追加されるまでは、Python テスト追加後に次を使う。

```bash
python -m pytest
```

Excel COM、Copilot CLI、skill 実行に依存する確認は live confirmation として扱い、自動テストだけで end-to-end 成功を主張しない。
