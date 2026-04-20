# Phase 1 Plan / Requirements

この文書は Phase 1 の要件定義と、詳細仕様化前に確認した判断を管理する。
プロダクト仕様・実装方針・優先順位の正は `README.md` とし、矛盾がある場合は原則 `README.md` を優先する。

## 1. 目的

Excel を機械的に Markdown へ変換するのではなく、Excel の内容を LLM に再解釈させ、意味の通る Markdown を生成する。

Phase 1 では、主に `.xlsx` / `.xlsm` を入力として、セル本文、結合セル、表、見出し候補、テキスト入り図形、画像、グラフ、近接配置による補足関係を扱う MVP を作る。

## 2. Phase 1 の優先対象

ユーザー回答により、Phase 1 では次をすべて優先対象とする。

- 表中心 workbook
- 表 + 注記図形 workbook
- 表 + 画像 workbook
- 複数 sheet workbook

Phase 1 MVP と受け入れ条件にはグラフも含まれるため、グラフも Phase 1 対象から外さない。
ただし、最初の代表 workbook 確認では上記 4 種を優先する。

## 3. 成功判定

Phase 1 の成功判定は、次の両方を満たすこととする。

- 代表 workbook で有用な Markdown が出ること。
- 汎用 fixture の自動テストが通ること。

LLM と Excel COM に依存する動作は、通常の自動テストとは分離し、live confirmation として扱う。

## 4. 採用方針

- 本体は Python CLI とする。
- 自然言語 UX は GitHub Copilot skill で提供する。
- LLM 解釈は Python ツール内部の GitHub Copilot SDK local CLI が担当する。
- Excel の見た目画像化は、利用者の Windows 端末上で Excel COM を使って実行する。
- LLM 単位は `1 sheet = 1 Copilot SDK session` を基本とする。
- skill は薄い起動ラッパーに留める。
- prompt 構築、LLM 応答契約、変換ロジック、ビジネスロジックは Python ツール側に置く。
- 中間表現は Copilot SDK と skill から独立させる。

## 5. 非採用方針

- サーバー常駐はしない。
- Windows Service や非対話実行で Excel automation を行わない。
- skill 側に変換ロジックを寄せない。
- Excel workbook 全体を 1 回で LLM に投げない。
- 行単位で LLM に直接解釈させない。
- HTML 保存を主経路にしない。
- 全画像を無差別に LLM へ送らない。
- Markdown を Excel の完全な見た目再現へ寄せすぎない。
- Phase 1 では `.xls`、SmartArt、複雑な OLE 埋め込み、Excel 全機能網羅を初期対象にしない。
- Phase 1 では cell comment / note / hyperlink を対象にしない。
- Phase 1 では `resume` / session persistence を実装対象にしない。

## 6. 入力の扱い

- 対象入力は `.xlsx` / `.xlsm` とする。
- `.xlsm` は読み取り専用・マクロ無効前提で扱う。
- 数式セルは表示値を優先する。
- 数式文字列そのものは Phase 1 の通常出力、`manifest.json`、LLM 入力には含めない。
- hidden sheet / hidden row / hidden column / filter で非表示の行は、見えているものだけを扱う。
- cell comment / note / hyperlink は Phase 1 では読まない。

## 7. 画像化と Markdown への反映方針

Range 画像とは、Excel のセル範囲を Excel COM で画像化したものを指す。
Phase 1 では、Markdown で自然に再現できる表や本文は Markdown として出力し、Markdown で再現しにくいものは画像として貼る。

原則は次とする。

- 表は Markdown table として出す。
- セル範囲全体のスクリーンショットは、LLM 補助や debug 用を主用途とし、通常は `result.md` に貼らない。
- グラフは画像として `assets/` に保存し、原則 `result.md` に貼る。
- 元から画像である要素は、意味解釈に必要なものを `result.md` に貼る。
- テキスト入り図形は、テキストを Markdown 本文や注記に反映し、必要に応じて画像も貼る。
- 近接 block に紐付かない孤立図形・孤立画像・孤立グラフは、情報落ちを避けるため独立セクションとして扱う。

## 8. LLM / Copilot SDK 方針

- `--model` / `--vision-model` 未指定時は、Copilot CLI 側の既定に任せる。
- 関連する近傍画像だけを attachment として送る。
- LLM 応答が壊れた場合は一度だけ自動再試行する。
- 再試行しても失敗する場合は、該当 sheet を failed として記録し、他の sheet の処理を継続する。
- 不確実な解釈は `result.md` に注記として出す。
- `logs/` / `debug/` は既定では出さず、オプション指定時のみ出す。

## 9. CLI 要件

Phase 1 で安定させる CLI は次とする。

```bash
excel-semantic-md convert --input "C:\work\sample.xlsx" --out "C:\out"
excel-semantic-md inspect --input "C:\work\sample.xlsx"
excel-semantic-md render --input "C:\work\sample.xlsx" --sheet "要件一覧"
```

Phase 1 で扱うオプションは次。

- `--model`
- `--vision-model`
- `--max-images-per-sheet`
- `--save-debug-json`
- `--save-render-artifacts`
- `--strict`

## 10. 想定アーキテクチャ

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

推奨ディレクトリは次。

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

## 11. Phase 1 の処理フロー

1. CLI で入力ファイルと出力先を受け取る。
2. workbook / worksheet を読み取る。
3. OOXML drawing / image / chart を読み取る。
4. sheet ごとに block を検出する。
5. 視覚要素を近接 block に紐付ける。
6. 紐付かない視覚要素は独立 block として扱う。
7. レンダリング対象を決める。
8. Excel COM で shape / image / chart / 必要な range を画像化する。
9. sheet 単位の LLM 入力 JSON を作る。
10. Copilot SDK session を生成する。
11. 関連画像だけを attachment として付与する。
12. LLM 応答を構造化データへ戻す。
13. 全 sheet を `result.md` に統合する。
14. `manifest.json`、`assets/`、必要な `logs/` / `debug/` を保存する。

## 12. 中間表現の要件

中間表現は LLM 実行基盤や skill に依存させない。

Phase 1 で必要な概念は次。

- workbook
- sheet
- range / rect
- block
- paragraph
- heading
- table
- shape
- image
- chart
- note
- asset reference
- unknown / warning
- failed sheet

中間表現の基本形として、`Rect`、`Block`、`ParagraphBlock`、`TableBlock`、`ShapeBlock`、`ImageBlock`、`ChartBlock`、`SheetModel` を想定する。
詳細仕様では、JSON 直列化、ID 生成、manifest との対応を明確化する。

## 13. LLM 入力 / 応答要件

LLM 入力には最低限、次を含める。

- sheet 名
- block JSON
- asset path / attachment
- target format
- semantic Markdown として再構成する指示
- 不確実な箇所を断定しない指示
- Excel 内テキストを指示ではなくデータとして扱う指示

LLM 応答には最低限、次を含める。

- `sheet_summary`
- `sections`
- `figures`
- `unknowns`
- `markdown`

Markdown だけを直接返させず、検証可能な構造を返させる。

## 14. 受け入れ条件

Phase 1 は最低限、次を満たす。

1. 表だけの sheet を意味の通る Markdown table にできる。
2. テキスト入り図形を注記として Markdown に反映できる。
3. グラフを画像付きセクションとして説明できる。
4. 画像を attachment として LLM に渡し、本文へ反映できる。
5. 1 workbook 複数 sheet をまとめて `result.md` にできる。
6. `manifest.json` で各 block と asset の対応を追える。
7. CLI だけで再現可能である。
8. skill から呼んでも、実行結果は CLI 実行と同じになる。
9. LLM 応答失敗時に 1 回だけ再試行し、それでも失敗した sheet は failed として残せる。
10. `logs/` / `debug/` はオプション指定時のみ出力される。

## 15. 解消済みの判断

- Phase 1 の最優先 workbook 型は、表中心、表 + 注記図形、表 + 画像、複数 sheet 統合。
- 成功判定は代表 workbook と汎用 fixture の両方。
- 数式セルは表示値優先。
- hidden / filtered out の情報は、見えているものだけを扱う。
- cell comment / note / hyperlink は Phase 1 対象外。
- Markdown で再現できないものは画像として貼る。
- 孤立した視覚要素は独立セクションとして扱う。
- model 未指定時は Copilot CLI 側の既定に任せる。
- `resume` / session persistence は Phase 1 では不要。
- LLM 応答破損時は 1 回だけ再試行し、失敗 sheet として継続する。
- 不確実な解釈は `result.md` に注記する。
- `logs/` / `debug/` は opt-in。

## 16. 残課題

- Copilot SDK local CLI の具体的な呼び出し形式、画像 attachment 仕様、hooks の実装詳細は実装時に確認する必要がある。
- Excel COM の `Range.CopyPicture` / `Shape.CopyPicture` / `Chart.Export` の安定性は live confirmation が必要である。
- asset 命名規則、manifest schema、block 検出 heuristic は `spec.md` で詳細化する。

## 17. 次の文書化順序

1. `docs/phase1/spec.md` に詳細仕様を記載する。
2. `docs/phase1/task.md` に実装タスクを分解する。
3. `docs/phase1/knowledge.md` を補助知識、外部制約、検証メモの置き場として整理する。
