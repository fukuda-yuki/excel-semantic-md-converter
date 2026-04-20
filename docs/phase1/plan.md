# Phase 1 Plan / Requirements Draft

この文書は Phase 1 の未確定な要件候補、計画案、詳細仕様化前に解消すべき疑問点を管理する。
この文書の内容は、`docs/phase1/spec.md` に反映されるまで実装仕様として扱わない。
プロダクト仕様・実装方針・優先順位の正は `docs/phase1/spec.md` とし、矛盾がある場合は `docs/phase1/spec.md` を優先する。

## 1. 目的

Excel を機械的に Markdown へ変換するのではなく、Excel の内容を LLM に再解釈させ、意味の通る Markdown を生成する。

Phase 1 では、主に `.xlsx` / `.xlsm` を入力として、セル本文、結合セル、表、見出し候補、テキスト入り図形、画像、グラフ、近接配置による補足関係を扱う MVP を作る。

## 2. 成果物

基本出力は次とする。

- `result.md`
- `manifest.json`
- `assets/` 配下の画像群

必要に応じて次を出力する。

- `logs/`
- `debug/`

## 3. 採用方針

- 本体は Python CLI とする。
- 自然言語 UX は GitHub Copilot skill で提供する。
- LLM 解釈は Python ツール内部の GitHub Copilot SDK local CLI が担当する。
- Excel の見た目画像化は、利用者の Windows 端末上で Excel COM を使って実行する。
- LLM 単位は `1 sheet = 1 Copilot SDK session` を基本とする。
- skill は薄い起動ラッパーに留める。
- prompt 構築、LLM 応答契約、変換ロジック、ビジネスロジックは Python ツール側に置く。
- 中間表現は Copilot SDK と skill から独立させる。

## 4. 非採用方針

- サーバー常駐はしない。
- Windows Service や非対話実行で Excel automation を行わない。
- skill 側に変換ロジックを寄せない。
- Excel workbook 全体を 1 回で LLM に投げない。
- 行単位で LLM に直接解釈させない。
- HTML 保存を主経路にしない。
- 全画像を無差別に LLM へ送らない。
- Markdown を Excel の完全な見た目再現へ寄せすぎない。
- Phase 1 では `.xls`、SmartArt、複雑な OLE 埋め込み、Excel 全機能網羅を初期対象にしない。

## 5. CLI 要件

安定した CLI を提供し、skill からも同じ CLI を呼び出す。

```bash
excel-semantic-md convert --input "C:\work\sample.xlsx" --out "C:\out"
excel-semantic-md inspect --input "C:\work\sample.xlsx"
excel-semantic-md render --input "C:\work\sample.xlsx" --sheet "要件一覧"
excel-semantic-md resume --job-id "20260421-001"
```

README に記載されている推奨オプション案は次。実装仕様として扱うには `docs/phase1/spec.md` への反映を必要とする。

- `--model`
- `--vision-model`
- `--max-images-per-sheet`
- `--resume`
- `--save-debug-json`
- `--save-render-artifacts`
- `--strict`

## 6. 想定アーキテクチャ

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

## 7. Phase 1 の処理フロー

1. CLI で入力ファイルと出力先を受け取る。
2. workbook / worksheet を読み取る。
3. OOXML drawing / image / chart を読み取る。
4. sheet ごとに block を検出する。
5. 視覚要素を近接 block に紐付ける。
6. レンダリング対象を決める。
7. Excel COM で range / shape / chart を画像化する。
8. sheet 単位の LLM 入力 JSON を作る。
9. Copilot SDK session を生成する。
10. 関連画像だけを attachment として付与する。
11. LLM 応答を構造化データへ戻す。
12. 全 sheet を `result.md` に統合する。
13. `manifest.json`、`assets/`、必要な `logs/` / `debug/` を保存する。

## 8. 中間表現の要件

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

README の例では `Rect`、`Block`、`ParagraphBlock`、`TableBlock`、`ShapeBlock`、`ImageBlock`、`ChartBlock`、`SheetModel` が示されている。実装仕様として扱うには `docs/phase1/spec.md` に反映し、JSON 直列化、ID 生成、manifest との対応を明確化する。

## 9. LLM 入力 / 応答要件

LLM 入力には最低限、次を含める。

- sheet 名
- block JSON
- asset path / attachment
- target format
- semantic Markdown として再構成する指示
- 不確実な箇所を断定しない指示

LLM 応答には最低限、次を含める。

- `sheet_summary`
- `sections`
- `figures`
- `unknowns`
- `markdown`

Markdown だけを直接返させず、検証可能な構造を返させる。

## 10. 受け入れ条件

Phase 1 は最低限、次を満たす。

1. 表だけの sheet を意味の通る Markdown table にできる。
2. テキスト入り図形を注記として Markdown に反映できる。
3. グラフを画像付きセクションとして説明できる。
4. 画像を attachment として LLM に渡し、本文へ反映できる。
5. 1 workbook 複数 sheet をまとめて `result.md` にできる。
6. `manifest.json` で各 block と asset の対応を追える。
7. CLI だけで再現可能である。
8. skill から呼んでも、実行結果は CLI 実行と同じになる。

## 11. 実装前の技術的負債化リスク

以下は、未確認のまま実装すると後から負債になりやすい点。

### 11.1 Copilot SDK local CLI 境界

- local CLI の具体的な呼び出し形式、標準入出力、終了コード、エラー形式が未確定。
- 画像 attachment の渡し方、上限、モデル要件、失敗時の挙動が未確定。
- hooks と session persistence を Phase 1 でどこまで実装するかが未確定。
- SDK が preview のため、依存バージョン固定や互換層の扱いが未確定。

### 11.2 Excel COM レンダリング

- `Range.CopyPicture` / `Shape.CopyPicture` / `Chart.Export` の使い分けとフォールバック方針が未確定。
- `CopyPicture` が clipboard 依存になる場合の安定化策が未確定。
- Excel プロセス終了時に、既存のユーザー Excel プロセスを巻き込まない設計が必要。
- `.xlsm` をマクロ無効・読み取り専用で開くための COM 設定と検証方法が未確定。
- hidden row / hidden column / filter / print area / zoom / freeze panes を画像化に反映する範囲が未確定。

### 11.3 OOXML / openpyxl 境界

- openpyxl で扱える範囲と raw OOXML で補う範囲の切り分けが未確定。
- chart、image、shape の anchor 取得精度と座標系の正規化方針が未確定。
- group shape、SmartArt、OLE など対象外要素をどう manifest に残すかが未確定。
- workbook 内画像の抽出ファイル名、重複、参照関係の扱いが未確定。

### 11.4 Block 検出

- 表、見出し、段落、注記、空白領域、複数表の判定ルールが未確定。
- 結合セルを見出しとして扱う条件が未確定。
- 数式セルについて、表示値、数式、cached value のどれを優先するかが未確定。
- cell comment / note / hyperlink を Phase 1 で扱うかが未確定。
- 大きすぎる sheet の section 分割条件が未確定。

### 11.5 LLM 品質と安全性

- Excel 内テキストを prompt instruction として扱わないための prompt injection 対策が未確定。
- LLM 応答 JSON の schema validation、再試行、部分失敗時の扱いが未確定。
- `unknowns` の粒度と最終 Markdown への出し方が未確定。
- 画像枚数、token budget、sheet 分割、再試行回数の上限が未確定。

### 11.6 出力と再開

- `manifest.json` の schema version、必須項目、エラー記録形式が未確定。
- `assets/` のディレクトリ構成と命名規則が未確定。
- `debug/` / `logs/` にユーザーデータが含まれるため、保存条件と redaction 方針が未確定。
- `resume` の job id、状態保存場所、冪等性、部分再実行範囲が未確定。

### 11.7 skill / 配布

- skill の配置場所、launcher script の呼び出し方法、Python 環境の解決方法が未確定。
- `allowed-tools` をどこまで許可するかが未確定。
- skill からの自然言語指示を CLI オプションへどう安全に変換するかが未確定。

### 11.8 テスト

- synthetic workbook fixture の作成方法が未確定。
- `.xlsm` の macro-disabled 契約を、危険な macro content を保存せずにどう検証するかが未確定。
- Excel COM が必要な live test と、通常の automated test の境界が未確定。
- Copilot SDK live access が必要な test と、adapter mock test の境界が未確定。

## 12. ユーザーへのヒアリング項目

詳細仕様化前に、以下を確認する。

### 12.1 MVP の優先順位

1. Phase 1 で最優先する workbook 型はどれか。
   - 表中心
   - 表 + 注記図形
   - 表 + グラフ
   - 表 + 画像
   - 複数 sheet の統合
2. Phase 1 の成功判定は「手元の代表 workbook で有用な Markdown が出ること」か、「汎用 fixture で自動テストが通ること」か、どちらを強く置くか。
3. Phase 1 で必ず扱いたい Excel 機能は README 記載分以外にあるか。

### 12.2 Excel の読み取り方

4. 数式セルは表示値を優先する方針でよいか。数式自体も manifest / debug に残す必要があるか。
5. hidden sheet / hidden row / hidden column は無視するか、見えていない情報として manifest に残すか。
6. cell comment / note / hyperlink は Phase 1 対象に含めるか。
7. filter 適用済み workbook は、見えている行だけを扱うか、元データ全体を扱うか。

### 12.3 画像化と asset

8. Range 画像は「LLM 補助用」であり、Markdown へ必ず貼る画像ではない、という扱いでよいか。
9. グラフ画像は Markdown に原則貼る方針でよいか。
10. 図形や画像は、近接 block に紐付く場合だけ Markdown に貼るか、孤立要素も独立セクションとして出すか。
11. asset ファイル名は安定性重視か、可読性重視か。

### 12.4 LLM / Copilot SDK

12. default model / vision model は CLI オプション未指定時にどうするか。README ではオプション名のみで既定値は未定義。
13. Copilot SDK hooks は Phase 1 から必須にするか、ログだけの最小実装でよいか。
14. session persistence / resume は Phase 1 で実動させるか、CLI surface と状態設計だけ先に置くか。
15. LLM 応答が壊れた場合、再試行するか、sheet を failed として継続するか。

### 12.5 出力ポリシー

16. `result.md` は sheet 順を保持して連結する方針でよいか。
17. 不確実な解釈は Markdown 本文に注記として出すか、`manifest.json` / `debug/` のみに出すか。
18. `manifest.json` は人間が読むことを重視するか、機械処理しやすい schema を重視するか。
19. `logs/` / `debug/` は既定では出さず、オプション指定時のみ出す方針でよいか。

### 12.6 セキュリティ / 配布

20. Excel 内のテキストは prompt injection の可能性があるため、LLM への指示ではなくデータとして扱う方針でよいか。
21. skill の `allowed-tools` は最小権限にし、広い事前承認は避ける方針でよいか。
22. 代表 workbook や生成 asset は、sanitized fixture 以外は repo に入れない方針でよいか。

## 13. 次の文書化順序

1. この `plan.md` のヒアリング項目を確認し、Phase 1 の判断を固める。
2. 固まった内容を `docs/phase1/spec.md` に詳細仕様として記載する。
3. 詳細仕様を `docs/phase1/task.md` に実装タスクとして分解する。
4. `docs/phase1/knowledge.md` は、README や仕様に含めない補助知識、外部制約、検証メモの置き場として整理する。
