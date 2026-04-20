# Excel Semantic Markdown Converter

`Python tool + GitHub Copilot skill wrapper + Copilot SDK(local CLI)` 方式

## 1. 目的

Excel を機械的に Markdown 化するのではなく、**内容を LLM に再解釈させたうえで、意味の通る Markdown を生成する**ツールを作る。

対象は主に `.xlsx` / `.xlsm`。
ユーザーの Windows 端末上でローカル実行する。
サーバー常駐化はしない。
自然言語から使えるように、GitHub Copilot の **skill** でラップして配布する。
LLM 実行基盤は **GitHub Copilot SDK** を前提とする。Copilot SDK はローカルの Copilot CLI を使う構成があり、画像入力、hooks、session persistence、custom skills を備える。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/set-up-copilot-sdk/local-cli) ([GitHub Docs][1]))

---

## 2. 最終方針

### 2.1 採用方針

* **本体は Python CLI ツール**
* **自然言語 UX は Copilot skill**
* **LLM 解釈は Python ツール内部の Copilot SDK が担当**
* **Excel の見た目画像化は Excel COM を使ってローカル端末で実行**
* **1 sheet = 1 Copilot SDK session** を基本単位にする
* **skill は薄いラッパー** に留め、変換ロジックは Python 側に集約する

この理由は、Copilot SDK は実行基盤として有効だが、変換品質の本体は Excel 前処理と block 化にあるため。skills は再利用可能な prompt module だが、ロジック本体の置き場所には向かない。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk/custom-skills) ([GitHub Docs][2]))

### 2.2 非採用方針

* サーバー常駐
* Windows Service / 非対話実行での Excel automation
* skill 側に変換ロジックを寄せる設計
* Excel 全体を 1 回で LLM に投げる設計
* 行単位での LLM 解釈
* HTML 保存を主経路にする設計

Microsoft は Office の server-side automation を非推奨・非サポートとしている。今回はローカル対話セッション前提なので、この制約を避ける。 ([support.microsoft.com](https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2) ([Microsoft サポート][3]))

---

## 3. 解くべき問題の定義

このツールは「Excel to Markdown converter」ではなく、実質的には **Excel to Semantic Markdown pipeline** である。

### 入力

* `.xlsx`
* `.xlsm`
  ただし `.xlsm` はマクロ実行リスクがあるため、読み取り専用・マクロ無効で扱う前提を置く

### 出力

* `result.md`
* `manifest.json`
* `assets/` 配下の画像群
* 必要に応じて `logs/` や `debug/` の中間成果物

### 変換対象

* セル本文
* 結合セル
* 表
* 見出しっぽい領域
* テキストを持つ図形
* 画像
* グラフ
* 近接配置による補足関係

### 変換対象外

* 完全な見た目再現
* `.xls` 初期対応
* SmartArt や複雑な OLE 埋め込みの完全解釈
* Excel の全機能網羅

---

## 4. 実装アーキテクチャ

## 4.1 全体像

```text
User
 └─ GitHub Copilot (natural language)
     └─ Skill (thin wrapper)
         └─ PowerShell / shell launcher
             └─ Python CLI
                 ├─ Excel structure extraction
                 ├─ Excel rendering (COM)
                 ├─ Copilot SDK interpretation
                 └─ Markdown / manifest output
```

GitHub Copilot skill は `SKILL.md` と補助スクリプトで構成でき、必要なら `allowed-tools` を frontmatter に持てる。だが tool 事前許可は強めの権限になるため、既定では広く pre-approve しない方針とする。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/use-copilot-agents/cloud-agent/add-skills) ([GitHub Docs][4]))

---

## 4.2 推奨ディレクトリ構成

```text
excel-semantic-md/
  pyproject.toml
  README.md

  src/
    excel_semantic_md/
      app/
        orchestrator.py
        config.py
        models.py

      excel/
        workbook_reader.py
        ooxml_drawings.py
        block_detector.py
        link_resolver.py

      render/
        excel_session.py
        range_renderer.py
        shape_renderer.py
        chart_renderer.py
        render_plan.py

      llm/
        copilot_client.py
        prompt_builder.py
        attachment_builder.py
        response_parser.py
        hooks.py
        session_store.py

      output/
        markdown_writer.py
        manifest_writer.py
        asset_store.py

      cli/
        main.py

  skills/
    excel-semantic-markdown/
      SKILL.md
      run_excel_semantic_md.ps1
      examples.md
```

---

## 4.3 レイヤ責務

### A. `excel/`

責務: Excel ファイルの構造抽出

想定内容:

* workbook / worksheet 読み取り
* 使用範囲の推定
* 結合セル取得
* 表候補検出
* 図形 / 画像 / グラフの OOXML メタデータ取得
* block 分割
* block と視覚要素の紐付け

想定技術:

* `openpyxl`
* `.xlsx` を zip として直接読み、drawing / chart / image 参照を追う raw OOXML 処理

openpyxl には chart anchor の概念があり、画像やグラフがセル本文とは別の描画オブジェクトである前提に沿って設計できる。 ([openpyxl.readthedocs.io](https://openpyxl.readthedocs.io/en/3.1/charts/anchors.html?utm_source=chatgpt.com) ([GitHub Docs][5]))

### B. `render/`

責務: Excel 本体を使った見た目の画像化

想定内容:

* Excel.Application セッション管理
* Range の画像化
* Shape の画像化
* Chart の画像化

Excel の公式 API として、`Range.CopyPicture`、`Shape.CopyPicture`、`Chart.Export`、`ExportAsFixedFormat` 系がある。今回は主経路だけでよいので、最低限 `Range.CopyPicture` / `Shape.CopyPicture` / `Chart.Export` を採用対象とする。 ([learn.microsoft.com](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chart?view=excel-pia) ([Microsoft Learn][6]))

### C. `llm/`

責務: Copilot SDK を使った sheet 単位の意味解釈

想定内容:

* local CLI 経由の Copilot session 生成
* prompt 組み立て
* 画像 attachment 付与
* 応答パース
* hooks で監査 / 失敗記録
* 必要なら session persistence

Copilot SDK は local CLI で使え、画像 input、hooks、session persistence をサポートする。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/set-up-copilot-sdk/local-cli) ([GitHub Docs][1]))

### D. `output/`

責務: 最終成果物生成

想定内容:

* Markdown 組み立て
* manifest 出力
* asset 配置
* debug / log 出力

---

## 5. 中間表現

LLM や skill に依存しない **中間表現** を持つこと。
これが今回の設計の最重要点。

```python
from dataclasses import dataclass
from typing import Optional, Literal

@dataclass
class Rect:
    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

@dataclass
class Block:
    id: str
    kind: Literal["heading", "paragraph", "table", "shape", "image", "chart", "note"]
    anchor: Rect

@dataclass
class ParagraphBlock(Block):
    text: str

@dataclass
class TableBlock(Block):
    headers: list[str]
    rows: list[list[str]]
    caption: Optional[str] = None

@dataclass
class ShapeBlock(Block):
    shape_type: str
    text: Optional[str]
    asset_path: Optional[str]

@dataclass
class ImageBlock(Block):
    asset_path: str
    hint: Optional[str]

@dataclass
class ChartSeries:
    name: str
    categories: list[str]
    values: list[str]

@dataclass
class ChartBlock(Block):
    title: Optional[str]
    series: list[ChartSeries]
    asset_path: str

@dataclass
class SheetModel:
    name: str
    blocks: list[Block]
```

この層は Copilot SDK 非依存にする。
将来、実行基盤を差し替えても block 化ロジックを使い回せるようにするため。

---

## 6. 実行フロー

## 6.1 全体フロー

1. CLI で入力ファイルと出力先を受け取る
2. Workbook を読む
3. OOXML drawing / image / chart を読む
4. sheet ごとに block を検出する
5. 視覚要素を近接 block に紐付ける
6. レンダリング対象を決める
7. Excel COM で range / shape / chart を画像化する
8. sheet 単位の LLM 入力 JSON を作る
9. Copilot SDK session を生成する
10. 画像 attachment を付与して解釈させる
11. 応答を構造化データに戻す
12. 全 sheet を Markdown に統合する
13. manifest と logs を保存する

## 6.2 LLM 単位

* **1 workbook = 複数 sheet**
* **1 sheet = 1 session**
* block 数が極端に多い sheet のみ section 分割

理由:

* 1 workbook は重すぎる
* 1 row は文脈不足
* 1 sheet が最も自然

---

## 7. CLI 設計

skill から叩く先は **安定した CLI** にすること。
自然言語 UX があっても、本体は再現性のあるコマンドラインであるべき。

### 推奨コマンド

```bash
excel-semantic-md convert --input "C:\work\sample.xlsx" --out "C:\out"
excel-semantic-md inspect --input "C:\work\sample.xlsx"
excel-semantic-md render --input "C:\work\sample.xlsx" --sheet "要件一覧"
excel-semantic-md resume --job-id "20260421-001"
```

### 推奨オプション

* `--model`
* `--vision-model`
* `--max-images-per-sheet`
* `--resume`
* `--save-debug-json`
* `--save-render-artifacts`
* `--strict`

---

## 8. Copilot SDK 利用方針

## 8.1 接続方式

* **local CLI path** を採用
* 利用者端末に Copilot CLI が入っており、サインイン済みである前提
* Python ツール内部で Copilot SDK を使う

Copilot SDK の local CLI 利用は、ローカルサインイン済み環境をそのまま使えるので、今回の「各利用者へ配布するローカルツール」に適している。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/set-up-copilot-sdk/local-cli) ([GitHub Docs][1]))

## 8.2 画像入力

* `file attachment` を基本にする
* 画像は block 近傍のものだけ送る
* 1 block あたり 0〜2 枚程度に制限
* 全画像を一括送信しない

Copilot SDK は画像入力をサポートし、vision 対応モデルが必要。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk) ([GitHub Docs][5]))

## 8.3 hooks

初期実装から hooks を入れる。用途は次。

* 監査ログ
* 失敗時の文脈保存
* prompt サニタイズ
* 添付ファイル一覧の記録

hooks は Copilot SDK のセッションライフサイクルの各点で挙動を差し込める。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-hooks) ([GitHub Docs][7]))

## 8.4 session persistence

v1 必須ではないが、設計余地は残す。
長い変換や再開要件が出るなら sheet 単位で復元可能にする。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk/session-persistence) ([GitHub Docs][8]))

---

## 9. prompt / output 契約

## 9.1 prompt 側の考え方

prompt は skill ではなく **Python ツール本体** に置く。

### system 方針

* これは Excel の意味再構成タスクである
* 見た目の完全再現は不要
* 表、注記、図形、画像、グラフを統合して解釈する
* 不確実な箇所は断定しない
* Markdown は読みやすさ優先

### input

* sheet 名
* block JSON
* asset path / attachment
* ユーザー追加指示

## 9.2 LLM 入力例

```json
{
  "sheetName": "要件一覧",
  "blocks": [
    {
      "id": "b1",
      "kind": "table",
      "anchor": "A3:H20",
      "headers": ["ID", "要件", "優先度", "備考"],
      "rows": [["R-001", "ログイン...", "A", "..."]]
    },
    {
      "id": "b2",
      "kind": "shape",
      "anchor": "J2:M6",
      "shapeType": "textBox",
      "text": "重要: 優先度Aは法対応",
      "assetRef": "assets/sheet1/shape-01.png"
    }
  ],
  "instructions": {
    "targetFormat": "markdown",
    "style": "semantic",
    "preserveUnknowns": true
  }
}
```

## 9.3 応答契約

LLM にはいきなり Markdown だけを返させず、最低でも以下を含める。

```json
{
  "sheet_summary": "...",
  "sections": [...],
  "figures": [...],
  "unknowns": [...],
  "markdown": "..."
}
```

理由:

* 品質検証しやすい
* Markdown 再生成がしやすい
* 不確実箇所を追跡できる

---

## 10. skill 設計

## 10.1 skill の責務

* どういう依頼でこのツールを使うかを Copilot に教える
* 入力ファイルパス / 出力先を確認させる
* PowerShell で Python CLI を呼ぶ
* 結果の保存先を案内する

## 10.2 skill に持たせない責務

* Excel block 化ロジック
* prompt 本体
* Markdown 出力規約本体
* 応答 JSON 契約
* 実質的なビジネスロジック

## 10.3 `SKILL.md` の方針

* use when:

  * Excel を Markdown 化したい
  * 図形やグラフも解釈して文書化したい
* do:

  * 入力パス確認
  * 出力ディレクトリ確認
  * CLI 起動
* don’t:

  * 独自に解釈しない
  * ツール実行前に勝手にファイルを書き換えない

skills は `SKILL.md` と補助ファイルで構成され、必要なら shell を `allowed-tools` に入れられるが、広い事前承認は慎重に扱う。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/use-copilot-agents/cloud-agent/add-skills) ([GitHub Docs][4]))

---

## 11. 実装フェーズ

## Phase 1: MVP

必須

* `.xlsx/.xlsm` 読み取り
* block 検出
* 画像 / 図形 / グラフの抽出と画像化
* Copilot SDK local CLI 接続
* sheet 単位解釈
* Markdown / manifest 出力
* CLI 実行
* 最小 skill

## Phase 2: 品質強化

* hooks 強化
* debug JSON 出力強化
* sheet 分割条件の最適化
* 再試行戦略
* ログ整備
* 一部 session persistence

## Phase 3: 運用性強化

* 配布スクリプト
* skill 更新自動化
* モデル選択ポリシー
* 設定ファイル化
* resume 強化

---

## 12. 受け入れ条件

最低限、以下を満たすこと。

1. 表だけの sheet を意味の通る Markdown table にできる
2. テキスト入り図形を注記として Markdown に反映できる
3. グラフを画像付きセクションとして説明できる
4. 画像を attachment として LLM に渡し、本文へ反映できる
5. 1 workbook 複数 sheet をまとめて `result.md` にできる
6. 出力に `manifest.json` があり、各 block と asset の対応が追える
7. CLI だけで再現可能
8. skill から呼んでも、実行結果は CLI 実行と同じになる

---

## 13. リスクと対策

## 13.1 Excel COM 依存

リスク:

* 利用者端末に Excel が必要
* Excel プロセスの残骸やダイアログで不安定化しうる

対策:

* 1 ジョブ 1 Excel session
* `finally` で確実に終了
* 読み取り専用で開く
* 失敗時にプロセス掃除を入れる

## 13.2 `.xlsm`

リスク:

* マクロ実行リスク

対策:

* 読み取り専用
* マクロ無効前提
* 信頼済みファイルのみ対象

## 13.3 二重 LLM 構造

外側:

* Copilot skill

内側:

* Python ツール内の Copilot SDK

リスク:

* 責務がぶれる

対策:

* 外側は起動支援のみ
* 解釈責任は内側だけ

## 13.4 SDK preview

リスク:

* 仕様変動

対策:

* `llm/` 層に閉じ込める
* 中間表現を SDK 非依存にする

Copilot SDK は public preview とされ、機能や可用性は変わりうる。 ([docs.github.com](https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk/custom-skills) ([GitHub Docs][2]))

---

## 14. 実装禁止事項

* skill に変換ロジックを入れない
* workbook 全体を 1 prompt にしない
* row 単位で直接 LLM に投げない
* server-side / service 実行にしない
* SaveToHTML を主経路にしない
* 全画像を無差別に LLM に送らない
* Markdown を見た目再現に寄せすぎない

---

## 15. Codex が最初に着手すべき順序

1. Python パッケージ骨格を作る
2. CLI `convert` の空実装を作る
3. workbook 読み取りと block モデル定義を作る
4. OOXML drawings 読み取りを足す
5. Excel COM で chart / shape / range の画像化を作る
6. Copilot SDK local CLI 接続を作る
7. sheet 単位 prompt / response 契約を実装する
8. Markdown / manifest 出力を作る
9. skill を最小構成で追加する
10. テスト用 workbook で E2E を回す

---

## 16. 最後に一文で要約

**本体は Python CLI、知能はその中の Copilot SDK、skill は薄い起動ラッパー、Excel の見た目取得はローカル Excel COM。**

---

[1]: https://docs.github.com/en/copilot/how-tos/copilot-sdk/set-up-copilot-sdk/local-cli?utm_source=chatgpt.com "Using a local CLI with Copilot SDK"
[2]: https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk/custom-skills?utm_source=chatgpt.com "Using custom skills with the Copilot SDK"
[3]: https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2?utm_source=chatgpt.com "Considerations for server-side Automation of Office"
[4]: https://docs.github.com/en/copilot/how-tos/use-copilot-agents/cloud-agent/add-skills?utm_source=chatgpt.com "Adding agent skills for GitHub Copilot"
[5]: https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk?utm_source=chatgpt.com "Use Copilot SDK"
[6]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chart?view=excel-pia&utm_source=chatgpt.com "Chart Interface (Microsoft.Office.Interop.Excel)"
[7]: https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-hooks?utm_source=chatgpt.com "Use hooks"
[8]: https://docs.github.com/en/copilot/how-tos/copilot-sdk/use-copilot-sdk/session-persistence?utm_source=chatgpt.com "Session persistence in the Copilot SDK"
