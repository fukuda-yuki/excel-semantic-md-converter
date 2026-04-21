# Phase 1 Task Breakdown

この文書は `docs/phase1/spec.md` を実装タスクに分解したチェックリストである。
プロダクト仕様・実装方針・優先順位の正は `docs/phase1/spec.md` とする。
Phase 1 では `resume` / session persistence を実装しない。

## 0. ドキュメント整合

- [x] `README.md` と Phase 1 文書が矛盾していないことを確認する。
- [x] `docs/phase1/plan.md` の判断と `docs/phase1/spec.md` の仕様が一致していることを確認する。
- [x] `docs/phase1/task.md` が placeholder ではなく実装タスクとして使えることを確認する。

## 1. プロジェクト骨格

- [x] `pyproject.toml` を作成する。
- [x] Python package 名を `excel_semantic_md` として定義する。
- [x] CLI entry point `excel-semantic-md` を定義する。
- [x] `src/excel_semantic_md/app/` を作成する。
- [x] `src/excel_semantic_md/excel/` を作成する。
- [x] `src/excel_semantic_md/render/` を作成する。
- [x] `src/excel_semantic_md/llm/` を作成する。
- [x] `src/excel_semantic_md/output/` を作成する。
- [x] `src/excel_semantic_md/cli/` を作成する。
- [x] `skills/excel-semantic-markdown/` を作成する。
- [x] runtime 出力が commit されないよう `.gitignore` を確認する。

## 2. 共通モデル

- [x] `Rect` を定義する。
- [x] `Block` の共通フィールドを定義する。
- [x] `HeadingBlock` を定義する。
- [x] `ParagraphBlock` を定義する。
- [x] `TableBlock` を定義する。
- [x] `ShapeBlock` を定義する。
- [x] `ImageBlock` を定義する。
- [x] `ChartBlock` を定義する。
- [x] `NoteBlock` を定義する。
- [x] `UnknownBlock` を定義する。
- [x] `SheetModel` を定義する。
- [x] `WorkbookModel` を定義する。
- [x] `AssetRef` を定義する。
- [x] `Warning` / `Unknown` / `Failure` の表現を定義する。
- [x] モデルを Copilot SDK と skill から独立させる。
- [x] JSON 直列化形式を定義する。
- [x] block ID 生成を安定化する。
- [x] asset ID 生成を安定化する。

## 3. CLI

- [x] `excel-semantic-md setup` を実装する。
- [x] `excel-semantic-md convert --input <path> --out <dir>` を実装する。
- [x] `excel-semantic-md inspect --input <path>` を実装する。
- [x] `excel-semantic-md render --input <path> --sheet <name>` を実装する。
- [x] `--model` を受け取り、未指定時は Copilot CLI 側の既定に任せる。
- [x] `--vision-model` を受け取り、未指定時は Copilot CLI 側の既定に任せる。
- [x] `--max-images-per-sheet` を受け取る。
- [x] `--save-debug-json` を受け取る。
- [x] `--save-render-artifacts` を受け取る。
- [x] `--strict` を受け取る。
- [x] `.xlsx` / `.xlsm` 以外を入力エラーにする。
- [x] 入力ファイルが存在しない場合に明確なエラーを返す。
- [x] 出力ディレクトリを作成する。
- [x] Phase 1 では `resume` コマンドを実装しない。
- [x] Phase 1 では `--resume` オプションを実装しない。

### 3.1 Setup

- [x] `setup` で Python package と CLI entry point を確認する。
- [x] `setup` で Windows 環境かどうかを確認する。
- [x] `setup` で Excel COM を利用できる可能性を確認する。
- [x] `setup` で Copilot CLI を利用できる可能性を確認する。
- [x] `setup` で Copilot CLI のサインイン状態を確認できる場合は確認する。
- [x] `setup` で skill launcher の配置先または利用方法を案内する。
- [x] `setup` で指定された出力先の書き込み可能性を確認する。
- [x] `setup` は自動的な外部インストールを行わない。
- [x] `setup` は認証情報を保存しない。
- [x] `setup` はユーザー workbook を開いたり変更したりしない。
- [x] `setup` の結果を、人間が読める診断レポートとして出力する。

## 4. Workbook 読み取り

- [x] openpyxl による workbook 読み取りを実装する。
- [x] `.xlsm` を読み取り専用・マクロ無効前提で扱う。
- [x] workbook を保存しないことを保証する。
- [x] sheet 順を保持する。
- [x] hidden sheet を除外する。
- [x] hidden row を除外する。
- [x] hidden column を除外する。
- [x] filter で非表示の行を除外する。
- [x] 数式セルは表示値を採用する。
- [x] 数式文字列を通常 manifest / LLM input に含めない。
- [x] cell comment を Phase 1 対象外として無視する。
- [x] note を Phase 1 対象外として無視する。
- [x] hyperlink を Phase 1 対象外として無視する。
- [x] セル内改行を保持する。
- [x] Excel 表示値を優先して日付・数値・パーセンテージを扱う。

## 5. Block 検出

- [ ] 使用範囲を推定する。
- [ ] 空行・空列による領域分割を実装する。
- [ ] 連続値領域から table 候補を検出する。
- [ ] 先頭行 / 先頭列の header 候補を検出する。
- [ ] 結合セルを取得する。
- [ ] 表上部の結合セルを heading または caption 候補にする。
- [ ] 表直上の単独テキスト行を heading 候補にする。
- [ ] 表ではない説明文領域を paragraph 候補にする。
- [ ] 過剰な意味推定を避けるため、warning を付けられるようにする。
- [ ] `inspect` で block JSON を確認できるようにする。

## 6. OOXML / Visual Metadata

- [ ] `.xlsx` / `.xlsm` を zip として読み、drawing relationship を取得する。
- [ ] image 参照を取得する。
- [ ] shape metadata を取得する。
- [ ] chart metadata を取得する。
- [ ] anchor / position を取得する。
- [ ] openpyxl で取得できない情報を raw OOXML で補う。
- [ ] SmartArt / OLE / group shape など未対応要素を warning または unknown として残す。
- [ ] workbook 内画像を asset 候補として抽出する。
- [ ] visual metadata を `inspect` 出力に含める。

## 7. 視覚要素の紐付け

- [ ] shape / image / chart と block の近接判定を実装する。
- [ ] table と隣接 shape を紐付ける。
- [ ] table と隣接 image を紐付ける。
- [ ] table と隣接 chart を紐付ける。
- [ ] heading 配下の visual を該当セクションに紐付ける。
- [ ] 紐付かない shape を独立 block にする。
- [ ] 紐付かない image を独立 block にする。
- [ ] 紐付かない chart を独立 block にする。
- [ ] 紐付け結果を manifest に残す。

## 8. Excel COM レンダリング

- [ ] Excel session wrapper を実装する。
- [ ] 1 job 1 Excel session を基本にする。
- [ ] workbook を読み取り専用で開く。
- [ ] 可能な範囲で Excel UI を非表示にする。
- [ ] 処理完了時に workbook を閉じる。
- [ ] 処理完了時に Excel session を閉じる。
- [ ] 既存ユーザー Excel プロセスを巻き込んで終了しない。
- [ ] `Range.CopyPicture` による Range 画像化を実装する。
- [ ] `Shape.CopyPicture` による Shape 画像化を実装する。
- [ ] `Chart.Export` による Chart PNG 出力を実装する。
- [ ] 元画像の asset 保存を実装する。
- [ ] `--save-render-artifacts` 指定時に補助 Range 画像を保存する。
- [ ] `render` コマンドで指定 sheet のレンダリング結果を確認できるようにする。
- [ ] Excel COM 依存箇所を live confirmation として分離する。

## 9. LLM Integration

- [ ] Copilot SDK local CLI adapter を実装する。
- [ ] sheet 単位で session を作成する。
- [ ] `--model` 指定時だけモデル指定を渡す。
- [ ] `--vision-model` 指定時だけ vision model 指定を渡す。
- [ ] prompt builder を実装する。
- [ ] Excel 内テキストを instruction ではなく data として扱う prompt を作る。
- [ ] 画像をもとにした LLM 分析を補足情報として扱う prompt を作る。
- [ ] attachment builder を実装する。
- [ ] 近傍画像だけを attachment にする。
- [ ] `--max-images-per-sheet` による画像数制限を実装する。
- [ ] 全画像を無差別に送らないことをテスト可能にする。
- [ ] LLM input JSON を構築する。
- [ ] LLM response JSON parser を実装する。
- [ ] response schema validation を実装する。
- [ ] 応答破損時に 1 回だけ再試行する。
- [ ] 再試行後も失敗した sheet を failed として記録する。
- [ ] Phase 1 では session persistence を実装しない。

## 10. 出力生成

- [ ] `result.md` writer を実装する。
- [ ] sheet 順を保持して Markdown を連結する。
- [ ] sheet ごとに見出しを出す。
- [ ] table を Markdown table として出力する。
- [ ] 図形テキストを本文または注記として反映する。
- [ ] 画像を必要に応じて Markdown に貼る。
- [ ] グラフ画像を Markdown に貼る。
- [ ] Markdown で再現できない要素を画像として貼る。
- [ ] 不確実な解釈を `result.md` に注記として出す。
- [ ] failed sheet を `result.md` に明示する。
- [ ] `manifest.json` writer を実装する。
- [ ] schema version を manifest に含める。
- [ ] input file name を manifest に含める。
- [ ] generated timestamp を manifest に含める。
- [ ] command options を manifest に含める。
- [ ] sheet list を manifest に含める。
- [ ] block list を manifest に含める。
- [ ] block id / kind / anchor を manifest に含める。
- [ ] asset path を manifest に含める。
- [ ] render status を manifest に含める。
- [ ] LLM status を manifest に含める。
- [ ] warning / unknown / failed sheet 情報を manifest に含める。
- [ ] `assets/` の sheet 別保存を実装する。
- [ ] 安定した asset 命名規則を実装する。
- [ ] `debug/` は `--save-debug-json` 指定時のみ出力する。
- [ ] `logs/` は既定では出力しない。

## 11. Orchestrator

- [ ] `convert` の全体フローを実装する。
- [ ] workbook 読み取りを呼び出す。
- [ ] visual metadata 抽出を呼び出す。
- [ ] block 検出を呼び出す。
- [ ] visual linking を呼び出す。
- [ ] render plan を作成する。
- [ ] Excel COM rendering を呼び出す。
- [ ] sheet 単位 LLM input を作成する。
- [ ] sheet 単位 LLM 解釈を実行する。
- [ ] sheet 単位失敗を failed として継続する。
- [ ] `--strict` 指定時は sheet failure を最終失敗にする。
- [ ] Markdown / manifest / assets を出力する。
- [ ] 途中失敗時も可能な範囲で中間状態を manifest に残す。

## 12. Copilot Skill

- [x] `skills/excel-semantic-markdown/SKILL.md` を作成する。
- [x] `run_excel_semantic_md.ps1` を作成する。
- [x] `examples.md` を作成する。
- [x] skill の責務を起動ラッパーに限定する。
- [x] 入力ファイルパス確認を記載する。
- [x] 出力ディレクトリ確認を記載する。
- [x] Python CLI 呼び出しを記載する。
- [x] 結果の保存先を案内する。
- [x] `SKILL.md` に prompt 本体を書かない。
- [x] `SKILL.md` に LLM response contract を書かない。
- [x] `SKILL.md` に変換ロジックを書かない。
- [x] `allowed-tools` を最小権限にする。

## 13. Fixture / Test

- [x] table only workbook fixture を作成する。
- [ ] table + text shape workbook fixture を作成する。
- [ ] table + image workbook fixture を作成する。
- [ ] table + chart workbook fixture を作成する。
- [x] multi-sheet workbook fixture を作成する。
- [x] hidden row / column / sheet workbook fixture を作成する。
- [x] formula display value workbook fixture を作成する。
- [x] workbook reading test を書く。
- [x] visible-only filtering test を書く。
- [x] formula display value test を書く。
- [x] merged cell handling test を書く。
- [ ] table detection test を書く。
- [ ] block ID stability test を書く。
- [ ] OOXML visual metadata extraction test を書く。
- [ ] visual linking test を書く。
- [ ] manifest generation test を書く。
- [ ] Markdown output composition test を書く。
- [ ] LLM response parser retry test を書く。
- [ ] Copilot SDK adapter は mock でテストできるようにする。
- [ ] Excel COM live confirmation は通常の自動テストから分離する。

## 14. Validation

- [x] `python -m pip install -e .` で local install できることを確認する。
- [x] `excel-semantic-md inspect --input <fixture>` を実行する。
- [ ] `excel-semantic-md setup` を実行する。
- [ ] `excel-semantic-md convert --input <fixture> --out <out>` を実行する。
- [ ] `result.md` が生成されることを確認する。
- [ ] `manifest.json` が生成されることを確認する。
- [ ] `assets/` が必要な画像を含むことを確認する。
- [ ] `--save-debug-json` 指定時のみ `debug/` が出ることを確認する。
- [ ] `--save-render-artifacts` 指定時のみ補助レンダリング成果物が出ることを確認する。
- [ ] LLM 応答破損時の 1 回再試行を確認する。
- [ ] 再試行失敗時に sheet failed として継続することを確認する。
- [ ] `--strict` 指定時に sheet failure が終了コードへ反映されることを確認する。

## 15. Review

- [x] README との整合性を確認する。
- [x] `resume` / `--resume` が Phase 1 対象外として整理されていることを確認する。
- [x] skill に変換ロジックが入っていないことを確認する。
- [ ] prompt と response contract が Python 側にあることを確認する。
- [ ] 画像をもとにした LLM 分析が主情報ではなく補足情報として扱われていることを確認する。
- [x] 中間表現が Copilot SDK 非依存であることを確認する。
- [ ] workbook 全体を 1 prompt にしていないことを確認する。
- [ ] row 単位 LLM 解釈になっていないことを確認する。
- [ ] 全画像を無差別に送っていないことを確認する。
- [ ] Excel COM cleanup の失敗リスクを確認する。
- [ ] `.xlsm` マクロ無効契約を確認する。
- [x] generated runtime outputs や private workbook を commit していないことを確認する。
