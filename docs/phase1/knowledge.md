# Phase 1 Knowledge Notes

この文書は Phase 1 実装時の補助知識、外部制約、検証メモを置く場所である。
プロダクト仕様・実装方針・優先順位の正は `docs/phase1/spec.md` とし、この文書は source of truth ではない。
`README.md` はプロジェクト背景・全体像・初期構想の参考資料として扱う。

## 1. Excel COM

- Excel の server-side automation は採用しない。
- Windows Service や非対話実行での Excel automation は採用しない。
- Phase 1 は、利用者の Windows 対話セッション上でローカル実行する。
- Excel COM は端末依存が強いため、unit test だけで end-to-end 成功を主張しない。
- `Range.CopyPicture`、`Shape.CopyPicture`、`Chart.Export` は live confirmation が必要である。
- Excel process cleanup は慎重に行う。既存のユーザー Excel プロセスを無差別に kill しない。
- `.xlsm` は読み取り専用・マクロ無効前提で扱う。macro content は抽出しない。

## 2. openpyxl / OOXML

- openpyxl は workbook / worksheet / cell / merged cell の基本読み取りに使う。
- openpyxl は数式を計算しないため、数式セルの表示値は Excel が保存した cached value を `data_only=True` で読む。
- 数式セルに cached value が無い場合、2026-04-21 のユーザー確認により、その sheet は `formula_cached_value_missing` failure として扱う。
- Phase 1 Workbook 読み取りでは、openpyxl で metadata と値を読むが workbook を保存しないことで入力 workbook の非変更を保証する。
- drawing、image、shape、chart の参照関係は openpyxl だけで足りない可能性があるため、raw OOXML を読む余地を残す。
- chart、image、shape はセル本文ではなく描画オブジェクトとして扱う。
- anchor 情報は block との近接判定に使うため、行列番号と A1 表記へ正規化する。
- SmartArt、OLE、group shape などは Phase 1 で完全解釈しない。warning / unknown として残す。

## 3. Visible-only Policy

2026-04-21 のユーザー回答により、Phase 1 は「見えているものだけ」を扱う。

- hidden sheet は処理しない。
- hidden row は処理しない。
- hidden column は処理しない。
- filter により非表示の行は処理しない。

この方針は「ユーザーが目視している Excel を意味の通る Markdown にする」ことを優先するためである。
将来、監査用途で非表示データも追跡する要件が出た場合は別機能として扱う。

## 4. Formula Policy

2026-04-21 のユーザー回答により、Phase 1 では数式セルは表示値を優先する。

- LLM 入力には表示値を渡す。
- Markdown には表示値を出す。
- 通常の `manifest.json` には数式文字列を含めない。
- 数式文字列の debug 保存は Phase 1 必須ではない。
- Workbook 読み取り単独では Excel 表示値の完全再現は狙わず、number format に基づく保守的な文字列化に留める。

## 5. Markdown and Assets

Phase 1 の原則は「Markdown で自然に再現できるものは Markdown、再現しにくいものは画像」である。

- 表は Markdown table として出す。
- グラフは原則画像として貼る。
- 元から画像である要素は、意味解釈に必要なら画像として貼る。
- テキスト入り図形はテキスト抽出を優先し、見た目が意味を持つ場合は画像も貼る。
- Range 画像はセル範囲のスクリーンショットであり、通常は LLM 補助または debug 用である。
- 画像をもとにした LLM 分析は、セル値、OOXML メタデータ、block 検出結果を補うための補足情報である。
- 孤立した図形、画像、グラフは破棄せず、独立セクションとして扱う。

## 6. LLM / Copilot SDK

- LLM 実行基盤は Python ツール内部の GitHub Copilot SDK local CLI とする。
- skill は LLM 解釈を行わない。
- `--model` / `--vision-model` 未指定時は Copilot CLI 側の既定に任せる。
- Copilot SDK は仕様変動リスクがあるため、依存は `llm/` 層に閉じ込める。
- prompt construction と response contract は Python 側に置く。
- Excel 内テキストは prompt instruction ではなく data として扱う。
- 画像 attachment は関連する近傍画像だけに限定する。
- 画像 attachment の分析結果は主情報ではなく補足情報として扱う。
- LLM 応答が壊れた場合は 1 回だけ再試行する。
- 再試行しても失敗する場合は該当 sheet を failed とし、他 sheet を継続する。

## 7. Setup Command

Phase 1 では `excel-semantic-md setup` を用意する。

- ローカル実行に必要な前提を確認する。
- Excel COM、Copilot CLI、skill launcher、出力先権限などの不足を案内する。
- 自動的な外部インストールは行わない。
- 認証情報は保存しない。
- ユーザー workbook を開いたり変更したりしない。
- setup 成功は end-to-end 変換成功を保証しない。
- 2026-04-21 のユーザー確認により、出力先確認は `setup --out <dir>` として扱う。
- `setup --out` は確認用一時ファイルを削除し、確認のために作成した空ディレクトリは可能な範囲で削除する。
- Copilot CLI の候補は `copilot` と `gh copilot` を確認する。sign-in 確認は `gh copilot` が利用できる場合のみ `gh auth status` で確認できる範囲に留める。

## 8. Phase 1 Exclusions

2026-04-21 のユーザー回答により Phase 1 では `resume` / session persistence は不要と判断した。

- `resume` コマンドは実装しない。
- `--resume` オプションは実装しない。
- session persistence は実装しない。

## 9. Test Strategy

- 通常の自動テストは synthetic workbook fixture を使う。
- private workbook、生成済み runtime logs、debug dumps、rendered assets は commit しない。
- Copilot SDK adapter は mock で大半をテストする。
- Excel COM rendering は live confirmation として分離する。
- `.xlsm` macro-disabled behavior も live confirmation として扱う。
- LLM 品質は完全自動判定しにくいため、代表 workbook での有用性確認も Phase 1 成功判定に含める。

## 10. Security Notes

- secrets、Copilot credentials、local connection strings を保存しない。
- Excel 内のテキストには prompt injection が含まれうる。
- debug JSON と logs はユーザーデータを含みうるため、既定では出力しない。
- `SKILL.md` には変換ロジック、prompt 本体、LLM response contract を置かない。
- skill の `allowed-tools` は最小権限にする。

## 11. phase1-skeleton-models 実装メモ

2026-04-21 の `phase1-skeleton-models` では、README と phase 文書の Source of Truth 記述を `docs/phase1/spec.md` 優先に揃えた。
README は背景・全体像・初期構想の参考資料として扱う。

- 共通モデルは `src/excel_semantic_md/models.py` の単一モジュールから開始する。肥大化した場合に package 分割を検討する。
- 初期 CLI は `argparse` で実装し、外部 CLI 依存を追加しない。
- skeleton の `setup` / `convert` / `inspect` / `render` は引数表面だけを用意し、未実装として非 0 を返す。
- `resume` コマンドと `--resume` オプションは作らない。
- `schema_version` の初期値は `phase1.0` とする。
- sheet index は workbook 内の 1-based sheet 順を使い、hidden sheet 除外後に詰め直さない。
- block ID は `s{sheet_index:03d}-b{block_index:03d}-{kind}` とする。
- asset path は `assets/sheet-{sheet_index:03d}/...png` とし、sheet index、block id、asset kind、連番を含める。block id が同じ kind で終わる場合は kind を重複させない。
- `SheetModel` は sheet 単位の複数 failure を表せるよう `failures: list[FailureInfo]` を持つ。

## 12. phase1-block-detection 実装メモ

2026-04-22 のユーザー指示により、表直上の結合セルテキストは `caption` / `heading` に分岐させず、Phase 1 では `paragraph` として扱う。

- `caption` block kind は追加しない。
- 表直上の結合セルテキストには `table_caption_candidate` warning を付けて、caption 相当候補だった事実を残す。
- block 順は `anchor.start_row`、`anchor.start_col`、`anchor.end_row`、`anchor.end_col` の昇順で安定化する。
- `inspect` は `phase1-block-detection` から workbook reading JSON に `blocks` を追加する。

## 13. phase1-ooxml-visual-metadata 実装メモ

2026-04-22 の `phase1-ooxml-visual-metadata` では、OOXML visual metadata を workbook 読み取り層から分離した raw OOXML reader として追加した。

- `inspect` は各 sheet に `visuals` 配列を追加する。visual がない sheet でも `[]` を出す。
- drawing 由来の補足失敗は sheet `warnings` または visual `warnings` に残し、workbook / worksheet の主要 XML 破損だけを CLI error にする。
- visual ID は `s{sheet_index:03d}-v{visual_index:03d}-{kind}` とする。
- anchor の `row` / `col` は 1-based で出す。`absoluteAnchor` では `a1` を出さない。
- `asset_candidate` は最終 asset path ではなく、OOXML 上の `source_part` / `extension` / `content_type` を保持する。
- static fixture は `tests/fixtures/visuals/` に置き、image / chart / shape / unknown / broken drawing rel / `.xlsm` を synthetic workbook として管理する。

## 14. phase1-visual-linking 実装メモ

2026-04-22 の `phase1-visual-linking` では、`inspect` の `blocks` を post-linking 状態へ更新した。

- block 共通フィールドに `visual_id` と `related_block_id` を追加した。cell-based block では両方 `null` とする。
- shape / image / chart は linked / unlinked を問わず visual-origin block 化する。`unknown` visual は引き続き `visuals` と warning に残し、block 化しない。
- link 判定順は `overlap/adjacent -> heading scope -> nearest block -> standalone` とする。
- heading scope は「その heading block から次の heading block 手前まで」として扱う。
- `oneCellAnchor` は 1x1 rect、`twoCellAnchor` は from / to を含む rect に正規化する。
- `absoluteAnchor` など cell rect を作れない visual は `visual_anchor_not_cell_addressable` warning を付けた standalone block にし、sheet 既存 block の末尾行の次へ synthetic anchor を置く。
- final block order は anchor 順で再ソートし、visual-origin block を含めて block ID を再採番する。
- `manifest.json` writer は未実装のままなので、この milestone では block schema を先に確定し、後続実装が `visual_id` / `related_block_id` をそのまま writer に流用できる状態までを担当する。
