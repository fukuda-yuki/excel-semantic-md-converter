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
- subagent review の指摘を受け、heading scope は insertion index ではなく row range で判定するよう修正した。具体的には「heading の `anchor.end_row` より下、次の heading の `anchor.start_row` より上」を scope として扱う。
- synthetic anchor の fallback row は cell block だけでなく addressable visual anchor も含めた sheet 最大行の次を使う。
- `link_visuals()` は入力 `WorkbookModel` を破壊的に変更しないよう、入力 block を clone してから再採番する。

## 15. phase1-excel-com-rendering 実装メモ

2026-04-22 の `phase1-excel-com-rendering` では、`render` を Excel COM live confirmation 用コマンドとして実装した。

- `render` は `--out` を増やさず、一時ディレクトリへ確認用成果物を書き、標準出力 JSON で `temp_dir` と artifact 一覧を返す。
- `render` は workbook 読み取り、block 検出、OOXML visual metadata、visual linking、render planning、Excel COM rendering までを実行するが、LLM は呼ばない。
- Excel COM が使えない環境や object matching 失敗では、可能な範囲で JSON `failures` を返し、live confirmation 不足を明示する。
- image block は `Shape.CopyPicture` による確認用画像に加え、OOXML の元画像 part が取得できる場合は original asset も確認用成果物として複製する。
- internal render planner は `save_render_artifacts` を受け取り、将来 `convert --save-render-artifacts` から補助 Range 画像を要求できる形にした。
- shape / image / chart の Excel COM object matching は、まず anchor rect の exact match、その後 nearest match とし、shape text / alt text / chart title は曖昧性解消の補助ヒントに使う。

## 16. phase1-llm-integration 実装メモ

2026-04-22 の `phase1-llm-integration` では、Copilot SDK 依存を `src/excel_semantic_md/llm/` に閉じ込めた。

- `models.py` を含む既存中間表現層には `copilot` import を追加しない。
- Copilot SDK の runtime import は adapter 内の helper に限定し、未インストール環境でも他層 import が壊れないようにした。
- LLM 入力は `SheetModel` 1 枚単位で作り、workbook 全体や row 単位では作らない。
- prompt には「Excel 内テキストは instruction ではなく data」「block 構造が主情報」「画像分析は補足情報」「応答は JSON のみ」を明記する。
- LLM session は追加権限要求を自動承認しない。prompt injection 境界を広げないため、adapter は SDK の deny-by-default に任せる。
- attachment 候補は `RenderSheetResult.artifacts` から作り、優先順位は `markdown` 用の chart / image / shape を最優先、`related_block_id` を持つ artifact を次点、`render_artifact` の range を最後にする。
- LLM input JSON に含める asset `path` はファイル名だけにし、絶対パスは SDK attachment payload にだけ残す。
- `--max-images-per-sheet=0` は attachment なしとして扱う。
- attachment の上限超過時は重要度に加えて related block との近接度でも並べ替える。
- parser は plain JSON と fenced JSON の両方を受け付け、必須キー欠落や空 `markdown` を validation failure として扱う。
- retry は parser / schema validation failure のときだけ 1 回行い、実行例外の再試行は行わない。
- SDK cleanup 失敗は生例外を外へ漏らさず、sheet-level `failed` として返す。
- `--vision-model` は指定時だけ session config に渡すが、実 SDK の live confirmation は未実施のため互換性確認は残課題である。

## 17. phase1-output-generation 実装メモ

2026-04-23 の `phase1-output-generation` では、`convert` を output writer まで接続し、`result.md` / `manifest.json` / `assets/` / `debug/` を生成できるようにした。

- `WorkbookModel` は extraction / linking 層の source of truth のまま維持し、`convert` では sheet ごとの `SheetModel`、`RenderSheetResult`、`LlmRunResult`、CLI option を束ねる集約 DTO を新設した。
- workbook reading の `warnings` / `failures` は `detect_blocks()` で `SheetModel` に引き継ぐよう修正し、`formula_cached_value_missing` などの sheet 失敗が `convert` 最終出力まで残るようにした。
- `result.md` の successful sheet 本文は LLM `markdown` をそのまま主本文として使い、writer は不足している Markdown 用 asset 参照だけを後置する。failed sheet に対して block から代替本文は再構成しない。
- 公開 asset は `assets/sheet-{sheet_index:03d}/...` に安定命名で保存する。`role=markdown` は常に保存し、`role=render_artifact` は `--save-render-artifacts` 指定時のみ保存する。
- `manifest.json` は top-level に `schema_version`、`input_file_name`、`generated_at`、`command_options`、`sheets`、`blocks` を持つ整形 JSON とし、sheet ごとに render / llm status、block ごとに最終公開 asset path を含める。
- `debug/` は `--save-debug-json` 指定時のみ保存し、`workbook_extraction.json`、`linked_blocks.json`、`render_plan.json`、`llm_input.json`、`llm_response.json` を出力する。
- `render` コマンドの JSON は live confirmation 用のまま維持し、`convert` の manifest には `temp_dir` や一時絶対パスを含めない。
- 非 `strict` では sheet failed を `result.md` / `manifest.json` に残したまま他 sheet を継続し、`strict` では同じ出力を残したうえで最終終了コードだけ失敗にする。
- 空 block の visible sheet は render / LLM を強制せず、empty-sheet short circuit で successful sheet として扱う。

## 18. phase1-validation-review 実装メモ

2026-04-23 の `phase1-validation-review` では、task 11〜15 の validation 証跡と stale task 状態を整理した。

- `python -m pytest -q` は `86 passed`。
- `python -m excel_semantic_md.cli.main setup --out .tmp-validation-setup` を実行し、package import、CLI entry point、Windows 判定、Copilot CLI 候補、skill launcher、`--out` 書き込み確認が診断出力されることを確認した。
- `C:\Users\mwam0\AppData\Roaming\Python\Python314\Scripts\excel-semantic-md.exe` でも `setup` / `convert` / `convert --strict` を実行し、外部公開 CLI entry point 経由でも同じ結果になることを確認した。
- 現環境では `pythoncom` / `win32com.client` が未導入のため、`setup` は Excel COM を「not available or not confirmed」と報告する。
- `python -m excel_semantic_md.cli.main convert --input tests/fixtures/visuals/no-visuals.xlsx --out .tmp-validation-convert` を実行し、非 `strict` では `result.md` / `manifest.json` を出力したうえで sheet failure を残せることを確認した。
- 同 fixture で `--strict` を付けると、同じ出力を残したまま終了コード `1` を返すことを確認した。
- 現環境の `convert` 実行では `pywin32` 不足により render failure となるため、Copilot SDK local CLI behavior と vision attachment behavior の live confirmation は pending のまま残す。
- editable install で生成される `excel-semantic-md.exe` は `C:\Users\mwam0\AppData\Roaming\Python\Python314\Scripts` に存在するが、この shell の `PATH` には含まれていない。
- sheet pipeline 内の未捕捉例外は、可能な限り sheet-level `FailureInfo` に正規化して他 sheet 継続と output writer 到達を優先する。
- `manifest.json` の sheet-level `llm` payload には、Copilot SDK から current model を取得できた場合だけ `used_model` を含める。

## 19. phase1-requirements-implementation-review 対応メモ

2026-04-23 の `phase1-requirements-implementation-review` 対応では、既存仕様を変更せず、受理済みレビュー指摘のうち仕様安全性・公開出力契約・回帰防止に関わる項目だけを修正した。

- `.xlsm` の Excel COM rendering は、`AutomationSecurity = 3` を設定できない場合に workbook を開かず fail closed とする。`.xlsx` は従来どおり best effort のまま扱う。
- `manifest.json` の warning / failure details は、キー名が `path` でない例外文字列でもローカル絶対パスを redacted にする。通常 manifest に temp dir や workbook 絶対パスを漏らさない方針を補強した。
- managed output replacement は、既存出力を backup へ退避してから staging 出力を移動し、公開移動に失敗した場合は新規移動済み出力を削除して旧出力を復元する。
- LLM request は `build_llm_request()` で attachments、LLM input、prompt を一括生成する。`convert` の debug 用 `llm_input_payload` と `GitHubCopilotSdkAdapter` の実送信用 prompt / attachments は同じ request 由来にする。
- workbook read -> block detection -> visual metadata -> linking の重複統合は今回は見送った。仕様不一致ではなく保守性改善であり、今回の安全性修正に混ぜると過剰リファクタになるため。
- 自動テストは `python -m pytest -q` で `95 passed`。Copilot SDK local CLI behavior、vision attachment behavior、実 Excel COM、実 `.xlsm` macro-disabled behavior は引き続き live confirmation 対象である。

## 20. 2026-04-24 再レビュー引き継ぎメモ

2026-04-24 の「要件定義をもとに、仕様と実装をレビュー」では、実装変更を入れずに review note だけを更新した。

- `convert` は現状、cell-based block に対しても常に `range_copy_picture` を計画するため、単純な table / paragraph workbook でも Excel COM が使えないと sheet failed になる。`--max-images-per-sheet 0` でもこの依存は消えない。
- `--max-images-per-sheet` 未指定時は render artifact を全件 Copilot 添付候補にするため、Phase 1 の「全画像を無差別に送らない」契約に対して過剰実装になっている。
- OOXML image の `target_part` は存在確認だけで publish / attach 対象になっており、`content_type` が画像かどうかを検証していない。細工した workbook では非画像 part や macro binary を asset として複製しうる。
- `render` CLI は `build_render_plan()` や `render_with_excel_com()` の予期しない例外を JSON failure に正規化せず、そのまま例外終了しうる。
- `render_with_excel_com()` は artifact ごとの failure 正規化が `RenderTaskError` だけなので、素の COM/OSError では残り artifact を継続せず sheet-level generic failure に崩れる。
- 表示値の文字列化は `%` と整数化された float を主に扱っており、通貨、桁区切り、小数桁指定など `number_format` 依存の表示値再現は未充足である。
- OOXML warning-only 継続経路と attachment 互換フォールバックには、追加の回帰テスト余地が残っている。

## 21. 2026-04-24 Re-review Fix Batch 実装メモ

- `convert` は render planning の結果から cell-based `range_copy_picture` を既定 render 対象から外す。render item が空の sheet は Excel COM を使わずに LLM へ進める。
- `--max-images-per-sheet` 未指定時の既定値は 3。既定添付候補は `chart` / `image` / `shape` の主要 visual に限定し、`range_copy_picture` は既定 attachment 候補から外す。
- `--max-images-per-sheet 0` は attachment 0 件を意味する。visual block が無い sheet では render を起動しない。
- OOXML image original asset は `image/*` content-type allowlist を満たす場合だけ `ooxml_image_copy` を計画する。non-image content type、missing target、missing part は warning-and-skip にする。
- `render` CLI は planning/rendering の unexpected exception を JSON failure に正規化する。`render_with_excel_com()` は artifact 単位の通常例外も failure 化して後続 item を継続する。
- 表示値フォーマットは conservative subset として percent / currency / grouping / fixed decimals を扱う。scientific / fraction / accounting は fallback を維持する。
