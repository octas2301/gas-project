# 変更台帳（復元用メモ）

コード全文ではなく **対象・目的・戻し方** のみ記録する（[AGENT_HANDOVER.md](AGENT_HANDOVER.md) §8・§9.1）。

| 日付 | 対象 | 目的 | 戻し方 |
|------|------|------|--------|
| 2026-07-25 | U3実機合格記録／`D_MENU_U2_…REQUIREMENTS.md`（新規）／CURRENT_PHASE 等 | U3クローズ＋U2起草。コード追加なし（本行） | **Git**: `git revert` |
| 2026-07-25 | `コード.js`（D amazon／full_amazon・`runBatchExportAmazonFacade`）／`AmazonApprovalExport.js`（21-①戻り値）／`D_MENU_U3_HUMAN_RUN.md`／要件・CURRENT_PHASE／AGENT_HANDOVER | U3 v1: 薄いファサード。トリガーにAmazon非搭載 | **Git**: `git revert`。Property OFF |
| 2026-07-24 | `D_MENU_AMAZON_FACADE_REQUIREMENTS.md`（新規）／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4・POCリンク | D本線Amazonファサード要件U0。T3保留。コードなし | **Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER（T2 PoC成功記録） | runId `R2T2_20260724_221107_7f9cf7`・URL画像表示。コード追加なし | **Property** `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`。**Git**: `git revert` |
| 2026-07-24 | `AmazonDriveImageExport.js` HMAC Byte[]揃え／clasp push | SigV4で (String,Byte[]) 例外を修正。T2のみ | **Git**: `git revert`／再push |
| 2026-07-24 | CURRENT_PHASE §0（clasp push済反映） | T2: `clasp push` 8 files 成功。コード差分なし（反映のみ） | 戻し不要（GASは前回リビジョンへ） |
| 2026-07-24 | CURRENT_PHASE §0 帰宅引き継ぎ／AGENT_HANDOVER | 自宅PC続行用に本日分を記録。コード追加なし（T2は既存） | **Git**: `git revert` |
| 2026-07-24 | `AmazonDriveImageExport.js`（新規）／`コード.js`（21-⑥）／LV4_T2_HUMAN_RUN／POC／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | T2: Drive→R2 MAIN1枚PoC。トグル既定false | **Property** `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`。**Git**: `git revert`／新規js削除＋メニュー削除 |
| 2026-07-24 | `RAKUTEN_NAV_STAGE1_HUMAN_RUN.md`／AI_ROUTING §5.1静的／`LV4_T2_IMPLEMENTATION_APPROVAL.md`／CURRENT_PHASE／AGENT_HANDOVER | Nav・AI静的＋T2承認書。GAS未実装 | **Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE／AGENT_HANDOVER／Drive `05` success JSON／FOOD_在庫確認_STATUS | FOOD v5成功（6バリエーション）。コードなし | **Git**: `git revert`／Drive success JSON削除可 |
| 2026-07-24 | Drive `03/…_v5_corrective.xlsm`・`04/…_MISSING5_MAIN_for_SC.zip`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | FOOD欠け5子corrective（親部分更新・30s175除外・袋数タイトル）。GASなし | **Drive**: v5ファイル削除。**Git**: `git revert` |
| 2026-07-24 | CURRENT_PHASE §0／LV4_R2_IMAGE_PIPELINE_POC（§2.1二モード）／AGENT_HANDOVER／CHANGE_LEDGER | 外出リモート引き継ぎ＋サブREUSE/ONLY方針。コードなし | **Git**: `git revert` |
| 2026-07-24 | Drive `04` 配下01〜06作成／POC §2 ID表／CURRENT_PHASE／`DRIVE_04_FOLDER_IDS.md`(Downloads) | フォルダ実体＋Folder ID控え。GAS・Property未設定 | **Drive**: 空フォルダ削除可。**Git**: `git revert` |
| 2026-07-24 | `LV4_R2_IMAGE_PIPELINE_POC.md`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4要件 | Drive起点GAS案を正に。xlsm自動は提案のみ。コードなし | **Git**: `git revert` |
| 2026-07-24 | `docs/org/LV4_R2_IMAGE_PIPELINE_POC.md`／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER／LV4要件リンク | R2×Amazon画像PoC設計（楽天・Yahoo対象外）。コードなし | **Git**: `git revert`／PoC md削除 |
| 2026-07-23 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER | 帰宅引き継ぎ: HPC suburl試験・FOOD再UP禁止。コードなし | **Git**: `git revert` |
| 2026-07-23 | LV4要件 §11.0・§11表・§11.5.3画像／CURRENT_PHASE／AGENT_HANDOVER／CHANGE_LEDGER | HPC §11-1〜8・10＋U5クローズ。画像ZIP優先。FOOD／M2／21-⑤は別ゲート。コードなし | **Git**: `git revert` |
| 2026-07-23 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER／accepted_values_db(Downloads) | 外出リモート引き継ぎ: HPC 21-③済・FOOD待機。コードなし | **Git**: `git revert`／DBはDownloads削除可 |
| 2026-07-22 | LV4要件 §11.5.4・§11.6／CURRENT_PHASE／`accepted_values_db/` | FOODテンプレ確認＋成功値DB方針。種子JSONはDownloads配下 | **Git**: `git revert`／DBフォルダ削除可 |
| 2026-07-22 | CURRENT_PHASE §0／AGENT_HANDOVER／CHANGE_LEDGER | 外出引き継ぎ: corrective画像18320・中期R2方針。コードなし | **Git**: `git revert` |
| 2026-07-22 | LV4要件 §6.1.2・§10・§11.5.3／CURRENT_PHASE / AGENT_HANDOVER / CHANGE_LEDGER | 修正登録（SKU維持）・21-⑤・実データ7行目起点。GASメニュー実装は後続。ローカル corrective_link.xlsm | **Git**: `git revert`／PACKAGEDファイル削除 |
| 2026-07-21 | LV4要件 §11.5・§4／CURRENT_PHASE / AGENT_HANDOVER | HPC列対応ドラフト＋マスタ必須候補方針＋成功値辞書化メモ（後続）。コードなし | **Git**: `git revert` |
| 2026-07-21 | LV4要件 §5・§6.1.1・§14・§17 Q5／CURRENT_PHASE / CHANGE_LEDGER | subBatchId方針を新番号に統一＋push手順。コードは既存実装どおり | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js | 再レビュー採用: DRY_RUN／冪等汚染／subBatchId単調／ブランド厳密／archive失敗停止 | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js | 実装レビュー採用修正（レジューム・SKIPPED記録・GTIN・冪等・追記ログ）。戻し: git revert | **Git**: `git revert` |
| 2026-07-20 | AmazonApprovalExport.js（新規）／ApprovalQueue.js／コード.js（メニュー21）／LV4要件・CURRENT_PHASE / AGENT_HANDOVER | Lv4実装。ENABLED=false既定。戻し: Property OFF＋git revert＋新規js削除 | **Git**: `git revert` / Property `APPROVAL_AMAZON_LV4_ENABLED=false` |
| 2026-07-20 | LV4要件 §3.2・§2・§13・§17／AMAZON_REQUIREMENTS §3／多数決 §4.2／CURRENT_PHASE / AGENT_HANDOVER | Q15/Q16＋列メモ注記。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §3.4–3.6・§6・§17／多数決／CURRENT_PHASE / AGENT_HANDOVER | Q11–Q14反映。方針Q&A一通り閉じ。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §3.1–3.2・§3.5・§17／多数決 §4.1／CURRENT_PHASE / AGENT_HANDOVER | Q7–Q10b: 3モール同一承認・TRACK=B強制・SKU/メーカー品番。コードなし | **Git**: `git revert` |
| 2026-07-20 | LV4要件 §1.4・§17／多数決メモ §4.1／CURRENT_PHASE / AMAZON_REQUIREMENTS | 社長Q&A: D-1 PACKAGED・JAN出力規則・バリエーションのみ・再生成＋ログ。コードなし | **Git**: `git revert` |
| 2026-07-20 | `docs/org/LV4_THREE_REVIEW_MAJORITY.md`・`LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN / AMAZON_REQUIREMENTS | Lv4三点レビュー多数決保存＋社長確定反映（在庫書込禁止・DONE分離・親SKU・GTINゲート）。コードなし | **Git**: `git revert` |
| 2026-07-20 | `docs/org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN / AMAZON_REQUIREMENTS | Lv4要件ドラフト（M1=B新規・M2=A既存・手動UP・限定ツール）。コードなし | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-20 | CURRENT_PHASE / LV3要件 §8.1・§8.2 / AGENT_HANDOVER / LEVELLED_PLAN | Lv3人間検収完了を共有。フォーカスを Lv4 Amazon へ。在庫列二重ヘッダー運用メモ | **Git**: `git revert` |
| 2026-07-20 | `YahooApprovalExport.js`（resolveログ） | Lv3スキップ切り分け: sheet名・在庫生値・ORPHAN/IN_STOCK を Logger 出力 | **Git**: `git revert`／当該 Logger 行を削除 |
| 2026-07-20 | `YahooApprovalExport.js`（新規）・`ApprovalQueue.js`（Yahoo読取）・`コード.js`（メニュー20）・Lv3 docs | Lv3: 承認①済Yahoo子を案A＋主25分／副ユニーク50で `runYahooExport` 呼出。Yahoo本体非改変 | **Property**: `APPROVAL_YAHOO_LV3_ENABLED=false`。**Git**: `git revert`／新規js削除＋メニュー20削除 |
| 2026-07-20 | `docs/org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md`・CURRENT_PHASE / AGENT_HANDOVER / LEVELLED_PLAN | Lv3要件ドラフト（runYahooExport呼出・主25分／副ユニーク50運用揃え）。コードなし | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-20 | CURRENT_PHASE / LV2要件 §8.1・§10.1 / AGENT_HANDOVER / LEVELLED_PLAN | Lv2人間検収完了を共有。フォーカスを Lv3 Yahoo へ。目視手順追記 | **Git**: `git revert` |
| 2026-07-20 | `RakutenApprovalExport.js`（案Aレ点）・Lv2 docs | 案Aを「承認済み親＋紐づく子も一時レ点」に変更。CSV本体非改変。バリエーション親のシングルSKU誤認を解消 | **Git**: `git revert`。関数 `rakutenApprovalLv2ApplyPlanACheckboxes_` を親だけONに戻す |
| 2026-07-21 | `AmazonApprovalExport.js`（価格解決）・LV4 docs | 親`販売価格amazon`空→承認済み子へフォールバック。`[Lv4Price]`ログ | **Git**: `git revert` |
| 2026-07-21 | `AmazonApprovalExport.js`（カテゴリ解決）・LV4 docs | T列`カテゴリー`フォールバック＋見出しゆれ対応。`[Lv4Cat]`ログ | **Git**: `git revert`／当該関数を旧2行lookupに戻す |
| 2026-07-17 | CURRENT_PHASE / LV1要件 §8.1 / AGENT_HANDOVER / LEVELLED_PLAN / LV2前提 | Lv1人間検収完了をプロジェクト全体へ共有。フォーカスを Lv2 へ | **Git**: `git revert` |
| 2026-07-17 | `ApprovalQueue.js`・`Yahoo.js`(doGet分岐のみ)・`コード.js`(メニュー18)・`appsscript.json` | Lv1承認キュー実装。シート追記＋Web承認。出品API/CSV非呼出。Property既定false | **Property**: `APPROVAL_QUEUE_V1_ENABLED=false`。シート削除可。**Git**: `git revert` |
| 2026-07-17 | `docs/org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md` 等 | Lv1承認キュー要件（シート列・batchId・レ点抽出・Web・EC書込なし） | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-17 | org 多数決・Lvプラン・マトリクス・CURRENT_PHASE 等 | Lv0最終承認・モール順楽天先・レ点/スキップ/U1手動上書きの文書化（コードなし） | **Git**: `git revert` |
| 2026-07-17 | `docs/org/LEVELLED_IMPLEMENTATION_PLAN.md`、CURRENT_PHASE / AGENT_HANDOVER 追記 | AI組織の Lv0〜5＋並行P の実装順叩き台（コードなし） | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-17 | `コード.js`（`menuDiagnoseRakutenNavigationGenreStage1Write` 等）、`docs/RAKUTEN_NAV_GENRE_STAGE1.md` | Stage1: 専用シート `▼診断(楽天ジャンルNav)` への追記のみ（マスタ/CSV非書込・楽天CSV非改変）。Property 既定 false | **Property**: `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED=false`。シート削除可。**Git**: `git revert`。メニュー17-⑥/99-⑩と Stage1 関数を削除 |
| 2026-07-17 | `docs/RAKUTEN_NAV_GENRE_STAGE1.md`（要件） | Stage0完了後の Stage1：テストシート追記隔離の要件正本化 | **Git**: `git revert` または当該ファイル削除 |
| 2026-07-16 | `parseRakutenNavigationGenreResponse_` | NavigationAPI 2.0 の `nameJa` / `nameJaPath` を解釈して namePath を出す | **Git**: `git revert` |
| 2026-07-16 | `diagnoseRakutenNavigationGenreGetOne_` の URL | NavigationAPI 2.0 を `/es/2.0/navigation/genres/{id}` に修正（GF0002 404 対策） | **Git**: `git revert` または URL 配列を旧 query 形式に戻す |
| 2026-07-16 | `コード.js`（`menuDiagnoseRakutenNavigationGenreGet` 等）、`docs/RAKUTEN_NAV_GENRE_DIAG.md` | 楽天ジャンルID NavigationAPI 読取疎通診断（マスタ/CSV書込なし・CSV経路非改変） | **Property**: `RAKUTEN_NAV_GENRE_DIAG_ENABLED=false`。**Git**: `git revert`。診断関数・メニュー17-⑤/99-⑨を削除 |
| 2026-07-15 | `docs/org/PHASE0_THREE_REVIEW_MAJORITY.md`・`THREE_REVIEW_RUNBOOK.md`、憲章・マトリクス更新、`.cursor/rules/three-review-runbook.mdc`、CURRENT_PHASE / AGENT_HANDOVER | 3者多数決の記録・採用項反映・親1＋並列サブ3運用の正本化（実装コードなし） | **Git**: `git revert`。ルール削除は `three-review-runbook.mdc` を削除 |
| 2026-07-15 | `docs/org/AI_ORG_CHARTER.md`・`docs/org/AI_APPROVAL_MATRIX.md`、および `CURRENT_PHASE.md` / `AGENT_HANDOVER.md` の参照追記 | 一人社長＋AI部門の組織憲章・承認マトリクスを docs 正本化（実装なし。Phase0→3者検証前提） | **ファイル削除**: `docs/org/` の当該2ファイルを削除し、CURRENT_PHASE / AGENT_HANDOVER / 本台帳の該当行を戻す。**Git**: `git revert` |
| 2026-07-11 | `Yahoo.js`（`findYahooMasterHeaderRowIndex_`・`getYahooMasterHeaderContext_`・`_loadMasterData`・`updateMasterYahooId`・`updateMasterDeleteFlag`・`showDeleteSelectionDialog`・`listDeletableItems`） | マスタヘッダー行の `data[7]` 固定を廃止し、出品 Builder と同じ動的検出に統一（`docs/REVIEW_Yahoo_js.md` 指摘#1） | **ファイル**: `Yahoo.js.bak_before_header_unify_20260711` または `Yahoo.js.local_backup` を `Yahoo.js` に上書きコピー。**Git**: 本変更コミット後は `git revert` |
| 2026-05-09 | `コード.js`（`keepaProductPrimaryImageToken_`・`keepaAmazonMediaUrlFromImageToken_`・`getKeepaProductImageUrl`・`getKeepaProductImageUrlsAll`・`getKeepaProductImageUrlsMaxForCompetitorTest_`・`buildKeepaTableRow`） | Keepa API が返す **images 配列（l / m）** を解釈する。従来の image / imagesCSV のみでは現行レスポンスで画像URLが空になる問題を修正 | **Git**: `git revert`。**GAS**: 同ブロックを旧実装に戻す |
| 2026-05-06 | `コード.js`・`.claspignore` | R-Cabinet 診断ログ（パイプラインタグ・送信 bytes・JPEG 比・cap 付きサムネ）、Capacity 時に長辺 cap を段階的に下げる再試行、**`clasp push` で本番反映**。`.claspignore` に `**/*.html` 等で誤 push 防止 | **Git**: `git revert`。**GAS**: script.google.com で誤追加の `.html` を削除（以前の誤 push 分）。再試行を単発に戻す場合は `executeRenameAndUploadFromMatrixProgrammatic_` 内の `capList` ループを旧 1 回 `tryFetch` に戻す |
| 2026-05-06 | `RAKUTEN_CABINET_MAX_IMAGE_BYTES`・`convertBlobToJpegForRakuten_`・`prepareBlobForRakutenCabinetUpload_`・Drive サムネ系（`コード.js`） | R-Cabinet `Capacity` 回避のため目安を **1,900,000 バイト**に締め、サムネ／オリジナルとも **送信直前に JPEG 化**し **JPEG 後バイト**で上限判定する一本化 | **Git**: `git revert`。**手動**: 定数を `2*1024*1024` に戻し `convertBlobToJpegForRakuten_` を削除、サムネ戻り値の `convertBlobToJpegForRakuten_` 呼び出しと `prepare` の二段ロジックを旧実装に戻す（判断は `docs/BATCH_EXPORT_IMAGE_GATE_REQUIREMENTS.md`） |
| 2026-05-06 | `generateRakutenCSV`（`コード.js`）・`getRakutenSpecHtmlCell_` | PC/スマホ説明文の「商品情報」表を、▼マスタ列なしの明示フォールバック（メーカー名・マスタシリーズ／特記すべき原材料・原材料／括弧付き賞味・保存＋従来名）で取得 | **Git**: `git revert`。スペック4行を旧 `getMVal(ブランド/シリーズ/原材料/賞味/保存)` に戻し `getRakutenSpecHtmlCell_` 関数削除 |
| 2026-05-02 | `generateRakutenCSV`（`コード.js`）・ヘルパー `isMasterCellEmptyForRakutenAttr_` / `isRakutenAttrValueEmptyForCleaning_` | マルチSKUで親の `楽天セット数`（0/空）が `virtualRow`・`parentDataValues` 経由で子の属性値を潰し、かつ `attr_value` の `!val` で項目・単位だけ消える問題を防止 | **Git**: 該当コミットを `git revert`。**手動**: `virtualRow` のコピー・SKU属性継承・`attr_value` クリーニング3点を元に戻す（判断経緯は `docs/RAKUTEN_CSV_ATTR_AND_SETCOUNT_REQUIREMENTS.md`） |
| 2026-03-22 | `runProductNameProposalsForRows`（`コード.js`） | OpenAI 失敗時もマスタ商品名があればバリエーション（単位・内容量）を実行。Script Property `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL`（既定 `true`）で旧挙動に戻せる | **Git**: 本変更のコミットを `git revert <commit>`。**運用**: Script Properties に `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL` = `false` で旧挙動（OpenAI 失敗行はバリエーションもスキップ） |
| 2026-03-22 | `inferVariationFromAsinCircleForJan_`・`pickPerSetContentFromCircleTitles_` 等（`コード.js`） | 同一 JAN ブロック内の **全 ◎ 行**から 1セットあたり内容量を総合判断（g/ml 優先）。`CIRCLE_COMBINED_PER_SET_CONTENT`（既定 `true`）で無効化すると旧「先頭ヒット打切り」 | **Git**: `git revert <commit>`。**運用**: `CIRCLE_COMBINED_PER_SET_CONTENT` = `false` |

**コミットハッシュ**: ローカルで `git log -1 --oneline` を実行したうえで、上表に追記すること。
