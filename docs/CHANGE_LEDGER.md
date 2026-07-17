# 変更台帳（復元用メモ）

コード全文ではなく **対象・目的・戻し方** のみ記録する（[AGENT_HANDOVER.md](AGENT_HANDOVER.md) §8・§9.1）。

| 日付 | 対象 | 目的 | 戻し方 |
|------|------|------|--------|
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
