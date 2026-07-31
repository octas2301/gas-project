# D入口 — Amazon 新規／既存相乗り（人間手順）

**状態**: **相乗り自己発＋新規同時 実機合格**（2026-07-31）。七味SEASONINGはSC再送済・在庫停止中確認中。FBA・フル（Yahoo再認証）は待ち  
**承認**: [LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)  
**PUT 詳細**: [D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)  
**前提**: v1.4 第1・第2段 API 実機合格／D×Amazon U3

---

## 0. できること

| D の選択 | 動き |
|----------|------|
| Amazonのみ＋新規 | 人間レ点子SKU → Da（`02` MAINゲートあり） |
| Amazonのみ＋相乗り自己発／FBA＋dry_run | レ点行 → VALID後、NF列`Amazon相乗りSKU`保存（送信在庫=常に0） |
| Amazonのみ＋相乗り＋prod | 保存済みNF列をsellerSkuとしてPUT（ALLOW_PROD＋確認必須） |
| 新規＋相乗り | 同じレ点子SKUを両方へ。新規=子SKU／相乗り=Amazon相乗りSKU（N列ASIN必須） |
| フル → Amazon | 開始前確認後、楽天→Yahoo→選択したAmazon方式 |

E／Z-21（⑩〜⑬）はテスト・復旧用に残る。X列は新規SKU式用で、相乗り時に変更不要。

---

## 1. 既存相乗りの手順（D）

1. `clasp push`（`コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`）
2. 対象**子行**にレ点、N列`ASINコード`を設定（X列は触らなくてよい）
3. NF列`Amazon相乗りSKU`は初回空欄で可。中央がすでにASINの子SKUはJAN置換せず、発送記号だけ `s→as`／`f→af`（D選択）に揃える（例 `…19s13`→`…19as13`）
4. `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=true`
5. D → Amazonのみ → **既存カタログに相乗り** → **相乗り自己発**または**相乗りFBA** → dry_run
6. status=VALID／issues=0、NF列保存を確認
7. （任意）`ALLOW_PROD=true` → 同じ選択でprod → 確認OK → ACCEPTED
8. トグルfalse

相乗り先はN列`ASINコード`のみ。O列／競合URLは使わない。  
マスタ在庫が0でなくても **Amazonへ送る数量は常に0**（非0出品はしない）。

### 1b. 実機合格記録（自己発）

| 日付 | モード | 結果 | runId | 備考 |
|------|--------|------|-------|------|
| 07-30 | dry_run／自己発 | OK=1 | `SPAPI_PUT_OFFER_CK_DRY_20260730_081735_99e` | X=`自己発送`のまま可。SKU=`sanky-B084RJSH7W-48as12` |
| 07-30 | prod／自己発 | OK=1 | `SPAPI_PUT_OFFER_CK_PROD_20260730_082325_9f2287` | NF列再利用。新規=未実行 |
| 07-31 | **新規＋相乗りprod 同時** | 新規=1／相乗りOK=1 | 新規 `LV4_20260731_010959_652276`／相乗り `SPAPI_PUT_OFFER_CK_PROD_20260731_011033_1f63ab` | 同一レ点行。新規=`CK_daba393f8055_B2_GENERATED.csv`（rows=2）／相乗り=`sanky-B01N5A6ESU-19as13` `ACCEPTED issues=0` qty=0 |

**同時実行の注意**: 新規側が冪等除外（`idempotentBlocked`）で0件になると、ファサードが例外で止まり**相乗りまで到達しない**。同じ親で作り直すときは 21-④ で該当 subBatchId を `UPLOAD_FAILED` にしてから再実行する。相乗りだけ進めたいときは新規のチェックを外す。

---

## 1c. Dレ点新規で在庫>0のとき

Dレ点新規（`source=child_ck`）は、マスタ在庫>0でも `SKIPPED_IN_STOCK` にせず GENERATED を作る（別カタログのノーブランドセットのため）。マスタ在庫は読取のみで書き換えない。バルクの在庫列は `inventoryMode` 準拠（既定 ZERO＝0）。

- ログ: `ckAllowInStock=true` と `[amazonApprovalLv4ResolveParents_] allowInStock parent=…`
- 旧挙動へ戻す: `APPROVAL_AMAZON_LV4_CK_ALLOW_IN_STOCK=false`
- 承認①経路（21-①等）は従来どおり在庫>0で `SKIPPED_IN_STOCK`
- `▼Lv4実行ログ(Amazon)` の1行目が `recordType` でないと停止する。別名退避して再実行する

---

## 1d. 新規カタログのGTIN免除証跡（カテゴリ別）

新規カタログ（track=B・`ノーブランド品`）は、**カテゴリごと**に免除証跡が必要。無いと `SKIPPED_GTIN_EXEMPTION` で停止する。HPCの証跡は食品には効かない。

1. 対象の**子SKU行に出品CK**、親行のカテゴリ列を埋める
2. メニュー **21-⑭ GTIN免除証跡を記録**
3. 同カテゴリのマスタASINがあればおすすめ文が表示される → OKでそのまま使う（キャンセルで手入力）
4. 確認ダイアログでOK → `▼Lv4実行ログ(Amazon)` に `EXEMPTION` 行が追記される
5. D を再実行

- おすすめは `過去成功ASIN: B0…（親SKU／カテゴリ）` ＋ `https://www.amazon.co.jp/dp/…`。SC申請URL／ケースIDは自動取得しない
- 既に有効な証跡があるカテゴリは追記されない（「登録済み」表示）
- 全カテゴリ `*` は `APPROVAL_AMAZON_LV4_EXEMPTION_ALL_CATEGORIES=true` のときのみ。追加警告あり。**原則はカテゴリ別**
- 手入力する場合は O=カテゴリ／P=`ノーブランド品`／Q=承認日／R=証跡。4つ揃わないと無効
- シートを別名退避すると証跡は引き継がれないため再記録が必要

---

## 1e. SC処理サマリで UPLOADED_OK 自動記録（21-⑮〜⑰）

**やること（初回だけ）**
1. Driveに監視フォルダを1つ作る → フォルダIDを `APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_ID` に設定  
2. `APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED=true`  
3. **21-⑯** でトリガー設置（間隔は既定15分。変更は `…_INTERVAL_MIN`＝5/10/15/30/60）

**日常**
- SC UP後、処理サマリ（`{subBatchId}_PACKAGED_…-processing-summary.xlsm`）を監視フォルダへ置くだけ  
- トリガーまたは **21-⑮** で `UPLOADED_OK` 追記 → ファイルは `_処理済` へ  
- 停止は **21-⑰** ＋ Property false  

**判定ルール（ファイル名のみ・中身は読まない）**
- 名前から `subBatchId` を抽出。`GENERATED` 行が無いIDは記録せずファイルも残す  
- 名前に `_NG`／`UPLOAD_FAILED` → `UPLOAD_FAILED`  
- 同じ状態が最新なら二重記録しない。1回最大20ファイル  
- **掲載完了の断定はしない**（実機確認は別）

---

## 1f. 七味（SEASONING）SC状況メモ（2026-07-31）

| SKU / ASIN | 経路 | SC在庫表示（外出先確認） |
|------------|------|--------------------------|
| 親 `sanky-4538872180149-oya`／`B0HC9S8PRN` | 新規カタログ | **停止中**（ブロック理由を確認） |
| 子 `sanky-B01N5A6ESU-19s13`／`B0HC9RRCBP` | 新規カタログ | **停止中**（同上） |
| 相乗り `sanky-B01N5A6ESU-19as13`／`B01N5A6ESU` | PUT prod | **停止中**（出品問題を修正） |
| 旧相乗り `sanky-B0B4RJSH7W-48as12` | PUT prod | **停止中**（同上） |

PACKAGED再送: `…_20260731_032459.xlsm`（KW1枠・粉末・グラム）。初回サマリは99016＋100521。再送結果はDownloads保存後に分析。  
サブ: U4でPT01〜03取得済 → 帰宅後手ZIP→Upload Images。

---

## 2. 合格目安

- [x] D新規＝レ点子SKUだけで従来Da
- [x] Dレ点新規が在庫>0でもGENERATEDされ、マスタ在庫が変わらない
- [x] X列が`自己発送`のままでも相乗りdry_runできる
- [x] Dで相乗り自己発を選べる（FBAは未検証）
- [x] 中央が既にASINの子SKUをそのまま使える
- [x] マスタ在庫>0でも quantity=0 でdry_run／prodできる
- [ ] N列空では停止
- [x] D prod＝ALLOW_PROD有り＋OKでACCEPTED（自己発1SKU）
- [ ] 確認キャンセルで PUT なし
- [ ] 21-⑫⑬・E-4 が従来どおり
- [ ] 作業後トグル false
- [ ] 相乗りFBA dry_run
- [x] 新規＋相乗り同時（2026-07-31・prod）
- [ ] フル → Amazon（楽天OK・**Yahooはrefresh token期限切れで失敗**）
- [ ] SEASONING七味のSC公開（停止理由解消）
- [ ] サブ画像ZIP UP

---

## 3. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-31 | §1e短文化。§1fに七味SC停止中メモ（親`B0HC9S8PRN`／子`B0HC9RRCBP`）。 |
| 2026-07-31 | **新規＋相乗りprod 同時 実機合格**（§1b）。`CK_daba393f8055_B2_GENERATED.csv`／`ACCEPTED issues=0`。 |
| 2026-07-31 | 21-⑮〜⑰ SC処理サマリ検知でUPLOADED_OK自動記録（§1e・ファイル名判定のみ）。 |
| 2026-07-31 | 21-⑭ にマスタ同カテゴリASINからの証跡おすすめを追加（OKでそのまま記録／キャンセルで手入力）。 |
| 2026-07-31 | 21-⑭ GTIN免除証跡の記録メニューを追加（§1d）。カテゴリ別が原則・`*` はProperty＋警告。 |
| 2026-07-30 | Dレ点新規は在庫>0でもGENERATED（§1c）。承認①経路は従来どおりスキップ。 |
| 2026-07-30 | 同一レ点行を新規＋相乗りへ同時出品（N列排他分割を廃止）。 |
| 2026-07-30 | 相乗り自己発 dry_run／prod 実機合格を記録。次＝FBA・新規・同時・フル。 |
| 2026-07-30 | Dで相乗り自己発/FBA選択。X非依存。ASIN済みSKUはそのまま。在庫>0でも送信0。 |
| 2026-07-30 | レ点本線・複数選択・Amazon相乗りSKU（NF列）へ更新。 |
| 2026-07-29 | 初版（実装に合わせた手順）。実機待ち。 |
