# D入口 — Amazon 新規／既存相乗り（人間手順）

**状態**: **P0実装済**。相乗り自己発＋新規同時・**相乗りFBA**は実機合格（[A1](D_MENU_LANE_A1_FBA_HUMAN_RUN.md)）  
**承認**: [LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)／[LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md](LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)／[LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md](LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)  
**PUT 詳細**: [D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)  
**前提**: v1.4 第1・第2段 API 実機合格／D×Amazon U3

---

## 0. できること

| D の選択 | 動き |
|----------|------|
| Amazonのみ＋新規 | 人間レ点子SKU → Da（`02` MAINゲートあり） |
| Amazonのみ＋相乗り自己発／FBA＋prod（既定） | レ点行 → PUT（ALLOW_PROD＋確認）。**SKU列空ならprodで生成・成功後に系統別列へ保存**（通常運用はdry_run不要） |
| Amazonのみ＋相乗り＋dry_run（上級・折りたたみ） | VALID後、自己発→`Amazon相乗りSKU`／FBA→`Amazon相乗りSKU_FBA`（任意検証） |
| 送信在庫＝0（既定） | 承認①相当。**新規GENERATED在庫列＝0**／相乗り quantity=0（FBAはquantity非送信） |
| 送信在庫＝マスタ在庫 | 承認②。`ALLOW_MASTER_QTY=true`＋確認に件数・qty内訳（**新規＋相乗りとも**）。**子の**空/負/非数は**停止**。Track B **親行の inventory は常に0**（親マスタ在庫は読まない）。マスタ列は非書込 |
| 新規＋相乗り | 同じレ点子SKUを両方へ。新規=子SKU／相乗り=系統別列（**N列ASINがある行だけ**相乗り。空は行スキップ） |
| フル → Amazon | 開始前確認後、楽天→Yahoo→選択したAmazon方式（prod既定に含む） |

段階実行は **D補助. Amazon段階実行（旧E）** または **Z → E互換**。Z-21（⑩〜⑬）はテスト・復旧用に残る。X列は新規SKU式用で、相乗り時に変更不要。

---

## 1. 既存相乗りの手順（D）

1. `clasp push`（`コード.js`／`AmazonApprovalExport.js`／`AmazonSpapiPut.js`）
2. 対象**子行**にレ点。相乗りする行だけ N列`ASINコード`を設定（X列は触らなくてよい）。**ASIN空のレ点行はスキップされ、全体は止まらない**（有効ASINが0件のときだけ停止）
3. **デュアル Phase1**: 自己発は `Amazon相乗りSKU`（初回空可）／FBAは `Amazon相乗りSKU_FBA`（列追加必須・初回空可）。中央がすでにASINの子SKUはJAN置換せず、発送記号だけ `s→as`／`f→af`（D選択）に揃える
4. （本番常時ON）`APPROVAL_AMAZON_SPAPI_PUT_ENABLED`／`ALLOW_PROD` は **未設定または true**（毎回切り替え不要。緊急停止だけ明示 false）
5. D → Amazonのみ → **既存カタログに相乗り** → 自己発orFBA → **prod（既定）** → 確認OK → ACCEPTED。**SKU列が空でも生成してPUT**し、成功後に当該系統列へ保存
6. （任意・上級）dry_run で VALID／列保存だけ先に確認してもよい
7. マスタ在庫で出す場合のみ `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY=true`（**作業後は false**。常時ONにしない）
8. ~~トグルfalse~~ **不要**（常時ONセット。MASTER_QTY だけ戻す）

### 0b. 本番常時ONセット（2026-08-10）

正本: [CURRENT_PHASE.md](../CURRENT_PHASE.md) §0。

| 未設定時ON | 常時OFF |
|------------|---------|
| `PUT_ENABLED`／`ALLOW_PROD`／`LV4_ENABLED`／`SC_SUMMARY_ENABLED` | `ALLOW_MASTER_QTY`／`P4B_PT_WRITE`／`U2`／`U4` |

既存が false のキーは **削除**か **true に1回**。`clasp push` 後に有効。
相乗り先はN列`ASINコード`のみ。O列／競合URLは使わない。  
送信在庫の既定は **0**（新規GENERATED／相乗りとも）。マスタ選択時のみ「在庫数」生値を GENERATED在庫列および自己発PUTへ（FBAは非送信）。  
開始前確認に **新規もqty内訳**を出す（2026-08-01補強）。 
A1暫定で NF に入った `…af…` は **人手で `Amazon相乗りSKU_FBA` へ移す**（自動移設なし）。  
**マスタ在庫>0の相乗り自己発検収**: [D_MENU_LANE_A3_HUMAN_RUN.md](D_MENU_LANE_A3_HUMAN_RUN.md) — dry `…49a49e`／prod `…f677a3` **OK**（2026-08-01）。

### 1b. 実機合格記録（自己発）

| 日付 | モード | 結果 | runId | 備考 |
|------|--------|------|-------|------|
| 07-30 | dry_run／自己発 | OK=1 | `SPAPI_PUT_OFFER_CK_DRY_20260730_081735_99e` | X=`自己発送`のまま可。SKU=`sanky-B084RJSH7W-48as12` |
| 07-30 | prod／自己発 | OK=1 | `SPAPI_PUT_OFFER_CK_PROD_20260730_082325_9f2287` | NF列再利用。新規=未実行 |
| 07-31 | **新規＋相乗りprod 同時** | 新規=1／相乗りOK=1 | 新規 `LV4_20260731_010959_652276`／相乗り `SPAPI_PUT_OFFER_CK_PROD_20260731_011033_1f63ab` | 同一レ点行。新規=`CK_daba393f8055_B2_GENERATED.csv`（rows=2）／相乗り=`sanky-B01N5A6ESU-19as13` `ACCEPTED issues=0` qty=0 |

**FBA**: [D_MENU_LANE_A1_FBA_HUMAN_RUN.md](D_MENU_LANE_A1_FBA_HUMAN_RUN.md) §0c — dry_run／prod **合格**（2026-08-01）。

### 1b2. デュアル Phase1（系統別列・2026-08-01）

| 日付 | モード | 結果 | runId | 備考 |
|------|--------|------|-------|------|
| 08-01 | dry_run／自己発 | OK=1 | `SPAPI_PUT_OFFER_CK_DRY_20260801_114254_8fa79e` | 保存列=`Amazon相乗りSKU`／`…19as13` |
| 08-01 | dry_run／FBA | OK=1 | `SPAPI_PUT_OFFER_CK_DRY_20260801_114820_d6ed67` | 保存列=`Amazon相乗りSKU_FBA`／`…19af13`・compliance ON |
| 08-01 | **prod／自己発** | OK=1 | `SPAPI_PUT_OFFER_CK_PROD_20260801_115554_4ed30e` | ACCEPTED／行503／NF |
| 08-01 | **prod／FBA** | OK=1 | `SPAPI_PUT_OFFER_CK_PROD_20260801_115648_eb2511` | ACCEPTED／`_FBA` |

詳細: [LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md](LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md) §6.1

### 1b3. デュアル Phase2（両系統1実行・2026-08-01）

| 日付 | モード | 結果 | runId | 備考 |
|------|--------|------|-------|------|
| 08-01 | dry_run／両系統 | OK=1+1 | 自己発 `…141229_f372b8`／FBA `…141243_40d85e` | VALID・行494・`…as19`／`…af19` |
| 08-01 | **prod／両系統** | OK=1+1 | 自己発 `…141420_8fcbc2`／FBA `…141432_cc7c72` | ACCEPTED・ASIN `B0D9VK8YPS` |

- 手順: [D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md](D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md) §3  
- 承認: [LV4_DUAL_OFFER_PHASE2_APPROVAL.md](LV4_DUAL_OFFER_PHASE2_APPROVAL.md) §5

**同時実行の注意**: 新規側が冪等除外（`idempotentBlocked`）で0件になると、**PT/Browse が前回GENERATEDから変わっていれば自動で冪等解除→再GENERATED**（案B）。同じままなら停止（既存CSVでPACKAGED可／強制はレ点「失敗後の再GENERATED」）。[LV4_D_REMAKE_MENU_APPROVAL.md](LV4_D_REMAKE_MENU_APPROVAL.md)

---

## 1c. Dレ点新規で在庫>0のとき

Dレ点新規（`source=child_ck`）は、マスタ在庫>0でも `SKIPPED_IN_STOCK` にせず GENERATED を作る（別カタログのノーブランドセットのため）。マスタ在庫は読取のみで書き換えない。バルクの在庫列は `inventoryMode` 準拠（既定 ZERO＝0／MASTER＝**子の**在庫数生値。Track B 親行は常に0）。MASTER選択時は開始前ダイアログに **[新規GENERATED] qty内訳**が出る。C1／B-T1は GENERATED の `inventory` を xlsm 在庫列へ写す。

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
3. （任意）間隔変更は `…_INTERVAL_MIN`＝5/10/15/30/60。既定15分  

**自動設置・自動削除（2026-07-31 承認）**
- **設置**: 本番 **GENERATED 成功時**（21-①／E-4／D新規）に待ちリストへ `subBatchId` を追加し、監視トリガーを自動設置（ENABLED＋FOLDER必須。DRY_RUNは対象外）
- **削除**: 待ちリストの ID がすべて `UPLOADED_OK`／`UPLOAD_FAILED` になったらトリガー削除。または追加から **72時間**超過で待ちから外し、空なら削除（メール通知）
- 手動の **21-⑯／21-⑰** も残置。21-⑰はトリガー＋待ちリストをクリア

**日常**
- SC UP後、処理サマリ（`{subBatchId}_PACKAGED_…-processing-summary.xlsm`）を監視フォルダへ置くだけ  
- トリガーまたは **21-⑮** で記録 → ファイルは `_処理済` へ  
- 閑散期は待ちが空ならトリガー無し（空走りしない）

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

PACKAGED再送: `…_20260731_032459.xlsm`（KW1枠・粉末・グラム）。  
**再送サマリ確定**（Downloads `…032459-processing-summary.xlsm`）: 処理2／その他エラー2／**100521のみ**（**99016解消**）。  
サブ: U4でPT01〜03取得済 → 手ZIP→Upload Images。  
**台帳**: [LANE_B_SC_ERROR_LEDGER.md](LANE_B_SC_ERROR_LEDGER.md)（手順: [D_MENU_LANE_B_LEDGER_HUMAN_RUN.md](D_MENU_LANE_B_LEDGER_HUMAN_RUN.md)）。

---

## 2. 合格目安

- [x] D新規＝レ点子SKUだけで従来Da
- [x] Dレ点新規が在庫>0でもGENERATEDされ、マスタ在庫が変わらない
- [x] X列が`自己発送`のままでも相乗りdry_runできる
- [x] Dで相乗り自己発を選べる（FBAは未検証）
- [x] 中央が既にASINの子SKUをそのまま使える
- [x] マスタ在庫>0でも quantity=0 でdry_run／prodできる
- [x] N列ASIN空は行スキップ（相乗りのみ／新規同時とも全体は止めない。有効0件のみ停止）
- [ ] **相乗りprod直**（SKU列空でも生成→成功後保存。dry_run不要）… 2026-08-02実装・実機待ち
- [x] D prod＝ALLOW_PROD有り＋OKでACCEPTED（自己発1SKU）
- [x] 確認キャンセルで PUT なし → **A2**（2026-08-01 `cancelled_by_user`）  
- [ ] 21-⑫⑬・E-4 が従来どおり → A2-cでE0到達済。⑫⑬は別途任意  
- [x] 作業後トグル false → **A2-b**（2026-08-01）  
- [x] 相乗りFBA dry_run（2026-08-01・`…111613_41ce9e`／compliance）  
- [x] 相乗りFBA prod（2026-08-01・`…111845_6bd20f` ACCEPTED）  
- [x] 新規＋相乗り同時（2026-07-31・prod）
- [ ] フル → Amazon（楽天OK・**Yahooはrefresh token期限切れで失敗**）
- [ ] SEASONING七味のSC公開（停止理由解消）
- [ ] サブ画像ZIP UP

---

## 3. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-10 | **本番常時ONセット**: PUT／ALLOW_PROD 未設定=ON。トグル毎回切替不要。MASTER_QTYのみ作業後OFF。 |
| 2026-08-05 | **条件付き自動再GENERATED（案B）**: 冪等0件＋PT/Browse差分で自動解除。要 clasp push。 |
| 2026-08-05 | Track B **親行 inventory 常に0**（MASTERでも親マスタ在庫非読取。子のみ厳密）。親 `#DIV/0!` で送信停止しない。 |
| 2026-08-02 | **失敗後再GENERATED＝D内レ点**（案A）。要 clasp push。[LV4_D_REMAKE_MENU_APPROVAL](LV4_D_REMAKE_MENU_APPROVAL.md)。 |
| 2026-08-02 | **相乗りprod直可**: SKU列空でも生成→PUT→成功後保存。dry_runは上級・任意。as/af正規化もprod可。 |
| 2026-08-01 | **相乗り ASIN空 soft skip**: 相乗りのみでも行スキップ・全体継続（有効0件のみ停止）。[D新規ゲート承認 §2.1](LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-08-01 | **デュアル Phase2検収OK**: dry `…f372b8`／`…40d85e`・prod `…8fcbc2`／`…cc7c72`（§1b3）。 |
| 2026-08-01 | **デュアル Phase2コード実装**（§1b3）。実機待ち。 |
| 2026-08-01 | §1f: 七味再送サマリ確定（100521のみ・99016解消）。 |
| 2026-08-01 | §1fからレーンB台帳へリンク。 |
| 2026-08-01 | **D在庫UI補強**: 新規MASTERもqty内訳確認。UIで新規＋相乗り共通を明示。 |
| 2026-08-01 | **A3検収OK**: dry `…49a49e`／prod `…f677a3`（MASTER）。 |
| 2026-08-01 | **A2検収OK**（キャンセル／トグル／E0）。§2一部完了。 |
| 2026-08-01 | §2のキャンセル／トグルを A2 HUMAN_RUN へ誘導。 |
| 2026-08-01 | **デュアル検収OK**: prod 自己発`…4ed30e`／FBA`…eb2511`（§1b2）。 |
| 2026-08-01 | **デュアル dry_run実機OK**: 自己発 `…8fa79e`／FBA `…d6ed67`（§1b2）。prod未。 |
| 2026-08-01 | **デュアル Phase1実装**: 自己発=`Amazon相乗りSKU`／FBA=`Amazon相乗りSKU_FBA`（他系統不変）。実機検収待ち。 |
| 2026-07-31 | §1e: GENERATED時の監視トリガー自動設置／終端・72hで自動削除。 |
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
