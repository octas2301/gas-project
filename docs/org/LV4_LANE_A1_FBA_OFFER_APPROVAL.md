# レーンA1 — 相乗り FBA（dry_run → 少件 prod）承認パッケージ

**日付**: 2026-08-01  
**状態**: **検収OK**（2026-08-01。FBA dry_run／prod 実機合格。属性実装込み）  
**レーン方針**: [LV4_P2_DC_123_INVESTIGATION_APPROVAL.md](LV4_P2_DC_123_INVESTIGATION_APPROVAL.md) **§7.4**  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)  
**親・前提**: [LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)／[LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md](LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)  
**手順**: [D_MENU_LANE_A1_FBA_HUMAN_RUN.md](D_MENU_LANE_A1_FBA_HUMAN_RUN.md)  
**三者レビュー**: **スキップ**

---

## 1. 目的（概要）

レーンAの穴を閉じる。**既存カタログへの相乗りを FBA（`AMAZON_JP`）で**、自己発と同程度まで検収する。

| 段階 | 内容 | 成功定義 |
|------|------|----------|
| **A1-a** | D → 相乗り **FBA** → **dry_run**（1SKU） | VALID／issues 要約。**暫定**: NF空から実行し NF に `…af…` 保存。**二列後**: `Amazon相乗りSKU_FBA` に保存（[デュアル](LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)） |
| **A1-b** | 同SKUで **prod**（在庫経路＝ZERO） | API **ACCEPTED**（または同等の成功）。SCでオファー／発送形態が期待どおり見えること（断定は人間） |
| **A1-c** | （任意）確認キャンセル・トグル戻し | PUT なし／作業後 Property 安全側 |

**既に実装済の前提（コード着手必須ではない）**

- D UI: 相乗り自己発／FBA ラジオ（`offerFulfillment=mfn|fba`）
- PUT: `fulfillmentChannel=AMAZON_JP`、**quantity は API 非送信**、SKU 接尾 **`af`**
- 送信在庫既定 **ZERO**（本A1の prod は **マスタ数量解禁を含めない**＝A3）

本包の主目的は **実機検収＋記録**。静的レビューや実機で不具合が出たときだけ **最小コード修正**を同包で許可する（§3）。

---

## 2. 社長へ確認したい方針（承認時にチェック）

| # | 論点 | 提案（既定案） | 承認 |
|---|------|----------------|------|
| 1 | 次チケット | **A1＝相乗りFBA検収**をレーンAの先頭とする（A3在庫>0・レーンCより先） | **済** |
| 2 | コード | **原則コードなし**。穴があれば `AmazonSpapiPut.js`／UI文言等の**最小差分のみ** | **済** |
| 3 | prod | **1SKU固定**→合格後に必要なら最大 `MAX_ITEMS` まで（既定5）。無人全件禁止 | **済（A1-bまで含む）** |
| 4 | 在庫 | A1では **ZERO のみ**。`ALLOW_MASTER_QTY` は触らない（A3） | **済** |
| 5 | 試験SKU | **FBA納品可能な既存ASIN**1件。自己発合格SKUと別でも可。納品実務が無い場合は dry_run のみ合格＋prodは保留可 | **済**（prodは実施する方針。技術的に不可なら保留可） |
| 6 | 三点 | **スキップ**（契約変更時は再判定） | **済（スキップ）** |
| 7 | 並列 | 七味ライブ確認は **止めない**（運用並行） | **済** |

---

## 3. 変更予定ファイル

### 3.1 承認時点（ドラフト〜方針承認）— docs のみ

| 種別 | パス | 内容 |
|------|------|------|
| **新規** | `docs/org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md` | 本承認包 |
| **新規** | `docs/org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md` | A1専用手順・合格記録 |
| 更新 | `docs/org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md` | §1b に FBA 行・チェックリスト連動 |
| 更新 | `docs/CURRENT_PHASE.md`／`AMAZON_DEV_ROADMAP_P0_P4.md`／`AGENT_HANDOVER.md`／`CHANGE_LEDGER.md` | 次＝A1 |

### 3.2 実装承認後 — 不具合時のみ（コード）

| 種別 | パス | 想定内容（例） |
|------|------|----------------|
| 改修 | `AmazonSpapiPut.js` | FBA body／`af` 正規化／確認文言／ログの穴埋め |
| 改修 | `コード.js` | Dモーダルの FBA 案内・確認ダイアログ文言（必要時） |
| 改修 | `AmazonApprovalExport.js` | 行分類・出荷方式との取り違え防止（必要時のみ） |
| 更新 | 上記 HUMAN_RUN／本承認／LEDGER | 実機結果・復元 |

**コード着手前に再提示**: 変更ファイル一覧／差分概要／リスク → **実装承認**（本包の方針承認とは別チェックでも可）。

### 3.3 やらない（本包）

- マスタ在庫>0 の本番解禁（**A3**）  
- 新規カタログの JSON／xlsm API UP（レーンB/C）  
- FBA 納品プラン作成・Shipment API・在庫転送の自動化  
- 全件マスタループ・日中無人トリガー・承認① AI 再接続  
- 楽天聖域・`Yahoo.js`・B統合 Step 境界  
- P2-③／P3／P4b  

---

## 4. 仕様（検収契約）

### 4.1 属性（既存どおり）

| 項目 | FBA（本A1） | 自己発（参考・済） |
|------|-------------|-------------------|
| `fulfillment_availability.fulfillment_channel_code` | `AMAZON_JP` | `DEFAULT` |
| `quantity` | **送らない** | 送る（ZEROなら0） |
| `Amazon相乗りSKU` 記号 | `…af…` | `…as…` |
| 価格 | `販売価格amazon` | 同 |
| ASIN | N列 `ASINコード` のみ | 同 |

### 4.2 手順の骨子

1. 子行レ点＋N列ASIN。NF列は空可（または記号修正を dry_run で）  
2. `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=true`  
3. D → Amazonのみ → 既存相乗り → **FBA** → **上級 dry_run**  
4. VALID かつ NF に `af` SKU 保存を確認  
5. `ALLOW_PROD=true` → 同選択で **prod** → 確認OK  
6. 全トグル **false** に戻す  

詳細は [D_MENU_LANE_A1_FBA_HUMAN_RUN.md](D_MENU_LANE_A1_FBA_HUMAN_RUN.md)。

### 4.3 Property（既存・新規キーなし）

| キー | A1作業時 | 作業後 |
|------|----------|--------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | true | **false** |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | dry_run後のみ true | **false** |
| `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` | ≤5（初回1推奨） | 戻さなくて可 |
| `APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0` | true のまま（ZERO経路） | — |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY` | **false のまま** | **false** |

---

## 5. 想定リスクと緩和

| リスク | 緩和 |
|--------|------|
| FBAオファー作成後、納品せず在庫0のまま／期待と違う表示 | 試験は1SKU・既知ASIN。SC目視。必要なら自己発に戻す手順を HUMAN_RUN に書く |
| `as` SKU を FBA に誤送信／記号混在 | dry_run 必須。prod は保存済み `af` 必須（既存ロジック）。別 sellerSku として扱う |
| quantity を誤って FBA に送る | 既存 `amazonSpapiPutBuildOfferBody_` が MFNのみ quantity。回帰はログで `fulfillment=fba`＋body確認 |
| 価格誤更新 | 試験価格を事前確認。dry_run で attributes 確認 |
| prod 誤爆・複数行 | MAX_ITEMS・レ点1行・ALLOW_PROD＋確認ダイアログ |
| EC書込（PUT） | 本包は **prod 段階で外部書込あり** → 社長の明示承認必須 |
| FBA未利用アカウント／権限不足 | dry_run 結果とエラーを記録し、prod は保留可（A1-aのみ合格扱い可） |

---

## 6. 検収チェックリスト

- [x] 方針承認（§2）… **2026-08-01**  
- [x] 属性実装承認… [LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md](LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)  
- [x] A1-a dry_run… `SPAPI_PUT_OFFER_CK_DRY_20260801_111613_41ce9e` VALID  
- [x] 対比: 自己発 OK／FBA 90220（属性前）  
- [x] A1-b prod… `SPAPI_PUT_OFFER_CK_PROD_20260801_111845_6bd20f` ACCEPTED  
- [x] HUMAN_RUN／PHASE／HANDOVER／LEDGER 更新  
- [x] 次ゲート: **A2** or **デュアル Phase1** or **A3**  

---

## 7. 復元

1. Property: PUT 系をすべて安全側（§4.3）  
2. コード変更があった場合: 当該 commit を `git revert`  
3. 誤って作った FBA オファー: SC／Listings で人間がクローズまたは自己発へ切替  
4. docs のみなら当該ファイル revert  

---

## 8. 次ゲート（A1完了後）

1. ~~**A2**~~ — docs済・実機は [HUMAN_RUN](D_MENU_LANE_A2_HUMAN_RUN.md)  
2. ~~**デュアル Phase1**~~ — 検収OK  
3. ~~**A3**~~ — dry／prod OK [HUMAN_RUN](D_MENU_LANE_A3_HUMAN_RUN.md)  
4. （並列）七味ライブ／レーンB台帳  

---

## 9. 社長確認

- [x] §2 方針承認… **2026-08-01**  
- [x] prod（A1-b）まで… **含む・実機合格**  
- [x] 三点スキップ  
- [x] 属性実装・実機… **2026-08-01**  

---

## 10. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | ドラフト起草。 |
| 2026-08-01 | **方針承認**。 |
| 2026-08-01 | A1暫定＝NF空＋七味。デュアル方針ロック。 |
| 2026-08-01 | **対比確定**: 自己発VALID／FBA 90220。 |
| 2026-08-01 | **検収OK**: dry_run `…111613_41ce9e`／prod `…111845_6bd20f`。 |
