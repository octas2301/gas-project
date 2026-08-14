# SP-API v1.4 GAS 直呼び（人間手順）

**状態**: 第1段・第2段とも **実機合格（API）**（2026-07-29）。SC最終更新の目視は **反映待ち**  
**承認**: [LV4_SPAPI_GAS_PUT_APPROVAL.md](LV4_SPAPI_GAS_PUT_APPROVAL.md)（第1段）／[LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)（第2段・**承認済・実装済**）  
**前提合格**: ローカル Listings／橋渡し v1.2〜v1.3／v1.2b／v1.2c  
**範囲**: 既存 ASIN・LISTING_OFFER_ONLY。新規カタログなし。

---

## 0. できること／できないこと

| できる | できない |
|--------|----------|
| **21-⑩** dry_run（子SKUレ点） | 全件・在庫>0無人・親レ点のみ全子 |
| **21-⑪** prod（子SKUレ点・ALLOW_PROD） | Restricted ロール・新規カタログ |
| **21-⑫** dry_run（承認①済 Amazon） | 承認②との自動連結 |
| **21-⑬** prod（承認①済・ALLOW_PROD） | Cloud Agent からの本番 PUT |
| 子SKU＋出品CK／承認①済子SKU | 楽天／Yahoo 改変 |

ローカル経路（21-⑧／⑨＋`--fetch-drive`）は **併用可**。第1段と第2段は **同じ Property を共用**。  
**本線入口**: D ラジオでも既存相乗りを選べる（[D入口 HUMAN_RUN](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)／[承認](LV4_SPAPI_D_ENTRY_APPROVAL.md)）。

---

## 1. Script Properties

| キー | 既定 | 内容 |
|------|------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | **true（未設定）** | 主スイッチ。明示 false で緊急停止 |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | **true（未設定）** | prod 許可。開始前確認は残る |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY` | **false** | マスタ在庫送信のみ。常時ONにしない |
| `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` | `10`（未設定時） | 1〜50。Property省略可 |
| `APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0` | **true** | 在庫0 |
| `SPAPI_LWA_CLIENT_ID` | — | LWA（秘密・Git禁止） |
| `SPAPI_LWA_CLIENT_SECRET` | — | LWA |
| `SPAPI_REFRESH_TOKEN` | — | LWA |
| `SPAPI_SELLER_ID` | — | 出品者トークン |
| `SPAPI_MARKETPLACE_ID` | 空→`A1VC38T7YXB528` | JP |
| `SPAPI_ENDPOINT` | **空推奨**→FE endpoint | `https:\` 誤記はコードが `https://` に正規化。空なら既定 |

**2026-08-10**: 本番常時ONセット（[CURRENT_PHASE](../CURRENT_PHASE.md) §0）。旧 false キーは削除 or true に1回。  
ローカル `tools/spapi_smoke/config.local.json` の値を Properties に転記（チャットに貼らない）。  
`SPAPI_ENDPOINT` は **空のまま**が安全（既定URL使用）。手入力する場合は必ず `https://`。

---

## 2. 手順（第1段＝子SKUレ点）

1. `clasp push`（`AmazonSpapiPut.js`＋`コード.js`）  
2. §1 の Properties を設定（ENDPOINT は空推奨）  
3. `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=true`  
4. マスタで **子行に出品CK**（親のみ不可）  
5. **21-⑩** dry_run → VALID／issues=0  
6. `ALLOW_PROD=true` → **21-⑪** prod（確認ダイアログ）→ SC 確認  
7. `PUT_ENABLED`／`ALLOW_PROD` を **false** に戻す  

---

## 2b. 手順（第2段＝承認①済）

1. `clasp push`（第2段メニュー 21-⑫⑬ を含む版）  
2. 承認キューに **APPROVED の Amazon 明細**があること（無いと明示停止）  
3. `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=true`  
4. **21-⑫** dry_run → ダイアログ／ログに **batchId** が出ること・VALID／issues=0  
5. `ALLOW_PROD=true` → **21-⑬** prod（確認に batchId）→ SC 確認  
6. 親行のみ（子SKU空）はスキップ件数として表示されること  
7. `PUT_ENABLED`／`ALLOW_PROD` を **false** に戻す  

**試験SKU候補**: `lifec-4560151300832-48s11`／`B07YND44VN`（v1.2b 合格・在庫0）。最新 APPROVED バッチに含まれること。

**失敗時**: 1件失敗しても残りは続行（ログに FAIL）。レートは 300ms sleep。`max_items` 超過は拒否（PUTしない）。

---

## 3. 合格目安（第1段）

- [x] Property OFF で拒否（運用上確認）  
- [x] 21-⑩ dry_run VALID（`…223012`／`…223300`）  
- [x] 21-⑪ prod ACCEPTED（`…223506`・issues=0）  
- [x] トグル戻し（作業後）  
- [ ] SC **最終更新日**が prod 時刻以降に進んだことの目視（**反映待ち**）  

### 3.1 実機記録・第1段（2026-07-29）

| 項目 | 結果 |
|------|------|
| SKU／ASIN | `lifec-4560151300924-48s11`／`B00A0J0D30` |
| dry_run | VALID・ok=1（例: `SPAPI_PUT_DRY_20260729_223300_91d3dc`） |
| prod | ACCEPTED・ok=1（`SPAPI_PUT_PROD_20260729_223506_190ff7`） |
| 障害メモ | 初回 `SPAPI_ENDPOINT=https:\…` で UrlFetch「無効な引数」。Property削除＋コード正規化で解消 |
| SC | 既存出品は見える。prod前の最終更新 `01:04` からの更新確認は **反映待ち** |

### 3.2 合格目安（第2段・**実機合格 2026-07-29**）

- [x] 21-⑫ dry_run VALID（batchId がログ・ダイアログに出る）  
- [x] 21-⑬ prod ACCEPTED（ALLOW_PROD・在庫0）  
- [x] 親行スキップ件数が表示される（1件）  
- [x] 21-⑩⑪ は未変更（第1段の runId 体系そのまま・回帰なし）  
- [x] 作業後トグル false  
- [ ] Property OFF での拒否は 21-⑩⑪ で確認済（21-⑫⑬ は同一関数のため未再現）  
- [ ] 承認①済なしの明示停止は未再現（APPROVED が存在したため）  

### 3.3 実機記録・第2段（2026-07-29）

| 項目 | 結果 |
|------|------|
| batchId | `A1_20260727_224939_b7a053` |
| SKU／ASIN | `lifec-4560151300832-48s11`／`B07YND44VN`（在庫0） |
| dry_run | **VALID**・ok=1／issues=0（`SPAPI_PUT_APPR_DRY_20260729_231605_1b24e2`） |
| prod | **ACCEPTED**・ok=1／issues=0（`SPAPI_PUT_APPR_PROD_20260729_232041_3d83f3`） |
| 親行スキップ | 1（子SKU空の親行。想定どおり） |
| メモ | 1回目の 21-⑬（`…231747_be4da5`）は確認ダイアログで **キャンセル**＝`state=FAILED cancelled_by_user`。PUTは未実行。再実行で OK を押して合格 |
| SC | 最終更新の目視は第1段と同様に **反映待ち** |

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-29 | **第2段実機合格**: 21-⑫ VALID／21-⑬ ACCEPTED（`…0832-48s11`・batch `…b7a053`）。 |
| 2026-07-29 | **第2段実装**: 21-⑫⑬／承認①済抽出（21-⑨相当）。 |
| 2026-07-29 | **実機合格（API）** dry_run／prod。ENDPOINT `https:\` 正規化。SC最終更新は反映待ち注記。 |
| 2026-07-29 | UrlFetch の `host` ヘッダー削除（無効な引数対策）。 |
| 2026-07-29 | 実装: `AmazonSpapiPut.js`／21-⑩⑪。 |
| 2026-07-29 | 起草。 |
