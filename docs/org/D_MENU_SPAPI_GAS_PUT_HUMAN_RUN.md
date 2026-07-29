# SP-API v1.4 GAS 直呼び（人間手順）

**状態**: **実機合格（API）**／SC最終更新の目視は **反映待ち注記**（2026-07-29）  
**承認**: [LV4_SPAPI_GAS_PUT_APPROVAL.md](LV4_SPAPI_GAS_PUT_APPROVAL.md)（第1段）／[LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)（第2段・**承認待ち＝未実装**）  
**前提合格**: ローカル Listings／橋渡し v1.2〜v1.3／v1.2b／v1.2c  
**範囲**: 既存 ASIN・LISTING_OFFER_ONLY。新規カタログなし。第1段＝**子SKUレ点**のみ。

---

## 0. できること／できないこと

| できる | できない |
|--------|----------|
| **21-⑩** dry_run（VALIDATION_PREVIEW） | 全件・在庫>0無人・親レ点のみ全子 |
| **21-⑪** prod（ALLOW_PROD 時のみ） | 承認①一括直呼び（第2段・**承認待ち／メニュー未実装**） |
| 子SKU＋出品CK のみ | 楽天／Yahoo 改変 |

ローカル経路（21-⑧／⑨＋`--fetch-drive`）は **併用可**。

---

## 1. Script Properties

| キー | 既定 | 内容 |
|------|------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | **false** | 主トグル |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | **false** | prod 許可 |
| `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` | `5` | 上限 |
| `APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0` | **true** | 在庫0 |
| `SPAPI_LWA_CLIENT_ID` | — | LWA（秘密・Git禁止） |
| `SPAPI_LWA_CLIENT_SECRET` | — | LWA |
| `SPAPI_REFRESH_TOKEN` | — | LWA |
| `SPAPI_SELLER_ID` | — | 出品者トークン |
| `SPAPI_MARKETPLACE_ID` | 空→`A1VC38T7YXB528` | JP |
| `SPAPI_ENDPOINT` | **空推奨**→FE endpoint | `https:\` 誤記はコードが `https://` に正規化。空なら既定 |

ローカル `tools/spapi_smoke/config.local.json` の値を Properties に転記（チャットに貼らない）。  
`SPAPI_ENDPOINT` は **空のまま**が安全（既定URL使用）。手入力する場合は必ず `https://`。

---

## 2. 手順

1. `clasp push`（`AmazonSpapiPut.js`＋`コード.js`）  
2. §1 の Properties を設定（ENDPOINT は空推奨）  
3. `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=true`  
4. マスタで **子行に出品CK**（親のみ不可）  
5. **21-⑩** dry_run → VALID／issues=0  
6. `ALLOW_PROD=true` → **21-⑪** prod（確認ダイアログ）→ SC 確認  
7. `PUT_ENABLED`／`ALLOW_PROD` を **false** に戻す  

---

## 3. 合格目安

- [x] Property OFF で拒否（運用上確認）  
- [x] 21-⑩ dry_run VALID（`…223012`／`…223300`）  
- [x] 21-⑪ prod ACCEPTED（`…223506`・issues=0）  
- [x] トグル戻し（作業後）  
- [ ] SC **最終更新日**が prod 時刻以降に進んだことの目視（**反映待ち**）  

### 3.1 実機記録（2026-07-29）

| 項目 | 結果 |
|------|------|
| SKU／ASIN | `lifec-4560151300924-48s11`／`B00A0J0D30` |
| dry_run | VALID・ok=1（例: `SPAPI_PUT_DRY_20260729_223300_91d3dc`） |
| prod | ACCEPTED・ok=1（`SPAPI_PUT_PROD_20260729_223506_190ff7`） |
| 障害メモ | 初回 `SPAPI_ENDPOINT=https:\…` で UrlFetch「無効な引数」。Property削除＋コード正規化で解消 |
| SC | 既存出品は見える。prod前の最終更新 `01:04` からの更新確認は **反映待ち** |

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-29 | **実機合格（API）** dry_run／prod。ENDPOINT `https:\` 正規化。SC最終更新は反映待ち注記。 |
| 2026-07-29 | UrlFetch の `host` ヘッダー削除（無効な引数対策）。 |
| 2026-07-29 | 実装: `AmazonSpapiPut.js`／21-⑩⑪。 |
| 2026-07-29 | 起草。 |
