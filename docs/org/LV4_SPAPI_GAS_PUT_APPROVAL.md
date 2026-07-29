# SP-API v1.4 — GAS から Listings 直呼び（実装承認パッケージ・起草）

**日付**: 2026-07-29  
**状態**: **承認済・実装済・実機合格（API）**（2026-07-29）。SC最終更新の目視は **反映待ち**。  
**前提**: ローカル Listings v1／v1.1 実機合格／橋渡し v1.2〜v1.3／v1.2b／v1.2c 実機合格  
**手順**: [D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)  
**三者レビュー**: **不要**（組織・マトリクス改定ではない。従来の SP-API 実装承認と同型。社長1承認で可）

---

## 1. 目的

スプシメニューから **GAS（UrlFetchApp）で SP-API Listings Items PUT** を行い、ローカル Python／Drive CSV を省略できるようにする。

- 属性範囲は既存ローカルと同一: **LISTING_OFFER_ONLY**（価格・在庫・状態・自己発送）
- **新規カタログ作成なし**
- 楽天／Yahoo／B統合境界は触らない

---

## 2. 変更予定ファイル（案・実装は承認後）

| 種別 | パス | 内容 |
|------|------|------|
| 新規 | `AmazonSpapiPut.js` | LWA＋GET/PUT Listings・dry_run／prod（**実装済**） |
| 改修 | `コード.js` | **21-⑩ dry_run**／**21-⑪ prod**（**実装済**） |
| 新規 | 本ファイル／HUMAN_RUN | 承認・手順 |
| 更新 | CURRENT_PHASE／HANDOVER／CHANGE_LEDGER | 進捗 |

**やらない（v1.4）**

- 全件マスタループ・在庫>0無人出品  
- Restricted ロール（発送住所等）  
- 親レ点のみで全子展開  
- 楽天聖域・Yahoo.js・B統合 Step 境界の改変  
- Cloud Agent からの本番 PUT  

---

## 3. 仕様（第1段＝子SKUレ点）

### 3.1 対象行

- マスタの **子SKUあり かつ 出品CK** のみ（v1.2c と同じ）
- 親行のみは出さない
- ASIN: `ASINコード` → `競合店ASIN` → URL（既存ヘルパ流用）
- 価格: `販売価格amazon`（行→親）
- 在庫: **FORCE_QTY_0 既定 true**

### 3.2 メニュー／モード

| モード | 挙動 |
|--------|------|
| dry_run | `mode=VALIDATION_PREVIEW`（ローカル dry_run 相当） |
| prod | 実 PUT。別 Property または確認ダイアログ必須 |

第2段（任意・別小さな承認でも可）: 承認①済 Amazon（現行 21-⑨相当）を直呼び対象に追加。

### 3.3 Script Properties（案）

| キー | 既定 | 内容 |
|------|------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | **false** | メニュー全体の主トグル |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | **false** | prod 許可（dry_run のみなら false のまま） |
| `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` | `5` | 1〜50。超過は拒否 |
| `APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0` | **true** | 在庫0強制 |
| `SPAPI_LWA_CLIENT_ID` / `SPAPI_LWA_CLIENT_SECRET` / `SPAPI_REFRESH_TOKEN` | （秘密） | ローカル smoke と同系統。**Git・チャット禁止** |
| `SPAPI_SELLER_ID` / `SPAPI_MARKETPLACE_ID` | （設定） | 出品者トークン／marketplace |
| `SPAPI_ENDPOINT` | 空推奨 | `https:\` 誤記は `amazonSpapiPutNormalizeEndpoint_` で矯正 |

既存 CSV 出力用 `APPROVAL_AMAZON_SPAPI_EXPORT_*` とは **分離**（誤爆防止）。

### 3.4 ログ

- `runId`／`stepName`／`functionName`／`state`／件数／SKU要約／HTTP 要約
- **トークン・client_secret はログに出さない**

### 3.5 互換

- 21-⑧／⑨＋ローカル `--fetch-drive` 経路は **残す**（フォールバック）

---

## 4. 想定リスク

| リスク | 緩和 |
|--------|------|
| 誤クリックで本番 PUT | 主トグル＋ALLOW_PROD 両方 false 既定／確認ダイアログ／max_items |
| 秘密漏洩 | Script Properties のみ。コード・docs・Git に書かない |
| 実行時間・レート | max_items 小／1件ずつ UrlFetch／失敗で停止方針を HUMAN_RUN に明記 |
| 属性スキーマ差異 | ローカル合格 JSON と同一組み立てを移植 |
| EC重要変更 | **本承認なしでは実装しない** |

---

## 5. 合格条件（実装後）

- [x] Property OFF ではメニューが拒否する  
- [x] dry_run: VALID／issues=0（試験SKU `…48s11`）  
- [x] prod: ACCEPTED（ALLOW_PROD=true・在庫0）  
- [ ] SC 最終更新日が prod 以降に進んだ目視（**反映待ち**・既存出品は確認済）  
- [x] 作業後トグル false  
- [x] HUMAN_RUN／CURRENT_PHASE 更新  

試験SKU: `lifec-4560151300924-48s11`／`B00A0J0D30`。prod runId `SPAPI_PUT_PROD_20260729_223506_190ff7`。

---

## 6. 社長承認欄

- [x] **承認する**（v1.4・第1段＝子レ点＋GAS UrlFetch Listings・コード実装可／2026-07-29）  
- [ ] 却下／条件付き（条件: ）

**実装**: 2026-07-29 `AmazonSpapiPut.js`＋21-⑩⑪。実機 API 合格。ENDPOINT 正規化・host ヘッダ削除込み。
