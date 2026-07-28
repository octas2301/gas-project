# SP-API 出品書込（Listings）— 実装承認パッケージ

**日付**: 2026-07-28  
**状態**: **承認済・ツール実装済**（2026-07-28「③書込承認」）  
**前提**: [D_MENU_SPAPI_SMOKE_HUMAN_RUN.md](D_MENU_SPAPI_SMOKE_HUMAN_RUN.md) 合格済  
**手順**: [D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)  
**当面の本線**: M2＝公式 ListingLoader＋SC手UP（継続）。本ツールは **API書込の後段**。

---

## 1. 目的（v1 提案）

既存 ASIN への **オファー作成／価格・在庫更新** を SP-API（Listings Items）で行う最小実装。

- TRACK=A（M2）相当を API で再現する第一歩  
- **新規カタログ作成（M1／C1）は対象外**  
- SC 自動ブラウザ操作はしない  

---

## 2. 変更予定ファイル（案）

| 種別 | パス | 内容 |
|------|------|------|
| 新規 | `tools/spapi_listings_write/`（または `spapi_smoke` 拡張） | dry_run／prod、PATCH/PUT 1SKU |
| 新規 | `docs/org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md` | 人間手順 |
| 更新 | `docs/CURRENT_PHASE.md`／`AGENT_HANDOVER.md`／`CHANGE_LEDGER.md` | 進捗 |
| （任意・別承認） | GAS Script Properties 読取ラッパ | ローカル成功後 |

**やらない（v1）**

- GAS からの本番一括出品ループ  
- Restricted ロール（発送住所等）  
- 全在庫置換・全件ループ  

---

## 3. 変更概要（v1 スコープ）

1. LWA（既存スモークと同方式）で access_token 取得  
2. Listings Items API で **1 SKU・既存 ASIN** に対し:
   - 価格・数量・コンディション・fulfillment（自己発送）等の最小属性  
3. `dry_run`（リクエスト組み立て＋検証のみ）／`prod`（実 PATCH、要明示フラグ）  
4. 試験SKU既定: 発汗 `lifec-4560151300832-48s11`／`B07YND44VN`／在庫0  

公式参考:

- https://developer-docs.amazon.com/sp-api/docs/listings-items-api  
- https://developer-docs.amazon.com/sp-api/docs/connecting-to-the-selling-partner-api  

---

## 4. 想定リスク

| リスク | 緩和 |
|--------|------|
| 本番出品・価格誤更新 | dry_run 必須／prod は在庫0・1SKU固定／トグル |
| API 属性名・JSONスキーマ相違 | Product Type Definitions 参照・最小属性から |
| 二重出品・SKU衝突 | 既存 M2 合格 SKU のみ／事前 GET |
| 秘密漏洩 | config.local／Script Properties のみ・Git禁止 |
| EC重要変更 | **本承認なしでは実装しない** |

---

## 5. 合格条件（v1）

- [x] dry_run でリクエスト妥当  
- [x] prod 1SKU 成功（発汗）  
- [x] Seller Central でオファー確認  
- [x] ログに runId・SKU・HTTP 結果（秘密なし）  
- [x] HUMAN_RUN 更新  

---

## 6. 社長承認欄

- [x] **承認する**（「③書込承認」・2026-07-28）  
- [ ] 却下／条件付き（条件: ）  

実装: `tools/spapi_listings_write/`（人間が dry_run→prod 実行）。
