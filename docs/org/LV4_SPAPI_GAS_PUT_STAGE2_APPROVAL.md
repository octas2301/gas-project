# SP-API v1.4 第2段 — 承認①済 Amazon を GAS から直 PUT（承認パッケージ・起草）

**日付**: 2026-07-29  
**状態**: **承認済・実装待ち**（2026-07-29 社長承認。コードは未着手）  
**親承認**: [LV4_SPAPI_GAS_PUT_APPROVAL.md](LV4_SPAPI_GAS_PUT_APPROVAL.md)（v1.4 第1段＝子レ点・API実機合格）  
**関連**: [LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md](LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)（v1.2b＝21-⑨ Drive CSV・実機合格）  
**手順（実装後に追記）**: [D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)  
**三者レビュー**: **不要**（組織・承認マトリクス改定ではなく、既存 v1.4 への小さな対象追加）

---

## 1. 目的

現行 **21-⑨（承認①済 Amazon → Drive CSV）** と同じ対象抽出で、CSV／ローカル Python を介さず **GAS から直接 Listings PUT** する。

- 属性範囲は第1段と同一: **LISTING_OFFER_ONLY**（価格・在庫・状態・自己発送）
- **新規カタログ作成なし**
- 楽天聖域・`Yahoo.js`・B統合 Step 境界は触らない

第1段（子SKUレ点＝21-⑩⑪）と 21-⑧／⑨＋`--fetch-drive` 経路は **すべて残す**。

---

## 2. 変更予定ファイル（実装は承認後）

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `AmazonSpapiPut.js` | 承認①済からの候補収集を追加（既存 LWA／PUT 部は再利用） |
| 改修 | `コード.js` | メニュー例: **21-⑫ 承認①済 dry_run**／**21-⑬ 承認①済 prod**（番号・文言は実装時確定） |
| 新規 | 本ファイル | 承認 |
| 更新 | GAS PUT HUMAN_RUN／CURRENT_PHASE／HANDOVER／CHANGE_LEDGER | 進捗 |

**やらない（第2段）**

- 全件マスタループ・在庫>0 無人出品  
- 親行のみ（子SKU空）の出品  
- Restricted ロール（発送住所等）  
- 承認②（在庫反映）との自動連結  
- Cloud Agent からの本番 PUT  
- 第1段メニュー（21-⑩⑪）の挙動変更  

---

## 3. 仕様（案）

### 3.1 対象抽出

21-⑨（`menuAmazonSpapiExportApprovedItemsCsv`）と **同一ロジックを流用**する。

- 最新 APPROVED バッチの `mall=amazon` かつ `lineStatus=APPROVED`
- **子SKU必須**。親行のみ（子SKU空）は **スキップ**（理由をログ・ダイアログに出す）
- 同一子SKUは重複排除
- マスタ行が見つからない行はスキップ
- ASIN: `ASINコード` → `競合店ASIN` → URL（既存ヘルパ流用）
- 価格: `販売価格amazon`（行→親フォールバック）
- 在庫: **FORCE_QTY_0 既定 true**
- `note` に `approved=<batchId>` を残す（21-⑨と同じ）

### 3.2 メニュー／モード

| モード | 挙動 |
|--------|------|
| dry_run | `mode=VALIDATION_PREVIEW`（永続化しない） |
| prod | 実 PUT。`ALLOW_PROD`＋確認ダイアログ必須 |

### 3.3 Script Properties

**第1段と同じキーを共用**（新規キーを増やさない＝UI 50件制限への配慮）。

| キー | 既定 | 備考 |
|------|------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | **false** | 主トグル |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | **false** | prod 許可 |
| `APPROVAL_AMAZON_SPAPI_PUT_MAX_ITEMS` | `5` | 超過は拒否 |
| `APPROVAL_AMAZON_SPAPI_PUT_FORCE_QTY_0` | **true** | 在庫0強制 |
| LWA 3点／`SPAPI_SELLER_ID`／marketplace／endpoint | — | 第1段と共用。`SPAPI_ENDPOINT` は空推奨 |

CSV 出力用 `APPROVAL_AMAZON_SPAPI_EXPORT_*` とは **分離**を維持。

### 3.4 ログ

- `runId`／`stepName`／`functionName`／`state`／`batchId`／件数／SKU要約／HTTP要約
- 親スキップ件数を明示
- **トークン・client_secret はログに出さない**

---

## 4. 想定リスク

| リスク | 緩和 |
|--------|------|
| 承認①バッチの取り違えで意図しないSKUへ PUT | dry_run 必須運用・`batchId` をログとダイアログに表示・`max_items` |
| 第1段（子レ点）との混同 | メニュー文言に「承認①済」を明記・ログの stepName を分離 |
| 誤クリックで本番 PUT | 主トグル＋ALLOW_PROD 両方 false 既定＋確認ダイアログ |
| 在庫誤爆 | FORCE_QTY_0 既定 true |
| 実行時間・レート | 1件ずつ UrlFetch・`max_items` 小・失敗時の停止方針を HUMAN_RUN に明記 |
| EC重要変更 | **本承認なしでは実装しない** |

---

## 5. 合格条件（実装後）

- [ ] Property OFF でメニューが拒否する  
- [ ] 承認①済が無いとき明示メッセージで停止  
- [ ] dry_run: VALID／issues=0（`batchId` がログに出る）  
- [ ] prod: ACCEPTED（ALLOW_PROD=true・在庫0）  
- [ ] 親行のみはスキップされ件数が表示される  
- [ ] 第1段 21-⑩⑪ が従来どおり動く（回帰なし）  
- [ ] 作業後トグル false  
- [ ] HUMAN_RUN／CURRENT_PHASE 更新  

試験SKU候補: v1.2b で合格済の `lifec-4560151300832-48s11`／`B07YND44VN`（在庫0）。実装時に HUMAN_RUN で固定。

---

## 6. 社長承認欄

- [x] **承認する**（v1.4 第2段＝承認①済 Amazon の GAS 直 PUT・コード実装可／2026-07-29）  
- [ ] 却下／条件付き（条件: ）

**実装開始条件**: 本 §6 の承認後のみ。承認前のコード追加は禁止 → **2026-07-29 承認済。実装可**。
