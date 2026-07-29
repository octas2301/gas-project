# SP-API × D入口 — Amazon 新規／既存相乗りの選択（承認パッケージ・起草）

**日付**: 2026-07-29  
**状態**: **承認済・実装待ち**（2026-07-29 社長承認。コードは未着手）  
**テーマ名**: **A（D入口版）** … 本線 D を Amazon 出品の起点にし、新規カタログと既存ASIN相乗りを人間が選ぶ  
**前提合格**: SP-API v1.4 第1段・第2段 API 実機合格／D×Amazon U3（Da）実機合格／E コース実装済  
**関連**:  
- [LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)（21-⑫⑬）  
- [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)（D 本線）  
- [D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)（一時 E・テスト期間残可）  
- [D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)  
**三者レビュー**: **不要**（組織・承認マトリクス改定ではない。D ラジオ拡張＋既存 PUT 呼出の薄いファサード）

---

## 1. 目的

人間が **D（一括出品）を唯一の本線起点**として、Amazon について次を選ぶ。

| 人間の選択（ラジオ文言案） | 中身 | TRACK 対応（括弧のみ） |
|----------------------------|------|------------------------|
| Amazonのみ・**新規カタログ**（Da＝GENERATED→PACKAGED→SC） | 現行 `amazon`／`menuApprovalAmazonLv4Run`（21-①） | （B＝M1） |
| Amazonのみ・**既存に相乗り**（SP-API **dry_run**） | 既存 `menuAmazonSpapiPutApprovedDryRun`（21-⑫）呼出 | （A＝M2） |
| Amazonのみ・**既存に相乗り**（SP-API **prod**＝実反映） | 既存 `menuAmazonSpapiPutApprovedProd`（21-⑬）呼出 | （A＝M2） |
| **フル → Amazon 新規カタログ**（Da・SC待ちで止まる） | 現行 `full_amazon` | （B＝M1） |

**方針（社長確認済）**

- **本線の起点は D**。楽天／Yahoo と同様に Amazon も D から選ぶ。  
- **D からも prod 可**。ただし `ALLOW_PROD`＋確認ダイアログ必須（21-⑬ と同等の硬さ）。  
- **テスト期間は E・Z-21（⑩〜⑬）が増えても許容**。逃げ道として残す。本線定着後の E 縮小は別承認。  
- 承認②・在庫>0無人・新規カタログの SP-API 作成は **対象外**。

用語はラジオ上で **「新規」／「既存に相乗り」** を使い、TRACK の A／B 記号は実装・ログの括弧対応に留める（取り違え防止）。

---

## 2. 変更予定ファイル（実装は承認後）

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `コード.js` | `showBatchExportModal` のラジオ拡張／`runBatchExportAmazonFacade` に経路分岐 |
| （必要なら）改修 | `AmazonSpapiPut.js` | D から呼ぶ際の `silent`／ダイアログ方針の微調整のみ。抽出・PUT 本体は再利用 |
| 新規 | 本ファイル | 承認 |
| 更新（実装後） | HUMAN_RUN（D または SP-API）／CURRENT_PHASE／HANDOVER／CHANGE_LEDGER／E HUMAN_RUN 1行案内 | 進捗 |

**やらない（本承認の範囲外）**

- E-4 を既存相乗りに置換する改変（E は残す）  
- `フル → Amazon 既存相乗り prod`（楽天・Yahoo 後に本番 PUT まで一気通貫）  
- Drive `02` MAIN ゲートを既存相乗り経路に適用すること  
- 承認②・販売中SKU無人上書き・Cloud Agent からの本番 PUT  
- 楽天聖域・`Yahoo.js`・B統合 Step 境界の改変  
- T3／ε／Dc の実装  

---

## 3. 仕様（案）

### 3.1 D ラジオ（確定案）

```text
○ フル（楽天 → Yahoo!）           … 既存のまま
○ 楽天のみ                         … 既存のまま
○ Yahoo!のみ                       … 既存のまま
○ Amazonのみ・新規カタログ（Da）
○ Amazonのみ・既存に相乗り（dry_run）
○ Amazonのみ・既存に相乗り（prod）
○ フル → Amazon 新規カタログ（Da・SC待ち）
```

コース値（実装時の内部名・案）:

| value | 挙動 |
|-------|------|
| `amazon` または `amazon_new` | 現行 Da（21-①）。Drive `02` MAIN ゲート **あり** |
| `amazon_offer_dry` | 21-⑫ 相当。`02` ゲート **なし**。即時実行 |
| `amazon_offer_prod` | 21-⑬ 相当。`02` ゲート **なし**。即時実行。§3.3 のゲート必須 |
| `full_amazon` | 現行どおり **新規カタログ Da のみ**（既存相乗りは載せない） |

### 3.2 既存相乗り経路の契約

- 対象抽出・PUT は **v1.4 第2段と同一**（承認①済・子SKU必須・親行スキップ・LISTING_OFFER_ONLY・FORCE_QTY_0 既定 true）  
- D は **薄いファサード**（dispatch＋完了メッセージ）。Listings 本体を D に複製しない  
- Amazon 系は **トリガー裏実行禁止**（現行 U3 どおり即時）  
- dry_run／prod の Script Properties は第1・第2段と **共用**（新規キー必須ではない。必要なら `AMAZON_D_SPAPI_ENTRY_ENABLED` を任意追加可＝テスト期間の空き枠利用）

### 3.3 prod ゲート（D から押す場合も同じ）

すべて満たすこと。1つでも欠ければ拒否（PUTしない）。

1. `APPROVAL_AMAZON_SPAPI_PUT_ENABLED=true`  
2. `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD=true`  
3. 確認ダイアログで **OK**（キャンセル＝`cancelled_by_user`・PUTなし）  
4. `max_items` 以内・FORCE_QTY_0 既定  

ダイアログには少なくとも `batchId`／件数／SKU例／在庫FORCE_0 を出す（21-⑬ 相当）。

### 3.4 明示禁止の組み合わせ

| 組み合わせ | 扱い |
|------------|------|
| フル（楽天→Yahoo）＋既存相乗り prod | **禁止**（ラジオに出さない） |
| フル＋既存相乗り dry_run | 当面 **出さない**（必要なら別承認） |
| 既存相乗り × Drive `02` MAIN ゲート | **適用しない**（新規専用） |

### 3.5 E／Z との関係（テスト期間）

- **E-0〜E-5**・**21-⑩〜⑬** は残す（段階テスト・復旧用）  
- E-0 の案内に「本線は D。新規＝Da／既存相乗り＝D のラジオまたは 21-⑫⑬」を **1行追加**してよい（実装時）  
- E 削除・統合は本承認に含めない  

### 3.6 ログ

- `runId`／`course`／`source=approved|new`／`state`／`batchId`（相乗り時）  
- トークン・secret は出さない  

---

## 4. 想定リスク

| リスク | 緩和 |
|--------|------|
| 新規と既存相乗りの取り違え | ラジオ文言を「新規カタログ」「既存に相乗り」に固定。TRACK記号はUIに出さない |
| D から誤って本番 PUT | ALLOW_PROD 既定 false＋確認ダイアログ＋max_items |
| `full_amazon` 後に本番 PUT | ラジオに載せない（§3.4） |
| 既存経路で `02` ゲートに誤停止 | 相乗り経路ではゲートを呼ばない |
| E／Z と D の二重運用で混乱 | E-0／HUMAN_RUN に「本線＝D」を明記。テスト期間は増メニュー許容 |
| EC重要変更 | **本承認なしでは実装しない** |

---

## 5. 合格条件（実装後）

- [ ] D で「新規カタログ」を選ぶと現行 Da（21-①）相当が動き、`02` ゲートが効く  
- [ ] D で「既存相乗り dry_run」→ VALID／`batchId` 表示（`02` ゲートなし）  
- [ ] D で「既存相乗り prod」→ ALLOW_PROD 無しで拒否／有り＋OK で ACCEPTED  
- [ ] 確認キャンセルで PUT されない（`cancelled_by_user`）  
- [ ] `full_amazon` は新規 Da のみ（相乗り prod が混ざらない）  
- [ ] 21-⑫⑬・E-4 が従来どおり動く  
- [ ] 作業後トグル false  
- [ ] HUMAN_RUN／CURRENT_PHASE 更新  

試験SKU候補: 既存相乗りは v1.4 第2段合格の `lifec-4560151300832-48s11`／`B07YND44VN`（在庫0）。

---

## 6. 社長承認欄

- [x] **承認する**（A＝D入口版・ラジオ案確定・Dから prod 可＝ALLOW_PROD＋確認必須・E/Z はテスト期間残可・コード実装可／2026-07-29）  
- [ ] 却下／条件付き（条件: ）

**実装開始条件**: 本 §6 の承認後のみ。承認前のコード追加は禁止 → **2026-07-29 承認済。実装可**（範囲は §2・§3。§2「やらない」と §3.4 の禁止組み合わせは厳守）。
