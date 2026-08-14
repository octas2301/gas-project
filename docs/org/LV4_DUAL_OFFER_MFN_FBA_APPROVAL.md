# デュアルオファー Phase1 — 自己発／FBA 列分離（承認・要件）

**日付**: 2026-08-01  
**状態**: **検収OK**（dry_run＋prod・両系統）。列範囲チェックは任意・未記録。三点スキップ  
**親**: [LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)／[LV4_LANE_A1_FBA_OFFER_APPROVAL.md](LV4_LANE_A1_FBA_OFFER_APPROVAL.md)／[AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md) §1.1 レーンA  
**三者レビュー**: **スキップ**（2026-08-01・実装承認時。条件＝系統別列のみ書込／自動移設なし／Z非対応）

---

## 1. 目的

同一マスタ行・同一相乗り先ASINで、Amazon上の **自己発オファーと FBA オファーを同時に維持**できるようにする。  
現行の「NF1本＋ラジオ排他＋実行のたびに付け替え」では、片方の sellerSku を失う。

本包は **易しい段階案（Phase1）のみ**。1実行で両系統PUTする Phase2 は含まない。

---

## 2. 社長確定方針（2026-08-01）

| # | 論点 | 決定 |
|---|------|------|
| 1 | 段階 | **易しい段階案（Phase1）** で進める |
| 2 | 新列名 | **`Amazon相乗りSKU_FBA`**（ヘッダ名で解決。物理列番号は固定しない） |
| 3 | 既存 NF | **`Amazon相乗りSKU`＝常に自己発専用** と宣言 |
| 4 | Dの操作 | **1回1系統のまま**（自己発／FBA ラジオ維持）。列だけ分離 |
| 5 | 初回範囲 | **D本線のみ**（Z-21／E互換の二列対応は後追い） |
| 6 | 同時更新 | Phase1では **両方まとめて1実行しない**。維持したい側は列に残し、更新したい側だけ選んで再実行 |
| 7 | A1との関係 | A1実機は当面 **NFを空にしてFBA試験**（二列実装前の暫定）。二列実装後は FBA 保存先を `_FBA` 列へ切替 |

---

## 3. Phase1 仕様

### 3.1 列

| ヘッダ | 系統 | 記号 | 入力者 |
|--------|------|------|--------|
| `Amazon相乗りSKU`（既存NF） | 自己発（MFN） | `…as…` | GAS（prod成功後／任意dry_run VALID後） |
| `Amazon相乗りSKU_FBA`（新規） | FBA | `…af…` | GAS（prod成功後／任意dry_run VALID後） |

- 人間手入力を正にしない  
- 列が無ければ当該系統の相乗りを拒否し、追加手順を案内  
- `子SKU` は変更しない  
- 一括クリア・一括上書き禁止（系統ごと・成功後のみ）  
- 保存済み値の ASIN と N列不一致なら停止  
- 列内の as/af が系統と不一致なら停止（移設案内）

### 3.2 実行契約（D）

```text
D → 相乗り自己発 → prod（通常）／dry_run（上級・任意）
  → 読取・保存先 = Amazon相乗りSKU のみ
  → Amazon相乗りSKU_FBA は触らない

D → 相乗りFBA → prod（通常）／dry_run（上級・任意）
  → 読取・保存先 = Amazon相乗りSKU_FBA のみ
  → Amazon相乗りSKU（自己発）は触らない
```

- **通常運用はprod直**。当該系統の列が空なら **prodでもSKUを生成**し、成功後に保存（dry_run必須にしない）  
- 既存列値の `s`→`as`/`af` 正規化も prod で可（成功後に保存）  
- 確認ダイアログ: 系統・保存列名・SKU例を明示  
- MAX_ITEMS・ZERO／MASTER・ALLOW_* は現行流用（FBAは quantity 非送信）  

### 3.3 含まない（Phase1）

- Dで自己発＋FBAを同時選択して1実行2PUT（**Phase2**）  
- Z-21／承認①経路／E互換の二列対応（初回は Dのみ）  
- 系統別価格  
- FBA納品・Shipment API  
- 90220（電池・危険物）属性のオファーbody拡張 → **別チケット** [LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md](LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)（本包では解けない）  
- 楽天／Yahoo／B統合／C1／画像コースの改変  

### 3.4 移行（実装時）

1. マスタに `Amazon相乗りSKU_FBA` 列を追加（人間作業エリア・ヘッダ解決）  
2. 既存 NF の `…af…` があれば **人手で `_FBA` へ移す**（自動一括移設はしない）  
3. セット構成の列範囲チェックを再検収  
4. HUMAN_RUN・レ点本線承認の「1列」記述を本包に合わせて改定  

---

## 4. 実装内容（2026-08-01 承認・実装）

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `AmazonSpapiPut.js` | `amazonSpapiPutOfferSellerSkuHeader_`／系統別読取・保存・as/af不一致停止 |
| 改修 | `コード.js` | D確認・完了・注記に保存列名 |
| 改修 | `AmazonApprovalExport.js` | コメントのみ（分類ロジック非改変） |
| 更新 | HUMAN_RUN／本ファイル／PHASE／HANDOVER／LEDGER／ROADMAP | 二列契約 |

**やらない**: 楽天聖域、`Yahoo.js`、B統合 Step 境界、Z二列の同時実装。

---

## 5. 想定リスクと緩和

| リスク | 緩和 |
|--------|------|
| FBA実行で自己発SKUを消す | FBAは `_FBA` のみ書込。NF読取禁止 |
| 列未追加で実行 | ヘッダ無なら当該系統を停止＋案内 |
| セット構成の列ずれ | 列追加後に範囲チェック実機 |
| Zが旧1列のまま | 初回はDのみと明記。Zは旧挙動（NFのみ）のまま注意書き |
| NFに残った `…af…` | 自己発実行時に不一致で停止＋移設案内 |
| 食品90220 | A1／属性別チケットで対応済 |

---

## 6. 検収（Phase1）

- [x] 方針ロック（§2）… **2026-08-01 済**  
- [x] 実装承認（§4）… **2026-08-01・三点スキップ**  
- [x] コード実装… **2026-08-01**  
- [x] 自己発 dry_run → NFのみ更新… **`SPAPI_PUT_OFFER_CK_DRY_20260801_114254_8fa79e`** VALID（`…19as13`／行503）  
- [x] FBA dry_run → `_FBA` のみ更新… **`SPAPI_PUT_OFFER_CK_DRY_20260801_114820_d6ed67`** VALID（`…19af13`／`Amazon相乗りSKU_FBA`）  
- [x] 各系統 prod 1SKU（ZERO）… 自己発 **`…115554_4ed30e`**／FBA **`…115648_eb2511`** ACCEPTED  
- [ ] 列範囲チェックOK… **任意・未記録**  
- [x] docs更新  

### 6.1 実機ログ要約（2026-08-01）

| 系統 | mode | runId | SKU | 保存列 |
|------|------|-------|-----|--------|
| 自己発 | dry_run | `…114254_8fa79e` | `sanky-B01N5A6ESU-19as13` | `Amazon相乗りSKU` |
| FBA | dry_run | `…114820_d6ed67` | `sanky-B01N5A6ESU-19af13` | `Amazon相乗りSKU_FBA` |
| 自己発 | prod | `…115554_4ed30e` | 同上 | `Amazon相乗りSKU` |
| FBA | prod | `…115648_eb2511` | 同上 | `Amazon相乗りSKU_FBA` |

FBA側: compliance attrs ON。GET 200 → ACCEPTED。行503。列欠如時は停止→ヘッダ追加後に成功。

---

## 7. Phase2

**別包**: [LV4_DUAL_OFFER_PHASE2_APPROVAL.md](LV4_DUAL_OFFER_PHASE2_APPROVAL.md) — Dで両系統チェック → 1実行で最大2PUT・部分成功。**検収OK（2026-08-01）**。  

---

## 8. 移行（二列実装後の人間手順）

1. マスタヘッダに **`Amazon相乗りSKU_FBA`** を追加  
2. A1暫定で NF に入った `…af…` があれば **カットして `_FBA` へ貼付**（NFは自己発 `…as…` 用に空or自己発SKU）  
3. clasp push 後、自己発／FBA それぞれ dry_run → 保存列が系統どおりか確認  

---

## 9. 社長確認

- [x] §2 確定（列名／NF自己発専用／1回1系統／Dのみ／易しい段階）… **2026-08-01**  
- [x] 実装着手の承認… **2026-08-01**（三点スキップ）  
- [x] dry_run 実機… **2026-08-01**（§6.1）  
- [x] prod 実機… **2026-08-01**（`…4ed30e`／`…eb2511`）  

---

## 10. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | **通常運用=prod直**: 相乗りSKU列空でもprodで生成・成功後保存。dry_runは上級・任意。[D_ENTRY](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。 |
| 2026-08-01 | 起草。社長方針ロック（易しい段階案）。実装未。 |
| 2026-08-01 | **実装承認＋実装**（三点スキップ）。実機検収待ち。 |
| 2026-08-01 | **dry_run 実機OK**: 自己発 `…8fa79e`／FBA `…d6ed67`。prod未提示。 |
| 2026-08-01 | **prod 検収OK**: 自己発 `…4ed30e`／FBA `…eb2511`。Phase1完了。 |
