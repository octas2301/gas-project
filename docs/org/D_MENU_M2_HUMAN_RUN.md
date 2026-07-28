# M2 — TRACK=A 既存ASINオファー（人間手順）

**状態**: **実機合格（発汗・2026-07-28）**／v1実装済  
**正本**: [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §8・§12.4  
**ギャップ**: [LV4_M2_TRACK_A_GAP_ANALYSIS.md](LV4_M2_TRACK_A_GAP_ANALYSIS.md)  
**承認**: [LV4_M2_IMPLEMENTATION_APPROVAL.md](LV4_M2_IMPLEMENTATION_APPROVAL.md)  
**PACKAGED**: `tools/m2_offer_packaged/`（案L CSV）＋ **`m2_listing_loader_fill.py`（公式 ListingLoader xlsm 埋め）**

---

## 0. 試験SKU（発汗チェッカー）— 合格実績

| 項目 | 値 |
|------|-----|
| 親SKU | `lifec-4560151300832-oya` |
| 子SKU | `lifec-4560151300832-48s11` |
| ASIN | `B07YND44VN` |
| 価格／在庫（試験） | `981`／`0` |
| subBatchId | `A1_20260727_224939_b7a053_A2` |
| SC結果 | processing-summary: 処理1／成功1／エラー0（`…_M2_ListingLoader_FILLED-processing-summary.xlsm`） |
| E-5 | **済**（2026-07-28）／Property・在庫片付け **済** |

---

## 1. Script Properties（M2）

| Key | 値 |
|-----|-----|
| `APPROVAL_AMAZON_LV4_ENABLED` | `true`（終了後 **false**） |
| `APPROVAL_AMAZON_LV4_TRACK` | **`A`** |
| `APPROVAL_AMAZON_LV4_BRAND_GATE_MODE` | **`manual_ok`**（人間が制限なし確認後のみ） |
| `APPROVAL_AMAZON_LV4_SKIP_EXPORT` | 初回 `true` → 本出力時 `false` |
| `APPROVAL_QUEUE_V1_ENABLED` | 承認①が必要なら `true` |

---

## 2. clasp push

```text
clasp push
```

含む: `AmazonApprovalExport.js`（ASIN解決＝競合列／URL、`SKIPPED_BRAND_GATE`）

---

## 3. 手順

```text
1. 発汗: ブランド制限なし・未オファー（または試験可）を目視
2. 対象子のマスタ在庫を 0（試験中）
3. 承認① amazon APPROVED（親／子）※Lv4は最新APPROVEDバッチを使う
4. Property: TRACK=A / BRAND_GATE_MODE=manual_ok / SKIP_EXPORT=true
5. 21-① or E-4 → ログで offer・asin=B07YND44VN を確認
6. SKIP_EXPORT=false → GENERATED 本出力
7. ローカル: tools/m2_offer_packaged で dry_run → prod
   → *_M2_OFFER_LOADER.csv（中間）
8. SC: 「カタログに既にある商品を出品」→ 出品ファイル(L)を入手（prefilled xlsm）
9. 公式 Loader 埋め（自動）:
   cd tools\m2_offer_packaged
   python m2_listing_loader_fill.py --template "<DLしたListingLoader.xlsm>" --offer-csv "<…_M2_OFFER_LOADER.csv>" --output "<…_M2_ListingLoader_FILLED.xlsm>" --mode dry_run
   python m2_listing_loader_fill.py ... --mode prod
   （--generated で GENERATED 直読みも可。列マップ=listing_loader_map.json）
10. FILLED xlsm を SC へUP（在庫全削除・置換は使わない）
11. processing-summary で成功確認 → 21-③ or E-5（同一 subBatchId）
12. ENABLED=false / BRAND_GATE_MODE 削除または空 / TRACK 戻す / 在庫を戻す
```

**重要**: 案Lの簡易CSV単独UPは SC で拒否され得る。**正は公式 ListingLoader xlsm**（発汗合格で実証済み）。

スモーク:

```text
python m2_offer_packaged.py --config config.smoke.json --mode dry_run
python m2_listing_loader_fill.py --template <公式xlsm> --offer-csv testdata/... --output out_smoke/fill_dry.xlsm --mode dry_run
```

---

## 4. 合格目安

- [x] GENERATED: `track=A` / `variationRole=offer` / asin あり  
- [x] `manual_ok` 無しでは `SKIPPED_BRAND_GATE`（設計どおり）  
- [x] `*_M2_OFFER_LOADER.csv` 生成（中間）  
- [x] 公式 ListingLoader 埋め → SC 受理（成功1／エラー0）  
- [x] UPLOADED_OK／E-5／片付け  

**残GAP**: 公式テンプレの **自動DL**（人手DL＋自動埋めが正）。SC自動UPは対象外。

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-28 | **`m2_listing_loader_fill.py`**: 公式 Loader 自動埋め（dry_run/prod）。 |
| 2026-07-28 | **実機合格**。公式 Loader 埋め→SC成功→E-5。簡易CSV単独UPは不可と確定。 |
| 2026-07-27 | v1実装。試験=発汗 `0832-48s11`／ASIN `B07YND44VN`。案L＋manual_ok。 |
| 2026-07-27 | 下書き初版。 |
