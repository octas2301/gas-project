# 領域4 P0 / F2 HUMAN_RUN — 適正在庫TSV（単品）

**日付**: 2026-11-16  
**対象**: [DOMAIN4_P0_REQUIREMENTS.md](DOMAIN4_P0_REQUIREMENTS.md) §5 F2  
**コード**: `tools/domain4_inventory/calc_proper_stock.py`  
**書込**: なし（ローカルTSVのみ・案A第1段）

---

## 0. 合格ゲート

| ID | 条件 | 記録 |
|----|------|------|
| G2a | 親行除外（子SKU空は出力対象外） | |
| G2b | 数値セット数で算出 | |
| G2c | 各行に `proper_qty` または `skip_reason` | |
| G2d | `monthly_sales_source=stub` かつファイル名に STUB | |
| G2e | ミックスは `mix_deferred_p6` | |
| G2f | 名簿に無いキーは `stock_key_not_in_inventory`（0にしない） | |

---

## 1. テスト

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\domain4_inventory
python test_formulas.py
python test_listing_read.py
python test_stock_recalc.py
```

すべて `OK` であること。

---

## 2. 実行

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\domain4_inventory
python calc_proper_stock.py --dry-headers
python calc_proper_stock.py --limit 50
```

- 出力: `out/proper_stock_STUB_YYYYMMDD_HHMMSS.tsv`  
- 全件は `--limit 0`（重い・別途意識）

### 見る列

| 列 | 意味 |
|----|------|
| child_sku | 売りSKU |
| set_count | セット数 |
| stock_key | 倉庫キー（代表JAN等） |
| stock_qty / safety_stock / available / proper_qty | 算出（成功時） |
| skip_reason | 空＝成功。`mix_deferred_p6` / `no_jan` / `stock_key_not_in_inventory` 等 |
| monthly_sales_source | 必ず `stub` |

---

## 3. ゲート確認のしかた

1. TSVを開き、親SKUだけの行が **無い**（G2a）  
2. 成功行の `proper_qty` が `floor(available/set_count)` になっている（G2b）  
3. 全行で proper か skip のどちらか（G2c）  
4. source=stub・ファイル名 STUB（G2d）  
5. 構成列あり行は mix_deferred_p6（G2e）  
6. 在庫アプリに無いJANは skip（stock_qtyを0埋めしていない）（G2f）  

---

## 4. 記録

| 項目 | 値 |
|------|-----|
| 実行日時 | |
| out ファイル | |
| 行数 / ok / skip | |
| G2a–f | 未／OK／NG |
| 次 | F2合格 → F3（P0検収）承認へ |
