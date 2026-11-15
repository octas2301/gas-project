# 領域4 P0 / F1 HUMAN_RUN — 現在庫読取スパイク

**日付**: 2026-11-15  
**対象**: [DOMAIN4_P0_REQUIREMENTS.md](DOMAIN4_P0_REQUIREMENTS.md) §5 F1  
**コード**: `tools/domain4_inventory/read_spike.py`  
**書込**: なし（Sheets `values.get` のみ）

---

## 0. 合格ゲート（再掲）

| ID | 条件 | 記録欄 |
|----|------|--------|
| G1a | 在庫・出品を読める | |
| G1b | サンプル JAN≥5 で画面と `stock_qty` **±0** | |
| G1c | モール／出品マスタ／inventory-app に更新痕跡なし | |

---

## 1. 事前

1. `tools/c1_hpc_packaged/secrets` に読取可能な token があること  
2. ユニットテスト:  
   `cd tools/domain4_inventory` → `python test_stock_recalc.py` → `OK`

---

## 2. スパイク実行

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\domain4_inventory
python read_spike.py --sample 5
```

- 成功時: `out/f1_stock_compare_*.tsv` ができる  
- `FAIL: missing_sheets` / `no_last_count_date` / `NO_CREDS` → ここで止め、原因をメモ

可能なら「棚卸後に入荷あり／出荷あり／確定後修正あり／在庫0寄り」を `--jans` で混ぜる。

---

## 3. 画面突合（G1b）

1. inventory-app（Webまたはスプシの差異・在庫一覧）を開く  
2. TSVの各 `stock_key` について、**実行直後**（間に入出荷を入れない）に画面の現在庫を見る  
3. TSVの `screen_qty_blank_for_human` に画面値を記入（手で可）  
4. `diff = screen − stock_qty` がすべて **0** であること  

不一致時: 日付境界・セッション・確定後修正を疑い、コード修正は別承認。解消できなければ GAS読取IFを依存タスク化。

---

## 4. 書込なし確認（G1c）

- 本コマンドは更新APIを呼ばない  
- 念のため: 実行前後で出品マスタ「在庫数」や inventory の操作ログが増えていないことをざっと確認  

---

## 5. 記録（実行後に埋める）

| 項目 | 値 |
|------|-----|
| 実行日時 | |
| out ファイル名 | |
| サンプルJAN数 | |
| G1a | 未／OK／NG |
| G1b | 未／OK／NG（差分メモ） |
| G1c | 未／OK／NG |
| 次 | F1合格なら F2実装承認へ |
