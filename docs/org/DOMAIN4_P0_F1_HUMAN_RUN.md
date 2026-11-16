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

## 3. 画面突合（G1b）— **Web（推奨）**

正とする画面は **在庫管理アプリのウェブ**（スプシ生表ではない）。

Webの「現在庫」は GAS `apiGetCurrentStockList` → `getCurrentStockBatch` と同じ式なので、F1の Python 再計算と突合する相手として正しい。

### 手順

1. `python read_spike.py --sample 5`（または `--jans …`）を実行し、`out/f1_stock_compare_*.tsv` を開く  
2. **すぐあと**（入荷・出荷・棚卸を挟まない）に、デプロイ済みの在庫アプリURLをブラウザで開く  
3. 下ナビ／メニューから **「在庫」**（`?view=stock`・見出し「現在庫」）を開く  
4. 表の **JAN** と **合計** 列を見る（場所別列は見なくてよい。突合は合計）  
5. TSVの各 `stock_key` について、Webの **合計** を `screen_qty_blank_for_human` に手書き  
6. `diff = screen − stock_qty` がすべて **0** なら G1b 合格  

### 注意

- URLは GAS「デプロイ → ウェブアプリ」のURL（`inventory-app/README.md` の手順で発行したもの）  
- ログインGoogleアカウントが「アクセス許可」シートに入っていること  
- 画面が古いときは **再読み込み**してから見る（アプリ側キャッシュあり）  
- サンプルが5未満しか商品マスタに無い場合は、あるだけ突合し件数を記録（不足はデータ側の制約としてメモ）  

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
