# domain4_inventory（領域4）

適正在庫・モール同期のツール置き場。  
**F1（読取スパイク）**: 現在庫再計算＋突合用TSV。モール／マスタへ書かない。

## 前提

- Python 3
- Sheets 読取トークン: `tools/c1_hpc_packaged/secrets/`（`competitor_store.client` 経由）
- 在庫スプシは `tools/common_ss_view/registry.json` の `在庫管理アプリ_Octas`

## F1 実行

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\domain4_inventory
python test_stock_recalc.py
python read_spike.py --sample 5
python read_spike.py --jans 4901234567890
```

出力: `out/f1_stock_compare_YYYYMMDD_HHMMSS.tsv`  
手順: [docs/org/DOMAIN4_P0_F1_HUMAN_RUN.md](../../docs/org/DOMAIN4_P0_F1_HUMAN_RUN.md)

## フェーズ

| フェーズ | 内容 |
|----------|------|
| F1 | 本フォルダの読取・再計算・突合（いま） |
| F2 | 適正在庫TSV（別承認） |
| F3 | P0検収 |
