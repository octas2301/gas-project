# domain4_inventory（領域4）

適正在庫・モール同期のツール置き場。  
**F1**: 現在庫再計算＋突合。**F2**: 適正在庫TSV（単品）。モール／マスタへ書かない。

## 前提

- Python 3
- Sheets 読取トークン: `tools/c1_hpc_packaged/secrets/`（`competitor_store.client` 経由）
- 在庫スプシは `tools/common_ss_view/registry.json` の `在庫管理アプリ_Octas`

## テスト

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\domain4_inventory
python test_stock_recalc.py
python test_formulas.py
python test_listing_read.py
```

## F1 実行

```powershell
python read_spike.py --sample 5
```

出力: `out/f1_stock_compare_*.tsv`  
手順: [DOMAIN4_P0_F1_HUMAN_RUN.md](../../docs/org/DOMAIN4_P0_F1_HUMAN_RUN.md)

## F2 実行

```powershell
python calc_proper_stock.py --dry-headers
python calc_proper_stock.py --limit 50
```

出力: `out/proper_stock_STUB_*.tsv`  
手順: [DOMAIN4_P0_F2_HUMAN_RUN.md](../../docs/org/DOMAIN4_P0_F2_HUMAN_RUN.md)

## フェーズ

| フェーズ | 内容 |
|----------|------|
| F1 | 読取・再計算・突合 |
| F2 | 適正在庫TSV（単品・いま） |
| F3 | P0検収 |
