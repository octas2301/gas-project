# Yahooカテゴリ仮想テスト

**目的**: B全体を回さず、カテゴリ選定ロジックだけを合否する。  
**実行**:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/yahoo_category_virt/run_virt.py
```

**比較**: 改修前ベースライン → [BASELINE_20260822.md](BASELINE_20260822.md)

---

## API ヒット率調査（実クエリ）

仮想 fixture とは別。Yahoo itemSearch を叩き、**どのクエリパターンがヒットしやすいか**を CSV 化する。  
パターン表: [HIT_RATE_PATTERNS.md](HIT_RATE_PATTERNS.md)

```powershell
cd C:\Users\takuy\Desktop\gas-project
# 初回: config.example.json を config.local.json にコピーし Client ID を入れる
# （または purchase_research_path3/config.local.json / 環境変数 YAHOO_SHOPPING_CLIENT_ID）
python tools/yahoo_category_virt/hit_rate_yahoo_queries.py --dry-run
python tools/yahoo_category_virt/hit_rate_yahoo_queries.py
```

出力: `tools/yahoo_category_virt/out/hit_rate_*.csv` と `hit_rate_report_*.md`  
読取のみ（マスタ非書込）。レート約 1.1 秒/クエリ。

**ロジック**

1. シリーズ核（len≥4）＋表記ゆれ（おだし⇔お出し）  
2. **同商品(+3)**／**シリーズメーカー(+2)**／ヒット≤5なら **ゆるい(+1・メーカー空可・衝突不可)**  
3. 重み付き多数決 → 票内売れ筋重み  
4. 全滅 → **ゆるいシリーズ第1FB**  
5. それでも0 → AI/Drive（商品語とpath無関係は **ガードで非書込**）  

食品path辞書は使わない。
