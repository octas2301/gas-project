# competitor_jitta — clasp（人間）

① (`purchase_research`) と **混ぜない**。

1. 競合スプシを開く → 拡張機能 → Apps Script → 新規プロジェクト  
2. プロジェクト設定の **スクリプト ID** を `.clasp.json` の `scriptId` に書く  
3. Prop: `GEMINI_API_KEY`（必須）  
4. push:

```powershell
cd C:\Users\takuy\Desktop\gas-project\competitor_jitta
clasp push
```

5. ①側 Prop: `PR_JITTA_OWNER=comp`  
6. 競合メニュー「実態AI（競合）」→ トリガー開始  

2026-10-24: 対象集めはサマリ／マスタ各1回読み（未調査が多いと旧実装は起動上限切れ）。push 後「1. 未調査を今すぐ」→ `実態AIログ` に `cj_`＋`resolve=`。Apps Script ログに `collect n=… ms=…`。

2026-10-24: 実態AIプロンプト案B（候補／確信度／入口型／大カテゴリのみ。出力768）。戻し `JittaResolve.js.local_backup_thinB_*`。

2026-10-25 P1: 未調査 U≥31 で日次ブースト。Prop 省略可（コード既定）。  
2026-10-27: ブースト既定 **1000**（旧500）。push 後「1. 未調査を今すぐ」→ `実態AIログ` に `U=` `mode=boost` `daily=…/1000`。切戻し `CJ_DAILY_CAP_BOOST=500` または通常と同じ値。

切戻し: トリガー停止 → ① `PR_JITTA_OWNER=hb`。対象集め戻しは `JittaResolve.js.local_backup_20261024_161152`
