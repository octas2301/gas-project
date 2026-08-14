# B統合ハード死対策 — HUMAN_RUN

**日付**: 2026-08-14  
**要件**: [B_HARD_DEATH_SCOPE_REQUIREMENTS.md](B_HARD_DEATH_SCOPE_REQUIREMENTS.md)

## 人間

1. Script Properties: `B_WATCHDOG_ENABLED=false`  
2. トリガーに `runBWatchdogFromTrigger` が残っていれば削除  
3. `clasp push`（Agentはしない）  
4. AI情報取得data にストックしてよい。1列目が `B済` になる（初回Bで自動挿入）  
5. 各行に商品名・卸値(税込)。シリーズは同じIDを連続行に  
6. Bを1回起動（続きダイアログはキャンセルで新規）。あとはパック連鎖  
7. 翌朝: `B済` と `▼B実行サマリ`  
8. 同じAI行を再挿入するときだけ `B済` を手で空にする  

マスタ全消去が必要なら人間が行う。

## Agent

コード・docs のみ。clasp push・本番B・トリガー削除はしない。
