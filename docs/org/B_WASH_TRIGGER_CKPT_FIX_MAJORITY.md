# 洗い ckpt 自己衝突対策 — 三点レビュー多数決＋社長決定

**日付**: 2026-09-01  
**対象**: `trigger_missed`（hint=`trigger_still_queued_or_stuck`）の誤判定  
**実例**: `pr_20260831_225524` / ハニーマザー  
**RUNBOOK**: [THREE_REVIEW_RUNBOOK.md](THREE_REVIEW_RUNBOOK.md)

---

## 1. レビューア結論

| ID | モデル | 結論 |
|----|--------|------|
| R1 | Claude Opus 5 | 条件付き YES |
| R2 | GPT-5.6 Sol | 条件付き |
| R3 | Cursor Grok 4.6 | 条件付き |

**総合**: 条件付き YES（ckpt 全廃は不可。短縮禁止＋overlap 修正＋検出抑制）

---

## 2. 社長決定（2026-09-01）

| # | 質問 | 決定 |
|---|------|------|
| 1 | ckpt 方針 | **全廃しない**。`wait_shortened` 無効＋実行終了見込み基準の ckpt を**先に実装して様子見** |
| 2 | overlap afterMin | **stale 寄せを実装**（`PR_WASH_OVERLAP_STALE_AWARE` 未設定=ON）。8分固定は Property OFF で切戻し。**様子見** |
| 3 | 実装ゲート | **dry 3〜5本 PASS ＋ clasp 前に実機1回**（発火済み trigger 残存確認）を必須 |

---

## 3. 実装内容（承認済）

| 変更 | 関数／Property |
|------|----------------|
| ckpt/overlap で `wait_shortened` 無効 | `prWashCkptNoShorten_` / `PR_WASH_CKPT_NO_SHORTEN`（未設定=ON） |
| overlap 必ず再予約＋stale 寄せ afterMin | `prOverlapWaitAfterMin_` / `PR_WASH_OVERLAP_STALE_AWARE`（未設定=ON） |
| inflight 生存中は `trigger_missed` しない | `prDetectWashTriggerMissed_` |
| 仮想 | `dry_cycle_trigger_watchdog.py` **18/18 PASS**（T-CTW-13〜18 追加） |

**切戻し**: Script Properties  
- `PR_WASH_CKPT_NO_SHORTEN=false` → ckpt も2分化（旧挙動）  
- `PR_WASH_OVERLAP_STALE_AWARE=false` → overlap 8分固定

---

## 4. clasp push 前チェック（人間）

1. `python tools/competitor_store/dry_cycle_trigger_watchdog.py` → 18/18 PASS  
2. `cd C:\Users\takuy\Desktop\gas-project\purchase_research` → `clasp push`  
3. **実機1回**: 洗い発火後、①ログで以下を確認  
   - `full_trigger_ckpt*` 予約に `wait_shortened=1` が**付かない**  
   - overlap 時 `overlap_wait トリガー予約` が出る  
   - 実行中に `trigger_missed` / `trigger_still_queued_or_stuck` が**出ない**  
   - （任意）GAS エディタで発火済み `.after()` トリガーが `getProjectTriggers()` に残るかメモ  

---

## 5. 様子見の判定目安

- **成功**: A-1 以降の `trigger_missed`（同一 armRunId・overlap 相関）が止まる  
- **未解決でも許容**: 取込み ~28分・20分設計線超過（別課題）  
- **ロールバック**: 上記 Property を false にして clasp push
