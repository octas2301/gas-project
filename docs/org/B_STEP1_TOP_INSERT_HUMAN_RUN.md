# B Step1 上挿入 — 人間向け手順（HUMAN_RUN）

**正本**: [B_STEP1_TOP_INSERT_REQUIREMENTS.md](B_STEP1_TOP_INSERT_REQUIREMENTS.md)  
**多数決**: [B_STEP1_TOP_INSERT_THREE_REVIEW_MAJORITY.md](B_STEP1_TOP_INSERT_THREE_REVIEW_MAJORITY.md)  
**状態**: 実装済み（ローカル）→ 本手順で実機確認

---

## 0. 実施前

1. 復元点: `_local_backup/pre_B_STEP1_TOP_INSERT_20260814_001210/`（済）  
2. スプシ Drive コピー（済）  
3. **`clasp push`**（`コード.js` 反映）  
4. Script Properties: **`B_STEP1_TOP_INSERT_ENABLED` = `true`**  
5. マスタのフィルタ／並べ替えを解除  
6. 古い行の出品CKを外し、**今回分のみ**レ点  
7. 前回失敗・`#REF!` の新規ブロックがあれば削除してから再テスト推奨  
8. **推奨: AI情報取得data は1商品だけ**

## 1. 合格条件（必須）

| # | 確認 |
|---|------|
| 1 | 既存行の **AK・IB が値のまま不変** |
| 2 | 新規 AK が `#REF!` でなくユニークSKU |
| 3 | 新規が**見本の下**（データ最上段） |
| 4 | 新規に出品CK |
| 5 | ログに `[STEP1_TOP]` の START → writeStartRow → insertDone → DONE（または FAIL） |

## 2. ログの見方

GAS「実行数」→ 当該実行 → ログで `[STEP1_TOP]` を検索。  
例: `writeStartRow=11` / `insertDone` / `DONE elapsedMs=...`
