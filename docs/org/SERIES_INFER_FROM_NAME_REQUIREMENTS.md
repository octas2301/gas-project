# マスタシリーズ — 不明禁止＋商品名推定

**文書種別**: 要件＋実装  
**日付**: 2026-08-12  
**状態**: **実装済**（要 `clasp push`）  
**方針**: **A**（不明を残さない。推定できなければ空）  
**コード**: `コード.js`（`writeSplitted`／`syncAiDataToMaster`／`inferSeriesFromProductName_`）

---

## ゴール

`マスタシリーズ`／`▼マスタ(シリーズ)` に **「不明」系を書かない**。AIが空・不明のときは **商品名から推定**し、できなければ **空**。

---

## 挙動

| 段階 | 処理 |
|------|------|
| Step5 `writeSplitted` | Gemini/GPT の `series` をマージ。不明系は捨て、空なら商品名＋メーカーから推定 → AI正解列 |
| Step6 `syncAiDataToMaster` | `▼マスタ(シリーズ)` と作業列 `マスタシリーズ` 書き込み時に同じ sanitize |
| プロンプト | `rakuten_attributes.series` に不明禁止・空文字可を明記 |

### 推定ルール（簡易）

1. 商品名内の `「…」`／`『…』`／引用符を優先  
2. メーカー名を除去  
3. 容量・個数・セット表記の手前まで  
4. 2〜40文字以外／不明系 → 空

---

## Property

| Key | 未設定 | 緊急停止 |
|-----|--------|----------|
| `SERIES_INFER_FROM_NAME_ENABLED` | **ON** | `false` で旧挙動（長い方を採用・不明も可） |

---

## 復元

- Property `false`  
- Git revert  

## 人間手順

1. `clasp push`  
2. 既存「不明」行: **Step6 同期だけ**でも sanitize される（商品名ベースがあれば推定）。精度を上げるなら Step5 再取得後に Step6  
3. ログ: `[シリーズ推定]`／`[syncAiDataToMaster] …シリーズ sanitize`  
