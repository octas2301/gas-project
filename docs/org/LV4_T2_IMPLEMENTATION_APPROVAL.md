# Amazon 画像 T2 — 実装承認リクエスト

**日付**: 2026-07-24  
**親設計**: [LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) §7.1  
**状態**: **承認済・実装済**（社長「T2承認」2026-07-24）。T3以降は未着手。

---

## 変更ファイル

| ファイル | 内容 |
|----------|------|
| **新規** `AmazonDriveImageExport.js` | Drive `02` から1枚 → R2 PUT → ログ |
| `コード.js` | メニュー **21-⑥** |
| docs（POC／HUMAN_RUN／CURRENT_PHASE／CHANGE_LEDGER／AGENT_HANDOVER） | 必須3点セット |

## 人間の次

[LV4_T2_HUMAN_RUN.md](LV4_T2_HUMAN_RUN.md) どおり `clasp push` → Property → 21-⑥ → URL 200 → トグル off。

## 復元

`AMAZON_DRIVE_R2_UPLOAD_ENABLED=false`／メニュー削除／`git revert`
