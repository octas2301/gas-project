# Amazon AI 生成＆一括採用（メニュー8）— 実装承認パッケージ

**日付**: 2026-07-26  
**要件正本**: [D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)  
**HUMAN_RUN**: [D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)  
**親**: [商品マスタ_人間作業エリアとマスタエリア_要件定義.md](../商品マスタ_人間作業エリアとマスタエリア_要件定義.md)／[AI_ROUTING_GEMINI_OPENAI.md](../AI_ROUTING_GEMINI_OPENAI.md)／[RESEARCH_AND_ESTIMATE.md](../RESEARCH_AND_ESTIMATE.md) §1.2  
**状態**: **承認済・実装済・実機未検収**（社長「Amazon AI一括採用 v1 を承認」2026-07-26）  

目的: Amazon 本線効率化のため、マスタのキーワード・商品名・バリエーション名・Amazon 登録項目の **生成→空欄のみ採用→要確認** を GAS メニュー8（実装ラベル **7.5**）で行う。価格・ブランドコードは対象外。

---

## 1. 変更ファイル一覧（実装）

| 種別 | パス | 内容 |
|------|------|------|
| 更新 | `コード.js` | 定数／Zメニュー7.5／`menuAmazonAiGenerateAndAdoptForCheckedParents` ほか採用・要確認 |
| 新規 | `docs/org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md` | 要件正本 |
| 新規 | `docs/org/LV4_AMAZON_AI_ADOPT_IMPLEMENTATION_APPROVAL.md` | 本ファイル |
| 新規 | `docs/org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md` | 人間手順 |
| 更新 | `docs/CURRENT_PHASE.md`／`docs/AGENT_HANDOVER.md`／`docs/CHANGE_LEDGER.md`／RESEARCH §1.2 | フェーズ・履歴 |
| **触らない** | `Yahoo.js`／`generateRakutenCSV` 本体／B統合 Step 順序／価格 CPO | 聖域・スコープ外 |

---

## 2. 実装サマリ

| 項目 | 内容 |
|------|------|
| メニュー | **Z → 7.5**（論理メニュー8。番号8は楽天CSV） |
| トグル | `AMAZON_AI_AUTO_ADOPT_ENABLED` 既定 **false** |
| 対象 | M-A（レ点親）。厳格は `AMAZON_AI_ADOPT_REQUIRE_LISTING_CK` |
| 生成 | **しない**（v1.4）。既存▼マスタ候補のみ |
| 採用 | 共有KW横断＋モール別検索/補足横断（完全一致）。最終名上限超は検索KWから弱語削除 |
| 上限 | Amazon/Yahoo **75**、楽天 **120** |
| 触らない | メーカー名／商品名ベース／メーカー品番の書込・商品名案・カテゴリ・バリ・ブランド・定価・sync・部分一致 |
| 採用 | 空欄のみ（プルダウン先頭／商品名推奨＋dedupe／検索KW150切詰／バリ値コピー） |
| 要確認 | `要確認_{短縮}`・#8B0000＋白字・セルメモ |

---

## 3. 復元

- Property false  
- `git revert`／メニュー 7.5 削除  

---

## 4. 承認記録

> **Amazon AI一括採用 v1 を承認（2026-07-26）**: メニュー8・空欄のみ採用・M-A・要確認列ごと（濃い赤白字＋メモ）・商品名dedupe（メモのみ）・価格/ブランド/接頭辞接尾辞は対象外。トグル既定 false。楽天CSV・Yahoo.js・B統合境界は触らない。

---

## 5. 次

1. `clasp push`  
2. HUMAN_RUN §3〜4 で1親スモーク  
3. 安定後の B／21-① 前自動実行は **別承認**

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-26 | **v1.4**: 3モール横断＋最終名上限。要 clasp push。 |
| 2026-07-26 | **v1.3**: 最終商品名amazon式の横断dedupe（完全一致）。部分一致なし。要 clasp push。 |
| 2026-07-26 | **v1.2**: 選択・重複削除のみに縮小。要 clasp push。 |
| 2026-07-26 | **緊急修正**: sync 除外。要 clasp push→再検収。 |
| 2026-07-26 | 承認・実装反映。次＝clasp push→実機。 |
| 2026-07-26 | 初版（実装承認待ち）。 |
