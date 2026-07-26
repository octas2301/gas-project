# 実装承認パッケージ — 楽天ジャンル Stage3（都度API）

**日付**: 2026-07-26  
**状態**: **承認済・実装済**（「楽天ジャンル Stage3（都度API・メニュー8）を承認」）  
**正本**: [../RAKUTEN_NAV_GENRE_STAGE3.md](../RAKUTEN_NAV_GENRE_STAGE3.md)

---

## 変更予定ファイル

| ファイル | 内容 |
|----------|------|
| `コード.js` | メニュー8にジャンル都度APIフェーズ追加。Ichiba検索で genreId 投票→Nav GET で名称→マスタ書込。共通ヘルパ |
| `docs/RAKUTEN_NAV_GENRE_STAGE3.md` | 要件（作成済） |
| `docs/org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md` | v1.7 追記（ジャンルフェーズ） |
| `docs/org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md` | 手順 |
| `docs/CHANGE_LEDGER.md` / `CURRENT_PHASE.md` / `AGENT_HANDOVER.md` | 記録 |

**新規**: 本承認パッケージ／Stage3 要件（上記）

**触らない**: `generateRakutenCSV`／Yahoo.js／B統合境界／AIジャンル生成プロンプト（正本にしないだけ）

---

## 変更概要

1. AIの★推奨楽天ジャンルは使わない  
2. 都度 **Ichiba Item Search**（既存認証）で市場ヒットの `genreId` を投票  
3. 最多IDを **NavigationAPI genres.get**（既存パーサ）で名称確定  
4. 親の `楽天ジャンルID` / `楽天ジャンルID名` に書く  
5. 専用トグル `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED`（既定 false）  

---

## 想定リスク

| リスク | 緩和 |
|--------|------|
| 誤ジャンル採用（似た商品の多数派） | 過半数／件数閾値。薄い票は要確認 |
| API失敗・クォータ | トグル・MAX_PARENTS・失敗時は非書込 |
| 実行時間 | レ点親のみ。親あたり数リクエスト |
| 楽天CSV誤出品 | CSV非改変。人間がジャンル確認後にCSV |
| EC重要変更 | 本承認後のみ実装 |

---

## 戻し方

- Property 両トグル false  
- `git revert`  

---

## 承認文（コピー用）

> **楽天ジャンル Stage3（都度API・メニュー8）を承認**

承認後に Agent が `コード.js` 実装へ進みます。
