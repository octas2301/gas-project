# 楽天ジャンルID — Stage3 要件（都度API・マスタ書込）

**文書種別**: requirements（実装前の正）  
**最終更新**: 2026-07-26  
**状態**: **実装済**（2026-07-26 承認「楽天ジャンル Stage3（都度API・メニュー8）を承認」）  
**承認パッケージ**: [org/LV4_RAKUTEN_GENRE_STAGE3_IMPLEMENTATION_APPROVAL.md](org/LV4_RAKUTEN_GENRE_STAGE3_IMPLEMENTATION_APPROVAL.md)  
**親**: メニュー8（[org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)）／AGENTS.md §10（RWS候補→Nav確認）

---

## 0. 方針（確定）

| 項目 | 内容 |
|------|------|
| 入手ソース | **都度API**（シートに全ジャンルマスタを持たない） |
| AIの★推奨楽天ジャンル | **正本にしない**（誤判定が多いため無視してよい） |
| 書込先 | `楽天ジャンルID` / `楽天ジャンルID名`（親行） |
| 入口 | **メニュー8**（KW・テーマの後フェーズ） |
| 聖域 | `generateRakutenCSV`／楽天CSV・FTP **非改変** |
| トグル | `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED` 既定 **false** |

---

## 1. なぜ都度APIか

NavigationAPI 2.0 の `genres.get` は **ID指定の取得**のみ（フリーテキストでジャンルを探すAPIではない）。  
商品に合うIDを当てるには、市場の実売商品からジャンルを集計する必要がある。

本PJでは既に **Ichiba Item Search**（`RAKUTEN_APP_ID` / `RAKUTEN_ACCESS_KEY`）がある。  
→ **都度**: 商品キーワードで市場検索 → ヒットの `genreId` を投票 → **ESA NavigationAPI** で名称確定。

---

## 2. 処理フロー（v1）

```text
メニュー8（レ点親）・トグルON
  → クエリ組み立て: JAN優先 → なければ 商品名ベース（＋メインKW短く可）
  → RWS Ichiba Item Search（都度・hits≤30）
       genreId をレスポンスから収集し頻度投票
  → 最頻 genreId を RMS NavigationAPI genres.get で検証
       nameJa / nameJaPath 取得（既存 diagnose パーサ流用）
  → 親行の 楽天ジャンルID / 楽天ジャンルID名 に書込
  → 失敗・低信頼 → 要確認_*（空欄のまま or 触らない）
```

### 2.1 投票ルール

| 条件 | 採用 |
|------|------|
| 同一 genreId がヒットの **過半数** または **件数≥3かつ最多** | 採用 |
| 最多でも件数が薄い（例: 1件のみ／同率首位） | 要確認・書込しない（または Property で緩和） |
| Nav GET 失敗 | 書込しない＋要確認 |

### 2.2 表示名

- `楽天ジャンルID名` には Nav の **`nameJaPath` 優先**（無ければ `nameJa`）
- ID は半角数字のみ

### 2.3 AI候補

- `▼マスタ(★推奨楽天ジャンルID)` は **読まない**（v1）
- 将来、投票が弱いときだけフォールバックする案は別承認

---

## 3. トグル

| Key | 既定 | 意味 |
|-----|------|------|
| `AMAZON_AI_AUTO_ADOPT_ENABLED` | false | メニュー8本体（既存） |
| `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED` | **false** | ジャンル都度APIフェーズ。false ならメニュー8でもジャンルを触らない |

認証:

- Ichiba: `RAKUTEN_APP_ID` / `RAKUTEN_ACCESS_KEY`（既存）
- Nav: `RAKUTEN_SERVICE_SECRET` / `RAKUTEN_LICENSE_KEY`（既存 ESA）

---

## 4. 制約

- 楽天CSV生成・Yahoo.js・B統合境界を変えない  
- シークレットをログ・セルに出さない  
- 全件ループはレ点親のみ（メニュー8の対象行）  
- UrlFetch 回数: 親あたり概ね Ichiba 1〜3回＋Nav 1回。上限親は既存 `AMAZON_AI_ADOPT_MAX_PARENTS` に従う  

---

## 5. 検収

- [ ] ジャンルトグル false でジャンル列が不変  
- [ ] レ点親1件で API→ID/名称が埋まる（または要確認）  
- [ ] AI推奨ジャンルが誤っていても、市場投票結果が優先される  
- [ ] 楽天CSVメニューが従来どおり（非改変）  
- [ ] Property を両トグルとも false に戻せる  

---

## 6. Stage 境界（更新）

| Stage | 内容 | 状態 |
|-------|------|------|
| 0 | Nav GET 疎通 | 完了 |
| 1 | 診断シート追記 | 完了 |
| 2 | 別 scriptId | 任意・本 Stage3 ではスキップ可 |
| **3** | **都度APIでマスタへジャンル書込（メニュー8）** | **実装済**（要 clasp push・トグル） |

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-26 | **実装**: メニュー8へ接続。承認済。要 clasp push。 |
| 2026-07-26 | 初版。都度API（Ichiba投票→Nav確認）。AI推奨は正本外。 |
