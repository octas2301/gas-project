# LV4 — ⑤競合店ASIN 空のみ自動本線（U3推奨）承認包

**状態**: **実装済**（2026-08-13・社長 A〜G 承認）／要 `clasp push`  
**三点**: スキップ  
**正本**: [B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md](B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md) §3.4  
**手順**: [B_SPAPI_RESEARCH_P0_P2_HUMAN_RUN.md](B_SPAPI_RESEARCH_P0_P2_HUMAN_RUN.md) U3b  
**関連**: U3=15-⑯／N列=[B_ASIN_N_AUTO_FILL_REQUIREMENTS.md](B_ASIN_N_AUTO_FILL_REQUIREMENTS.md)

---

## 1. 目的

U3 の高信頼推奨を、**競合店ASINコードが空の親だけ**へ自動記入する。人間◎を置き換えない。

---

## 2. 社長確定（A〜G）

| # | 決定 |
|---|------|
| A | 書込列＝**`競合店ASINコード` のみ** |
| B | **空のみ**（上書き禁止） |
| C | 高信頼＝`brand_yes`／`brand_set_match`／`brand_prefer`／`unanimous`／`majority`／`own_child_set_match`（`brand_catalog_hint` 等は書かない）。空判定は **O列のみ**（P列URLにASINがあっても O 空なら対象） |
| D | 人間◎があり推奨と異なる → **書かない** |
| E | 黄セル＋note＋要確認フラグ（`COMPASIN`） |
| F | 入口＝Z **15-⑳**（B Step は未） |
| G | OFF＝`B_COMP_ASIN_AUTOFILL_ENABLED=false` |

---

## 3. 変更

| パス | 内容 |
|------|------|
| `コード.js` | 15-⑳／`fillCompetitorAsinFromU3Vote_`／`computeCompetitorAsinVoteU3ForParent_` |
| docs | 本承認・要件 §3.4・HUMAN_RUN・PHASE |

**戻し**: Property false／git revert／黄セルは手戻し

---

## 4. 本線化ゲート（運用）

| 指標 | ライン |
|------|--------|
| 試験 | レ点親で 15-⑳ → 黄セル目視 |
| 誤書込 | 重大0（1週間） |

---

## 5. 承認欄

| 項目 | 記入 |
|------|------|
| 日付 | 2026-08-13 |
| 結果 | **承認**（推奨 A〜G） |
| コメント | Agent 実装 |
