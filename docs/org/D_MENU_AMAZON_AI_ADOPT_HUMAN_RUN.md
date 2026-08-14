# メニュー8 — 人間手順（v1.13＋Yahoo 7.5／7.6）

**正本**: [D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)／Yahoo [../YAHOO_CATEGORY_BRAND_STAGE.md](../YAHOO_CATEGORY_BRAND_STAGE.md) §0.1／[E確認](D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md)

## KW・商品名（従来）

**聖域**: **`商品名ベース` は可変不可**（読取のみ。dedupe／trim／正規化もメニュー8では禁止）。正本 [D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)「不変列」。

1. 版履歴で **7.5／7.6前**に戻す必要がある場合のみ（商品名ベース回復）  
2. `clasp push`  
3. `AMAZON_AI_AUTO_ADOPT_ENABLED=true` → Z→**7.5 または 7.6** **1回のみ**  
4. 確認: GD≒70–75／**商品名ベースが実行前と同一**（空・短縮していない）  
5. Property を **false**

## Yahoo モード

| メニュー | Yahooカテゴリ | ブランド |
|----------|---------------|----------|
| Z **7.5** | 価格参照（`price_aware`） | 市場／AI → SHP |
| Z **7.6** | 売れ筋のみ（`popular_only`） | メーカー SHP → **38074** |
| **B統合** | **7.6 と同じ**（Step7.6） | 同上 |

Bゲート: `B_INTEGRATED_MENU8_ENABLED` 未設定＝ON。

`clasp logs` の二重行は Logger+console。実行2回ではない。完了ダイアログ＋トーストで窓が2つに見えることがある。
