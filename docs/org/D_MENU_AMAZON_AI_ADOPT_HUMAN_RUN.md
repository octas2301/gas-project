# メニュー8 — 人間手順（v1.8）

**正本**: [D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)  
**ジャンル**: [../RAKUTEN_NAV_GENRE_STAGE3.md](../RAKUTEN_NAV_GENRE_STAGE3.md)  
**Yahoo**: [../YAHOO_CATEGORY_BRAND_STAGE.md](../YAHOO_CATEGORY_BRAND_STAGE.md)／[D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md](D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)

## Properties

| Key | 値 |
|-----|-----|
| `AMAZON_AI_AUTO_ADOPT_ENABLED` | `true`（実行後 false） |
| `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED` | ジャンルもやるとき `true`（実行後 false） |
| `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED` | Yahooカテゴリ／ブランドもやるとき `true`（実行後 false） |
| `RAKUTEN_APP_ID` / `RAKUTEN_ACCESS_KEY` | Ichiba（既存） |
| `RAKUTEN_SERVICE_SECRET` / `RAKUTEN_LICENSE_KEY` | Nav ESA（既存） |
| `YAHOO_SHOPPING_CLIENT_ID` | Yahoo itemSearch（既存） |

## 手順

1. `clasp push`  
2. 上記 Property を設定  
3. Z → 7.5  
4. 確認: KW／テーマ／（トグルON時）楽天ジャンル／（YahooトグルON時）YahooカテゴリID・`:`名／ブランド  
5. 使ったトグルをすべて **false**  

ジャンルトグルOFFならジャンル列は不変。YahooトグルOFFなら Yahoo 3列は不変。
