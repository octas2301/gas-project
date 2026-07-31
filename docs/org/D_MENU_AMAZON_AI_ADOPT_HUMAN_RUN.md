# メニュー8 — 人間手順（v1.10）

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

## 運用ルール（v1.9〜1.10）

1. **KW区切りは半角スペースのみ**（`,` `、` 禁止）。
2. **メインKWは1語だけ**（大〜中カテゴリSEO）。
3. **メイン／CTR／特徴／用途は方向性の違う語**（`大学共同開発品` と `大学共同開発` のような類似はメニュー8が後列を除去）。
4. 最終名超過時は式の後ろ列・後ろワードから削る（Amazon75／楽天120／Yahoo75）。

## 手順

1. `clasp push`  
2. 上記 Property を設定  
3. Z → 7.5  
4. 確認: 4列が似すぎていない／最終名文字数／テーマ／（トグルON時）ジャンル・Yahoo  
5. 使ったトグルをすべて **false**  

ジャンルトグルOFFならジャンル列は不変。YahooトグルOFFなら Yahoo 3列は不変。
