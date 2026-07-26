# Amazon 画像 U4 — 実装承認リクエスト（パッケージ）

**日付**: 2026-07-26  
**要件正本**: [D_MENU_U4_R2_URL_EMBED_REQUIREMENTS.md](D_MENU_U4_R2_URL_EMBED_REQUIREMENTS.md)  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) §9 U4／POC T4  
**状態**: **承認済・実機合格**（社長「U4 v1 承認」＋HUMAN_RUN 2026-07-26）  
**実機**: [D_MENU_U4_HUMAN_RUN.md](D_MENU_U4_HUMAN_RUN.md) §0（runId `U4_20260726_090920_1366af`）

前提クローズ: U2 実機合格・T2・T2再検証合格（80s10・URL単独・18320なし）。

---

## 実装サマリ

| 項目 | 内容 |
|------|------|
| メニュー | **21-⑦** `menuAmazonU4UrlEmbed` |
| コード | `AmazonDriveImageExport.js`（U4）／`AmazonApprovalExport.js`（URL優先）／`コード.js` |
| マスタ列 | `Amazon MAIN URL`／`Amazon PT URL` |
| トグル | `AMAZON_U4_URL_EMBED_ENABLED` 既定 **false** |
| 上限 | `AMAZON_U4_MAX_SKUS`（既定 20） |

## 復元

- Property false  
- `git revert`／メニュー 21-⑦ 削除  

## 承認記録

> **U4 v1 承認（2026-07-26）**: Drive02→R2→マスタURL＋GENERATED優先。xlsm直編集・T3・楽天CSV・Yahooは触らない。トグル既定 false。件数上限あり。
