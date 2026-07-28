# SP-API 読取スモーク（人間手順）

**状態**: **実機合格（2026-07-28）**／ツール実装済  
**範囲**: **読取のみ**（LWA トークン取得 ＋ Catalog Items 1件 GET）。出品・在庫書込なし  
**正本接続**: https://developer-docs.amazon.com/sp-api/docs/connecting-to-the-selling-partner-api  
**エンドポイント（日本）**: https://sellingpartnerapi-fe.amazon.com  
**Marketplace**: `A1VC38T7YXB528`

---

## 0. 合格実績

| 項目 | 値 |
|------|-----|
| 日時 | 2026-07-28 20:57（JST） |
| アプリ | `jidousyuppin260728`（本番・自己承認済） |
| LWA | OK（`expires_in=3600`） |
| Catalog | OK／ASIN `B07YND44VN`／title=発汗チェッカー1500 |
| レポート | `tools/spapi_smoke/out/SPAPI_SMOKE_20260728_115740_REPORT.json`（トークン本文なし） |

---

## 1. 前提

- 開発者プロフィール承認済
- 本番アプリ自己承認済
- 手元に **Client ID／Client Secret／Refresh Token** がある
- ロールに **商品の出品（Product Listing）** がある（Catalog 読取用）

---

## 2. Script Properties（将来 GAS 用・命名）

値は **Script Properties のみ**（コード・シート・Git 禁止）。

| Property Key | 内容 |
|--------------|------|
| `SPAPI_LWA_CLIENT_ID` | LWA Client ID |
| `SPAPI_LWA_CLIENT_SECRET` | LWA Client Secret |
| `SPAPI_REFRESH_TOKEN` | 自己承認で得た Refresh Token |
| `SPAPI_MARKETPLACE_ID` | 既定 `A1VC38T7YXB528` |
| `SPAPI_ENDPOINT` | 既定 `https://sellingpartnerapi-fe.amazon.com` |

※ いまのスモークは **ローカル Python**。GAS 実装・**出品書込**は別承認（[LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md](LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md)）。

---

## 3. ローカル実行

```text
cd tools\spapi_smoke
python -m pip install -r requirements.txt
copy config.example.json config.local.json
（config.local.json に3点の秘密を記入）
python spapi_smoke.py
```

環境変数でも可: `SPAPI_LWA_CLIENT_ID` / `SPAPI_LWA_CLIENT_SECRET` / `SPAPI_REFRESH_TOKEN`

成功時:

1. `LWA OK expires_in=3600` 付近  
2. `Catalog OK` ＋ ASIN（既定 `B07YND44VN`）のタイトル断片  
3. `tools/spapi_smoke/out/SPAPI_SMOKE_*_REPORT.json`（トークン本文は書かない）

---

## 4. 合格目安

- [x] LWA access_token 取得成功  
- [x] Catalog HTTP 200  
- [x] REPORT.json に秘密が載っていない  

---

## 5. やらないこと

- Listings Items の PUT／PATCH／削除（書込は別承認）  
- Refresh Token／Secret のチャット・Git 共有  
- Sandbox アプリの認証情報で本番 endpoint を叩く  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-28 | **実機合格**（LWA＋Catalog／発汗 ASIN）。 |
| 2026-07-28 | 初版。`tools/spapi_smoke` 読取スモーク。 |
