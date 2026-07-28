# SP-API Listings 書込（人間手順）

**状態**: **v1／v1.1 実機合格**（2026-07-28）／橋渡し v1.2 経由の運用も合格  
**承認**: [LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md](LV4_SPAPI_LISTINGS_WRITE_IMPLEMENTATION_APPROVAL.md)（v1）／[LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md](LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md)（v1.1）／[LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md](LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)（v1.2）  
**範囲**: **既存 ASIN・LISTING_OFFER_ONLY**（価格・在庫・状態・自己発送）。新規カタログ作成なし。GAS からの SP-API 直呼びなし（CSV橋渡しのみ）  
**公式**: https://developer-docs.amazon.com/sp-api/docs/submit-listings-data  

---

## 0. 試験既定

### 0.1 発汗（v1・単行でも可）

| 項目 | 値 |
|------|-----|
| SKU | `lifec-4560151300832-48s11` |
| ASIN | `B07YND44VN` |
| 価格 | `981`（当時）／在庫 `0` |

### 0.2 安眠・アルコール相乗り（v1.1 CSV）

| note | seller SKU（試験） | 相乗りカタログ ASIN | 価格 | 在庫 |
|------|-------------------|---------------------|------|------|
| 安眠 | `lifec-4560151300924-ride01` | `B00A0J0D30` | **1000** | **0** |
| アルコール | `lifec-4560151300139-ride01` | `B0091G3AHY` | **1000** | **0** |

- ASIN は **相乗り先カタログ**（自社C1新規カタログASINではない）  
- `…-ride01` は C1 出品済みSKU（`…-19s124` 等）との衝突回避用。SCの別自社SKUに変えてよい  
- `seller_id` = Seller Central の **出品者トークン**（ポータルの `amzn1.pa.o.…` ではない）

---

## 1. seller_id

Seller Central → 設定（歯車）→ Account Info → **出品者情報** → **出品者トークン**。  
チャットには貼らない。

---

## 2. 設定

```text
cd tools\spapi_listings_write
python -m pip install -r requirements.txt
copy config.example.json config.local.json
copy items.example.csv items.csv
```

| キー | 内容 |
|------|------|
| LWA 3点 | `auth_config_path` → `../spapi_smoke/config.local.json` 可 |
| `seller_id` | **必須**（出品者トークン） |
| `items_csv` | 既定 `items.csv`。複数行 |
| `max_items` | 既定 `5`。超過は実行拒否 |
| `mode` | 初回 `dry_run` |
| `allow_prod` | prod 時のみ `true`（終わったら **false に戻す**） |

`items_csv` を空／削除すると、従来どおり config の `sku`/`asin` 1件モード。

---

## 2.5 Drive CSV から（スプシ橋渡し）

- **推奨（v1.3）**: `python spapi_listings_write.py --fetch-drive --mode dry_run`  
- 手動: Drive の `*_SPAPI_ITEMS.csv` を本フォルダの `items.csv` に置き換えてから §3  
- 正本: [D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)（21-⑧＝**子SKUレ点**／選択行廃止＝v1.2c）

---

## 3. 実行

### dry_run

```text
python spapi_listings_write.py --mode dry_run
```

成功目安:

- `LWA OK`
- 各行 `PUT HTTP 200`（または 202）
- ERROR issue なし
- `summary ok=… fail=0`
- `tools/spapi_listings_write/out/SPAPI_LISTINGS_WRITE_*_REPORT.json`

### prod

1. `"allow_prod": true`  
2. `python spapi_listings_write.py --mode prod`  
3. SC で各 SKU／価格1000／在庫0を目視  
4. **`allow_prod` を false に戻す**

---

## 4. 合格目安

- [x] dry_run で全行 ERROR なし  
- [x] prod 受理（HTTP 200/202・ACCEPTED 等）  
- [x] SC でオファー確認（発汗／安眠 ride01／アルコール ride01）  
- [x] REPORT に秘密が載っていない  
- [x] allow_prod を false に戻した  

---

## 5. やらないこと

- 全件ループ・max_items 無断拡大  
- `allow_prod=true` のまま放置  
- GAS からの SP-API 直呼び（v1.4・別承認）  
- 秘密のチャット／Git 共有  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-28 | **v1.2b 実機**: 21-⑨ APPR→`--fetch-drive` prod（`…0832-48s11`／`B07YND44VN`）。 |
| 2026-07-28 | **v1.2c＋v1.3 実機**: `--fetch-drive` dry_run／prod（`…48s11`／`B00A0J0D30` ACCEPTED）。 |
| 2026-07-28 | **v1.3**: `--fetch-drive`／**v1.2a** UTF-8・cp932。 |
| 2026-07-28 | **実機合格**記録（v1発汗／v1.1 ride01／v1.2橋渡し）。 |
| 2026-07-28 | **v1.2橋渡し**: Drive CSV 手順を §2.5 に追加。 |
| 2026-07-28 | **v1.1**: CSV複数行・max_items・安眠/アルコール相乗り（1000/0）。 |
| 2026-07-28 | v1。`putListingsItem` LISTING_OFFER_ONLY。dry_run=VALIDATION_PREVIEW。 |
