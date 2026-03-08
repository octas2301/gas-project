# Keepa API 送料データ取得可否の調査

## 結論

- **送料込み価格**: 取得している。`product.csv[7]`（NEW_FBM_SHIPPING）で**送料込みの出品者最安価格履歴**を取得し、競合価格Amazonやモール横断のAmazon候補に利用している。
- **送料単体（円）**: **取得していない**。Keepaの product オブジェクト／csv 配列には「送料のみ」の項目はない。
- **送料無料フラグ**: **取得していない**。公式ドキュメント・検索結果から、`shippingCost` や `isFreeShipping` に相当するフィールドは見当たらない。

このため、**楽天（postageFlag）・Yahoo!（shippingCode）のように「送料無料なら仮定送料0、それ以外は自社送料を加算」という切り替えは、現状のKeepa APIでは実現できない。**

---

## 根拠

### 1. 当コードでの利用

- `コード.js` では `KEEPA_INDEX_NEW_FBM_SHIPPING = 7` で `product.csv[7]` を参照している。
- Keepaの説明では **NEW_FBM_SHIPPING** は「3rd party New price history **including shipping costs**（FBMのみ）」であり、**すでに送料込みの価格**である。
- つまり「価格＋送料」が1つの値として返るだけで、**送料額や送料有無のフラグは返っていない**。

### 2. 公式・二次資料での整理

- **NEW**: Marketplace New 価格履歴。**Shipping and handling costs not included** と明記されている。
- **NEW_FBM_SHIPPING**: FBM の New 価格履歴で、**shipping included**。
- product の `csv` 配列は「価格タイプごとの履歴」であり、送料単体や送料無料フラグ用のインデックスは一般的な説明には出てこない。
- Marketplace Offer 系では `isPrime` や `isShippable` などの記述はあるが、**送料額・送料無料フラグ**に相当する項目は見つかっていない。

### 3. 現状のモール横断での扱い

- モール横断では Amazon 候補の `assumedShipping` は **0** としている（送料込み価格をKeepaで取っているため、追加で送料を足さない）。
- ドキュメント（`モール横断セット数判定_テスト検査結果.md`）の「Amazon（Keepa）現状」は、実装では **送料込み価格は取得して利用している** が、**送料単体・送料無料フラグは取得していない** という理解が正確。

---

## 今後の方針

- **現状のまま**: 競合価格Amazon・モール横断のAmazon候補は「送料込み価格」のみ利用し、仮定送料は 0 のままでよい。
- **楽天・Yahoo!と同様に「送料無料なら0、それ以外は自社送料」としたい場合**: Keepaの公式API仕様で送料または送料無料フラグに相当する項目が提供されていない限り、**APIだけでは実現できない**。別手段（スクレイピングや他サービス）の検討が必要。
