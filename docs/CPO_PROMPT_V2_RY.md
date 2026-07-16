# CPO値決めプロンプト V2（Gemini用）— 楽天・Yahoo!

GAS では `getRakutenYahooCPOPromptTemplateV2()` に同一思想を埋め込み、`{{ }}` をマスタで置換する。対象モールごとに JSON キーは **`rakuten`** または **`yahoo`** のいずれか一方。

## 方針

- **送料・利益の最低ラインはここでは決めない**（②.5 と ③／再③ の役割）。Amazon V2（[CPO_PROMPT_V2.md](CPO_PROMPT_V2.md)）と同じ。
- **競合価格（当該モール）とセット数**に基づき、**税込・整数の売価案**を JSON で返す。
- 後処理は `applyCpoV2PricePostProcess_`（競合−1・補間・単調）。**旧⑤の利益率レールは使わない**。

## 入力テーブル

- **②.5 前（既定）**  
  `セット数` / `セット卸値(税込)` / `競合価格(楽天)` または `競合価格(Yahoo!)`  
  ※送料は **含めない**。

- **送料参照モード（任意・②.5 後など）**  
  上記に加え `確定送料(参考)` を付ける（**参考のみ**）。

## 出力（厳守）

**必ず最初に** `## 5. 【JSONデータ】` に続く **1つの** ` ```json ` ブロック。

楽天の例:

```json
{"rakuten":[{"setCount":1,"price":749},{"setCount":2,"price":1498}]}
```

Yahoo! の例:

```json
{"yahoo":[{"setCount":1,"price":749},{"setCount":2,"price":1498}]}
```

JSON 以外の本文は **親行**の `楽天価格戦略` または `Yahoo!価格戦略` に反映する。

## 参照

- [PRICING_CPO_RY_V2_REQUIREMENTS.md](PRICING_CPO_RY_V2_REQUIREMENTS.md)
