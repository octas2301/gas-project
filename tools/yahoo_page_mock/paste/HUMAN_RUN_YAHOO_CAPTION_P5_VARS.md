# Yahoo説明 P5（共通骨格＋変数自動埋込）人間手順

正本: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

## 目的

手製の第2テンプレIDを増やさず、**食べるおだし系**は共通骨格（`katsuo_odashi`）に  
魚種・内容量・セットリンクをマスタから自動埋込して inject する。

## 事前

```powershell
cd C:\Users\takuy\Desktop\gas-project
clasp push
```

Script Properties: `YAHOO_CAPTION_RICH_LOCK_ENABLED=true`

## 検証SKU（まぐろ・ぶり系）

| 項目 | 値 |
|------|-----|
| JAN | `4906283045317` |
| 例・子SKU | `ishi-beruodashicaz-64s18-2`（1袋）ほか |
| 親のテンプレID列 | **空でよい**（JAN／商品名で自動解決） |
| 手製 `maguro_odashi` | **作らない** |

## 手順

1. まぐろ・ぶり系列の**子SKUに出品CK**（1袋だけで可。初回は1本推奨）
2. メニュー **Yahoo説明をAPI反映**（または通常 ▶ 出品実行）
3. ログを確認:
   - `[YAHOO_CAPTION_RICH] vars tmpl=katsuo_odashi fish=まぐろ・ぶり net=35g sets=…`
   - `[YAHOO_CAPTION_RICH] inject …` または `OK … tmpl=katsuo_odashi`
4. Pro／公開ページで（**写真は見ない**）:
   - 名称が「食べるおだし（まぐろ・ぶり）」等になっている
   - 内容量が 35g 系
   - セットコーナーのリンクが **当該親の子SKU**（`ishi-beruodashicaz-…`）になっている
   - （画像がかつおのままなのは P5 では合格扱い）

## かつお退行チェック（必須）

1. かつお検証SKU（既存 `ishi-4906283045119-…`）1本だけCK → 同メニュー
2. 文言がかつお・50gのまま／セットリンクがかつお子SKUのまま

## 画像について（P5第1スライス）

**P5検収では写真は見ない**（文章・内容量・セットリンクのみ）。本文画像の自動用意・UPは P6 以降。

列 `▼マスタ(Yahoo説明画像1..3)` が空のままだと、HTML内の画像URLはかつお既定のまま残る（今回は許容）。


## 切戻し

- Property `YAHOO_CAPTION_RICH_LOCK_ENABLED=false` → 装飾同送オフ
- コード: Git revert（`YahooCaptionRich.js` / `Yahoo.js` の P5 差分）
