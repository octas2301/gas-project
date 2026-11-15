# P3 かつおシリーズ同一HTML（他袋へAPI反映）

正本: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

**方針**: 袋数SKUすべてに **同じ** `katsuo_odashi` HTML（内容量は `1袋あたり50g`）。袋数差はセットリンクのみ。

## 対象子SKU（JAN 4906283045119）

| 袋 | item_code | 公開URL |
|----|-----------|---------|
| 1 | `ishi-4906283045119-32s14-2` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-32s14-2.html |
| 2 | `ishi-4906283045119-64s18` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-64s18.html |
| 3 | `ishi-4906283045119-97s12` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-97s12.html |
| 4 | `ishi-4906283045119-12s196` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-12s196.html |
| 5 | `ishi-4906283045119-16s120` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-16s120.html |
| 6 | `ishi-4906283045119-19s144` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-19s144.html |
| 8 | `ishi-4906283045119-25s192` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-25s192.html |
| 10 | `ishi-4906283045119-32s140` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-32s140.html |
| 12 | `ishi-4906283045119-38s188` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-38s188.html |
| 15 | `ishi-4906283045119-48s160` | https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-48s160.html |

※1袋はP4c済。未反映の袋だけでも可。全件まとめても可。

## 手順（人間）

1. Property `YAHOO_CAPTION_RICH_LOCK_ENABLED=true`（済ならそのままで可）
2. マスタで上記の**子SKUに出品CK**（親は不要。推奨: 親に `▼マスタ(Yahoo説明テンプレID)=katsuo_odashi`）
3. メニュー **🛒 Yahoo!出品 → Yahoo説明をAPI反映（テンプレ→caption/FS）**
4. ▼ログ: `captionRich.editItem` OK が対象件数ぶん／BATCH 失敗0
5. 公開URLを数本スポット確認（最低: 2袋・10袋・15袋）
   - CTA（`cta_gourmet`）あり
   - 本採用サブ（`katsuo_body_t03` 等）あり
   - **「N袋／○○g」専用表記が無い**（`1袋あたり50g` のみ）

## 任意（日常確認）

反映後、1本だけ通常「出品実行」→ 装飾が消えない（案A'）。

## 失敗しやすい点

- **ed-00006**「反映またはアップロード中のため更新ができません」→ 連続反映ロック。コード側でリトライ＋間隔あり（要 clasp）。失敗2件だけ再CKしてメニュー再実行でも可。

## 合否をAgentに頼むとき

「P3合否判定して」で、ログ＋公開2〜3袋を見ます。
