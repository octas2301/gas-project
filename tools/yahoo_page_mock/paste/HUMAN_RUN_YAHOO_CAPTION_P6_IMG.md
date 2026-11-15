# Yahoo説明 P6a（本文画像）— 欠画FIX後の手順

正本: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

## 実測（2026-09-15）

楽天サブURL（`image.rakuten.co.jp`）を caption に直載せすると、Proプレビューで**欠画**（壊れたアイコン＋alt「シズル・食感」）。  
文言（国産まぐろ）はOKでも画像だけ落ちる。

→ **楽天CDNは使わない**。表示できるのは `yimg.jp` 等Yahooホスト／追加画像libのみ。

## いまの挙動

1. 説明画像列（libファイル名 or Yahoo安全URL）
2. 楽天サブがあっても **スキップ** → ログ `imgSrc=katsuo_lib_default_skip_rakuten`
3. かつお本採用 lib 3枚

## 復旧手順

```powershell
cd C:\Users\takuy\Desktop\gas-project
clasp push
```

まぐろ1袋（`ishi-4906283045317-32s14-2`）だけCK → **Yahoo説明をAPI反映**  
→ 欠画が消え、かつおlib写真が表示されれば復旧PASS。

魚種専用写真は **P6b**（追加画像にUPしてから差し込み）。
