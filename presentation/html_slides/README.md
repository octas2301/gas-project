# 南の海テンプレ風 HTML スライド（Zenn 方式）

[Zenn: Cursorで発表スライドを自動生成する(pptx化も説明)](https://zenn.dev/jam0824/articles/76ecf1bf060e71) および [ENECHANGE: Marp&Cursorで始める、爆速！AIプレゼン資料作成術](https://tech.enechange.co.jp/entry/2025/08/15/100000) を参考に、**HTML でスライドを組み、PDF 保存後に pptx に変換する**方式です。

## イメージ

- **タイトル**：全面に南の海の画像、右に半透明の円形でタイトル（黄アクセント）
- **コンテンツ**：左に「CONTENTS」風の青バー、右に番号付きリスト
- **セクション**：左に大きな円（番号＋見出し）、右に箇条書き
- **表**：ネイビーヘッダー・南の海カラー
- **THANK YOU**：全面画像＋中央に白文字

## 手順

### 1. 画像の準備

`presentation` フォルダで実行し、4枚の南の海画像を `html_slides/images/` にコピーします。

```bash
cd presentation
python copy_images_for_html.py
```

`html_slides/images/` に `beach1.png` ～ `beach4.png` ができていることを確認してください。

### 2. ブラウザで確認

`html_slides/index.html` をブラウザで開きます。

- スライドショー風にしたい場合：URL に `#present` を付けて開く（例: `file:///.../index.html#present`）と、先頭スライドのみ表示され、矢印キーで前後移動できます。
- 通常表示：全スライドが縦に並んだ状態で表示されます（PDF 用）。

### 3. PDF で保存（Zenn 記事と同じ）

1. ブラウザのメニューから **印刷**（Ctrl+P / Cmd+P）を開く
2. **送信先** で **PDFに保存** を選択
3. **レイアウト** を **横向き** に設定
4. **余白** を **なし** にし、**背景のグラフィック** を有効にする
5. 保存して PDF を出力

### 4. PDF を pptx に変換

- [Adobe Acrobat](https://acrobat.adobe.com/) の「PDFを書き出し」→「Microsoft PowerPoint」で pptx に変換できます。
- その他、PDF to pptx のオンラインサービスやツールでも変換可能です。

## ファイル構成

| ファイル | 説明 |
|----------|------|
| `index.html` | 全スライドの HTML（20枚） |
| `slides.css` | 南の海テーマ・レイアウト |
| `slides.js` | キーボード操作（#present 時） |
| `images/` | beach1.png ～ beach4.png（`copy_images_for_html.py` で生成） |

## 参考リンク

- [Cursorで発表スライドを自動生成する(pptx化も説明)【リポジトリ公開】](https://zenn.dev/jam0824/articles/76ecf1bf060e71)
- [前編：非エンジニアでもできる！Marp&Cursorで始める、爆速！AIプレゼン資料作成術](https://tech.enechange.co.jp/entry/2025/08/15/100000)
- [Marp + Cursorでスライドを作成する](https://blog.mksc.jp/contents/create_slides_in_marp/)
