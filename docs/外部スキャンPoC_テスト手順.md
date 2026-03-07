# 外部スキャンPoC テスト手順

## 概要
この PoC は、**カメラ起動だけを GAS の外に出し**、スキャン結果を GAS と連携できるか確認するための最小テストです。

- カメラ起動: `inventory-app/external-scan-poc.html`
- GAS 連携:
  - 商品 lookup: `api=externalLookupJsonp`
  - 診断ログ: `api=externalScanLog`
- 戻り先画面: `?test=1&view=scantest`

## 事前準備
1. `inventory-app/Code.gs` の変更を含めて、GAS を **新バージョンで再デプロイ**する
2. Web アプリの `exec` URL を控える
3. `inventory-app/external-scan-poc.html` を開き、以下を書き換える

```html
var GAS_EXEC_URL = 'PASTE_GAS_EXEC_URL_HERE';
```

例:

```html
var GAS_EXEC_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
```

## 公開方法
このHTMLは **HTTPS で公開**してください。簡易的には次のどちらかがやりやすいです。

1. GitHub Pages
2. Netlify Drop

ローカルファイルのままでは、スマホでの実機確認がしづらいため、HTTPS 公開を推奨します。

## テスト手順
1. 公開した `external-scan-poc.html` をスマホの Chrome で開く
2. 「スキャン開始」を押す
3. カメラが起動し、JAN を読み取れるか確認する
4. 読み取り後、画面に以下が出るか確認する
   - `JAN: xxxxx`
   - 商品検索結果
5. 「GAS に戻る」を押す
6. GAS の `scantest` 画面に戻り、`受信した JAN:` が表示されるか確認する
7. スプレッドシートの `デバッグログ` シートで `[外部スキャンPoC]` のログが残るか確認する

## URL クエリでの上書き
HTML を編集せず、URL クエリでも設定できます。

- `gasUrl`: GAS の `exec` URL
- `returnUrl`: 戻り先 URL

例:

```text
https://YOUR-HOST/external-scan-poc.html?gasUrl=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2FAKfycb...%2Fexec
```

戻り先を指定する例:

```text
https://YOUR-HOST/external-scan-poc.html?gasUrl=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2FAKfycb...%2Fexec&returnUrl=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2FAKfycb...%2Fexec%3Ftest%3D1%26view%3Dscantest
```

## 期待する結果
- 外部ページではカメラが起動する
- JAN 読み取りができる
- GAS の商品 lookup が返る
- GAS の `scantest` に JAN を戻せる

## 失敗時の見方
- カメラが起動しない:
  - 外部ページ側の HTTPS やブラウザ権限を確認
- 商品検索結果が出ない:
  - `GAS_EXEC_URL` が誤っていないか確認
  - GAS が再デプロイされているか確認
  - `デバッグログ` に `[外部スキャンPoC]` が出ているか確認
- `access denied`:
  - そのスマホブラウザで、GAS にアクセス許可された Google アカウントにログインしているか確認
