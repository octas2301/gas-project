# 消費税輸出免除不適格用連絡一覧表 GAS（clasp 連携）

このフォルダのコードを `clasp push` でスプレッドシートに紐づいた Apps Script に反映できます。

## 初回だけ：Script ID の取得と .clasp.json の設定

1. **転記用スプレッドシート**（「設定」シートがある方）を開く。
2. メニュー **拡張機能** → **Apps Script** でスクリプトエディタを開く。
3. 左の **プロジェクトの設定**（歯車アイコン）をクリック。
4. **スクリプト ID** をコピー（英数字とハイフンの長い文字列）。
5. このフォルダの **`.clasp.json`** を開き、`"scriptId"` の値 **`"ここをスプレッドシートのApps ScriptのScript IDに書き換えてください"`** を、コピーした Script ID に **書き換えて保存**する。

   ```json
   {
     "scriptId": "1Ab2Cd3Ef4Gh5Ij6Kl7Mn8Op9Qr0St1Uv2Wx3Yz",
     "rootDir": "."
   }
   ```

   （上記の `1Ab2Cd3Ef4Gh5Ij6Kl7Mn8Op9Qr0St1Uv2Wx3Yz` の部分を、あなたの Script ID に置き換える。）

## 普段の流れ：ローカルで編集 → push

1. **PowerShell** でこのフォルダに移動する：
   ```powershell
   cd "c:\Users\takuy\Desktop\gas-project\export-tax-list"
   ```

2. ローカルで `Code.gs` などを編集したあと、GAS に反映する：
   ```powershell
   clasp push
   ```

3. 必要なら、GAS 側のコードをローカルに取り込みたい場合：
   ```powershell
   clasp pull
   ```

## 注意

- **初回に `clasp push` する前に**、必ず `.clasp.json` の `scriptId` を書き換えてください。
- `clasp push` は、このフォルダ内の `Code.gs` と `appsscript.json` で、紐づいた Apps Script の内容を上書きします。別プロジェクト用の clasp を使っている場合は、このフォルダでだけ `clasp push` を実行するようにしてください。
