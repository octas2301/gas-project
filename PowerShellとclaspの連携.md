# Windows PowerShell と clasp の連携（新しい GAS プロジェクトへ）

載せ替え後、ローカル（Cursor のコード）を **新しい GAS プロジェクト** に push する手順です。

---

## 前提

- **clasp** がインストール済み（`npm install -g @google/clasp`）
- 現在の `.clasp.json` の `scriptId` は **元の（古い）プロジェクト** のままです。  
  → このまま `clasp push` すると **古いプロジェクト** に反映されます。  
  → **新しいプロジェクト** に切り替える必要があります。

---

## 手順 1：新しい GAS プロジェクトの script ID を取得する

1. **contact@octas2301.com** で、**新しいスプレッドシート**（本番にした方）を開く。
2. **「拡張機能」→「Apps Script」** で、新しい GAS プロジェクトを開く。
3. **script ID を確認する**（どれか一方でOK）:
   - **方法 A**: 左の **歯車（プロジェクトの設定）** を開き、**「スクリプト ID」** をコピーする。
   - **方法 B**: ブラウザのアドレスバーを見る。  
     `https://script.google.com/home/projects/【ここが script ID】/edit` の **「/projects/」と「/edit」の間** が script ID です。  
     例: `1AbCdEfGhIjKlMnOpQrStUvWxYz1234567890`

---

## 手順 2：.clasp.json を新しいプロジェクトに切り替える

### 方法 A：手動で書き換え（おすすめ）

1. ローカルの **`.clasp.json`** を開く。
2. **`scriptId`** の値を、手順 1 でコピーした **新しいプロジェクトの script ID** に書き換える。
   ```json
   {"scriptId":"【新しい script ID をここに】","rootDir":"."}
   ```
3. 保存する。

### 方法 B：clasp clone で作り直す

1. 既存の `.clasp.json` を **別名で退避**（例: `.clasp.json.old`）。
2. PowerShell でプロジェクトフォルダに移動し、以下を実行:
   ```powershell
   clasp clone "【新しい script ID】"
   ```
3. 新しい `.clasp.json` が作られる。  
   ※`clasp clone` はプロジェクト直下にファイルを pull するため、既存の `コード.gs` / `Yahoo.gs` などと競合することがあります。**方法 A（scriptId の書き換え）の方が安全**です。

---

## 手順 3：ログインと push

1. PowerShell でプロジェクトのフォルダに移動:
   ```powershell
   cd C:\Users\takuy\Desktop\gas-project
   ```
2. 未ログインなら clasp でログイン（**Contact で**）:
   ```powershell
   clasp login
   ```
   ブラウザが開くので **contact@octas2301.com** でログインして許可する。
3. 新しい GAS に push:
   ```powershell
   clasp push
   ```
4. 成功すると、ローカルの `コード.js` / `Yahoo.js` / `appsscript.json` などが **新しい GAS プロジェクト** に反映されます。

---

## 注意

- **push すると、GAS エディタ側のコードが上書き**されます。新しいプロジェクトで手動編集した部分は、push 前にエディタで確認し、必要ならローカルに取り込んでから push してください。
- 本番は **新しいプロジェクト** なので、**今後の編集はローカル（Cursor）で行い、`clasp push` で反映**する運用で問題ありません。

---

## まとめ

| やること | コマンド／作業 |
|----------|----------------|
| 新しい script ID を取得 | GAS の「プロジェクトの設定」または URL から |
| ローカルを新プロジェクトに切り替え | `.clasp.json` の `scriptId` を新しい ID に変更 |
| ログイン（未実施なら） | `clasp login`（Contact で） |
| コードを反映 | `clasp push` |
