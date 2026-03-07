# 広告運用GASを gas-project に共有する手順

## 前提
- **gas-project** は現在、**1つのGASプロジェクト**（AI出品ツール用スプレッドシート紐づけ）を `clasp` で管理しています（ルートの `.clasp.json`）。
- **広告運用スプレッドシート** は別のスプレッドシートであり、その「拡張機能 → Apps Script」は**別のGASプロジェクト**です。
- ここでは、その広告運用用GASを gas-project フォルダ内に取り込み、コードの確認・開発ができるようにする手順を説明します。

---

## 方法A：clasp で clone する（推奨）

広告運用スプシに紐づいたスクリプトを、**そのまま clone** して gas-project 内のフォルダに置きます。以降、そのフォルダから `clasp push` / `clasp pull` で同期できます。

### ステップ1：広告運用用の Script ID を取得する

1. **広告運用スプレッドシート**を開く。
2. メニュー **「拡張機能」** → **「Apps Script」** を開く。
3. 左の **歯車アイコン（プロジェクトの設定）** をクリック。
4. **「スクリプト ID」** をコピーする（長い英数字の文字列です）。

![Script ID は「プロジェクトの設定」画面にあります](https://developers.google.com/apps-script/guides/clasp#clone)

### ステップ2：clasp でログイン済みか確認する

PowerShell で gas-project のフォルダに移動し、次を実行します。

```powershell
cd c:\Users\takuy\Desktop\gas-project
clasp login
```

ブラウザで Google アカウント認証が求められたら、**広告運用スプシを編集できる同じアカウント**でログインしてください。すでに `clasp login` 済みならこのステップは不要です。

### ステップ3：広告運用用のフォルダを用意し、その中で clone する

**「gas-project」** = あなたのPCにある、このプロジェクトのフォルダ（パス例：`c:\Users\takuy\Desktop\gas-project`）のことです。  
その**中に**、広告運用のGAS用の**新しいフォルダを1つ作ります**。既存の コード.js や Yahoo.js と同じ階層に「広告運用GAS」という名前のフォルダが増えるイメージです。

#### やり方1：エクスプローラーでフォルダを作る（分かりやすい）

1. エクスプローラーで **`c:\Users\takuy\Desktop\gas-project`** を開く。
2. 空白のところで右クリック → **新規作成** → **フォルダー**。
3. フォルダ名を **`広告運用GAS`** にする。
4. その **`広告運用GAS`** フォルダを開く（ダブルクリック）。
5. フォルダの中で **Shift + 右クリック** → **「PowerShell ウィンドウをここに開く」** を選ぶ（または「ターミナルをここに開く」）。
6. 開いた PowerShell で、次を実行（Script ID はステップ1でコピーしたものに置き換え）：

```powershell
clasp clone "ここにScript IDを貼り付け"
```

例：`clasp clone "1ABCdefGHIjklMNOpqrSTUvwxYZ1234567890"`

#### やり方2：PowerShell のコマンドでフォルダを作る

すでに gas-project のフォルダで PowerShell を開いている場合：

```powershell
cd c:\Users\takuy\Desktop\gas-project
mkdir 広告運用GAS
cd 広告運用GAS
clasp clone "ここにScript IDを貼り付け"
```

- `mkdir 広告運用GAS` … gas-project の**中に**「広告運用GAS」という名前の**新しいフォルダを1つ作る**コマンドです。
- `cd 広告運用GAS` … これからコマンドを打つ場所を、その**今作ったフォルダの中**に移すコマンドです。

---

成功すると、`広告運用GAS` フォルダの中に次のようなファイルができます。

- `.clasp.json`（このプロジェクトの Script ID が書かれた設定）
- `Code.gs` や `*.gs`（既存のスクリプトファイル。スプシ側の名前のまま）

### ステップ4：共有できたか確認する

```powershell
dir
```

- `広告運用GAS` 内に `.clasp.json` と `.gs` ファイルがあれば、共有完了です。
- 以降、広告運用の**編集・push**は、必ず **広告運用GAS フォルダに cd して** 行います。

```powershell
cd c:\Users\takuy\Desktop\gas-project\広告運用GAS
clasp push    # ローカルの変更をスプシのスクリプトに反映
clasp pull    # スプシ側の変更をローカルに取得
```

### 注意（clasp が2つある場合）

- **ルート**（`c:\Users\takuy\Desktop\gas-project`）で `clasp push` すると、**AI出品ツール用**のスクリプトにだけ反映されます。
- **広告運用**のスクリプトを更新するときは、必ず **`広告運用GAS` フォルダに cd してから** `clasp push` してください。

---

## 方法B：手動でコードをコピーする（clasp を使わない場合）

広告運用のGASを clasp で管理したくない場合や、まずは「コードだけ見せたい」場合は、次のようにします。

### ステップ1：Apps Script の内容を取得する

1. **広告運用スプレッドシート** → **「拡張機能」** → **「Apps Script」** を開く。
2. 左のファイル一覧で、各 **.gs** ファイルをクリックして開く。
3. 内容を**すべて選択**（Ctrl+A）→ **コピー**（Ctrl+C）。

### ステップ2：gas-project にファイルとして保存する

1. gas-project 内に、例えば **`広告運用GAS`** フォルダを作成する。
2. その中に、**ファイル名.gs ではなく .js** で保存する（例：`Main.js`、`広告取り込み.js` など、スプシ側のファイル名に合わせてよい）。
3. コピーした内容を貼り付けて保存。

複数ファイルがある場合は、1ファイルずつ同じ手順で作成します。

### 注意

- この方法では **clasp の紐づけはありません**。gas-project 側で編集した内容を反映するには、**Apps Script のエディタに手動で貼り直す**か、後から方法Aで clasp clone してから差分をマージする必要があります。
- コードの「共有・確認」が目的なら十分です。

---

## まとめ

| 方法 | やること | メリット |
|------|----------|----------|
| **A: clasp clone** | 広告運用の Script ID で `clasp clone` を **広告運用GAS** フォルダ内で実行 | 以降、そのフォルダから push/pull で同期できる。開発しやすい。 |
| **B: 手動コピー** | Apps Script の .gs の内容をコピーし、gas-project 内の .js に貼り付けて保存 | clasp 不要。コードだけすぐ共有できる。 |

**推奨**：今後の開発（①取り込みなど）も gas-project で行うなら、**方法A** で clone しておくと、こちらでコードを確認・提案しやすく、あなた側でも `広告運用GAS` で `clasp push` するだけで反映できます。

共有ができたら、「広告運用GAS を clone した」「手動で ○○.js を追加した」など、どの方法でどこに置いたかを教えていただけると、次の要件定義・設計に進みます。
