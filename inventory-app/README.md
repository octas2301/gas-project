# 在庫管理アプリ (Octas) — GAS

## セットアップ手順

### 1. GAS プロジェクトの作成

```powershell
cd "C:\Users\takuy\Desktop\gas-project\inventory-app"
clasp create --title "在庫管理アプリ" --type standalone
```

既に .clasp.json がある場合は、上記で作成された scriptId が書き込まれます。

### 2. スプレッドシート ID の設定

- 使用するスプレッドシート: `1_kqBcQTcL0Q5--KsAdxMYbPMsaToLEXxOMuDw_z-kwE`
- GAS エディタで **プロジェクトの設定** → **スクリプト プロパティ** に次を追加:
  - プロパティ: `SPREADSHEET_ID`
  - 値: `1_kqBcQTcL0Q5--KsAdxMYbPMsaToLEXxOMuDw_z-kwE`

（未設定の場合は上記 ID がデフォルトで使われます。）

### 3. 初回：シート作成・初期データ投入

- GAS エディタで **Init.gs** の関数 **`initializeSpreadsheet`** を選択し、**実行** する。
- 初回は権限の承認が求められます。承認後、もう一度実行する。
- これで「アクセス許可」「設定」「日報宛先」「商品マスタ」「場所マスタ」「理由マスタ」「担当者マスタ」および各リスト・操作ログのシートが作成され、マスタに初期データが入ります。

### 4. アクセス許可の設定

- スプレッドシートの **「アクセス許可」** シートを開く。
- 2行目以降の **メールアドレス** 列に、アプリにログインを許可する Google アカウントのメールアドレスを1行に1件ずつ入力する。

### 5. デプロイ（Web アプリ）

- GAS エディタで **デプロイ** → **新しいデプロイ** → **種類** で **ウェブアプリ** を選択。
- **次のユーザーとして実行**: **自分**
- **アクセスできるユーザー**: **全員**（匿名は不可。ログインすると doGet でアクセス許可をチェックします）
- **デプロイ** を押し、表示された URL をスマホやブラウザで開く。

### 6. 日報メールのトリガー（任意）

- 毎日 20:00 に在庫移動日報を送るには、GAS エディタで **MailJob.gs** の **`installDailyReportTrigger`** を1回実行する。
- または **トリガー** から **sendDailyMoveReport** を **日ベースのタイマー** で 20:00（Asia/Tokyo）に設定する。
- 宛先はスプレッドシートの **「日報宛先」** シートにメールアドレスを1行1件で列挙する。

## フォルダ構成

- Code.gs … doGet / doPost / アクセスチェック / クライアント用 API
- Init.gs … スプレッドシート初期化（シート作成・初期データ）
- SheetService.gs … 各シートの読書き
- InventoryLogic.gs … 現在庫・平均原価・差異
- MailJob.gs … 在庫移動日報メール
- index.html … スマホ向け UI

## clasp push

```powershell
cd "C:\Users\takuy\Desktop\gas-project\inventory-app"
clasp push
```

## 注意

- 親フォルダ（gas-project）の .clasp.json には `"skipSubdirectories": true` を推奨。inventory-app は別 scriptId で管理するため。
