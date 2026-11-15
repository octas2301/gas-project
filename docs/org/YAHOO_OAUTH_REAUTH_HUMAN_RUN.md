# Yahoo OAuth 再認証（半自動）— HUMAN_RUN

**最終更新**: 2026-10-23  
**実装**: `Yahoo.js`（再認可・C14記録・期限メール）／`コード.js`（メニュー）  
**非対象**: `editItem`／画像／在庫／出品本体（`getYahooAccessToken` は refresh 回転時の B14 更新のみ）

---

## 0. この店の確定値（忘れ防止・2026-10-23）

| 項目 | 正 |
|------|-----|
| Yahooアプリ名（例） | Octas商品登録システム 系 |
| **コールバックURL（アプリ登録）** | **`http://localhost/`**（末尾 `/` 含む） |
| **Script Property `YAHOO_OAUTH_REDIRECT_URI`** | 上と**完全一致** → **`http://localhost/`** |
| マッピング B14 | リフレッシュトークン（ここには URL を入れない） |
| マッピング C14 | 取得日（再認証成功で自動） |

**混ぜない:**

| 入れ場所 | 入れてよいもの | 入れてはいけないもの |
|----------|----------------|----------------------|
| `YAHOO_OAUTH_REDIRECT_URI` | 戻り先URL（`http://localhost/`） | リフレッシュトークン・Client Secret |
| B14 | リフレッシュトークン | 戻り先URL |

ブラウザが `http://localhost/` で「接続できない」と出ても正常。アドレス欄の `code=` 付きURL全体をダイアログに貼る。

エージェント向け: 人間が Yahoo 再認証・`invalid_grant`・「戻り先どれ？」と聞いたら、**先に本節と §2.1bis を開いて案内する**（値を推測で変えない。アプリ側を変えたときだけ §0 を更新）。

---

## 1. 何をするか

Yahoo! ショッピングAPIの **refresh_token**（マッピング **B14**）は通常 **約4週間**で期限切れになる。切れ前に再認可すると期限を延長できる（公式。文字列が同じままの場合あり）。

本機能:

1. メニューから認可URLを開き、戻り `code` を貼る → Token交換 → **B14 更新**＋**C14＝取得日（東京）**
2. C14＋28日で目安期限を自前計算（Yahooに期限APIは無い）
3. 残り ≤ 警告日数（既定 **7日**）または期限切れ／C14未記録時に **メール**（宛先 **B17**）
4. 任意で日次トリガー（毎日9時前後）

切れ切ったあとの典型ログ（▼ログ）:  
`トークン取得失敗: {"error":"invalid_grant",...,"error_description":"refresh token has expired."}`  
→ 本手順で再認証 → **D の Yahooコースだけ**再実行（楽天をやり直す必要はない）。

---

## 2. 初回セットアップ（1回）

### 2.1 Script Properties

| Property | 必須 | 説明 |
|----------|------|------|
| `YAHOO_OAUTH_REDIRECT_URI` | **必須**（再認証メニュー利用時） | Yahooアプリ管理のコールバックと**完全一致**。**この店の正は §0＝`http://localhost/`** |
| `YAHOO_OAUTH_SCOPE` | 任意 | 既定 `openid`（アプリ／API要件に合わせて変更可） |
| `YAHOO_OAUTH_WARN_DAYS` | 任意 | 警告閾値（日）。既定 `7` |

通常の Yahoo出品だけ（B14 がまだ有効）には `YAHOO_OAUTH_REDIRECT_URI` は不要。再認証メニューで認可URLを組むときだけ必須。

### 2.1bis スクリプトプロパティ画面が 50件超で編集できないとき

GAS「プロジェクトの設定 → スクリプト プロパティ」は **50件超だと一覧が読み取り専用**になる。未登録ならプログラムで入れる（UIに見えなくても未設定のまま）。

出品 GAS エディタに一時関数を貼って実行し、終わったら消してよい:

```javascript
function setYahooOauthRedirectUriOnce() {
  PropertiesService.getScriptProperties().setProperty(
    'YAHOO_OAUTH_REDIRECT_URI',
    'http://localhost/'
  );
  Logger.log('OK: YAHOO_OAUTH_REDIRECT_URI');
}
```

件数を減らして UI を戻すのは任意（Z → Script Properties 退避系）。再認証の必須手順ではない。

### 2.2 マッピングシート `▼設定(Yahooマッピング)`

| セル | 用途 |
|------|------|
| B12 | Client ID |
| B13 | Client Secret |
| B14 | Refresh Token（再認証で上書き） |
| **C14** | **取得日** `yyyy-MM-dd`（B14の右。再認証成功時に自動記録） |
| B17 | 成否メール送信宛先（期限リマインドも同じ） |

**このシートに戻り先URLの欄はない。** 戻り先は Yahooアプリ管理の「コールバックURL」と Script Property のみ。

初回だけ、いま手元の refresh がいつ取得したか分かれば **C14 に手入力**してもよい（以降はメニューが自動更新）。

### 2.3 戻り先URLの調べ方（アプリ側を変えたとき）

1. Yahoo!デベロッパーネットワーク → 当該アプリケーション  
2. **コールバックURL** を確認  
3. 変わっていたら §0 と `YAHOO_OAUTH_REDIRECT_URI` を同じ値に更新（推測で変えない）

---

## 3. 操作手順（再認証）

1. `YAHOO_OAUTH_REDIRECT_URI` が §0 どおり入っているか確認（無ければ §2.1bis）
2. スプレッドシートを再読込（メニュー反映）
3. **🛒 Yahoo!出品** → **Yahoo再認証（認可コード貼付）**  
   （または Z → 14-③）  
   ※「Yahoo説明をAPI反映…」や「認証期限を確認」はトークン更新メニューではない
4. 表示された認可URLをブラウザで開き、ログイン・同意
5. 戻り先（`http://localhost/?code=…`）の URL全体、または `code=` の値をダイアログに貼付
6. 成功ダイアログで **C14 取得日**と目安期限を確認
7. **D の Yahooコースだけ**再実行（楽天済みなら楽天コースは不要）

失敗時: `YAHOO_OAUTH_REDIRECT_URI` 不一致・code 期限切れ・Client 不一致を疑う。実行ログで `menuYahooOauthReauth state=FAILED` を確認。

**注意:** 一括出品ログの「Yahoo!出品完了」「D result=OK」は、トークン FATAL を握りつぶしたあとでも出ることがある。正は ▼ログの `FATAL`／`invalid_grant` と、実際の editItem 成否。

---

## 4. 期限確認・メール

- **Yahoo認証期限を確認（メール可）**: 残り日数をダイアログ表示。警告条件なら **即メール**（B17）
- **日次トリガー設置**: `runYahooOauthExpiryReminderFromTrigger` を毎日実行。同日・同理由は再送抑制
- **トリガー削除**: 上記を削除

警告条件:

- C14 ありかつ残り日数 ≤ `YAHOO_OAUTH_WARN_DAYS`（または負＝期限切れ）
- C14 空 かつ B14 に refresh あり（取得日未記録の促し）

---

## 5. 復元・停止

- トリガーのみ止める: メニュー「トリガー削除」
- コード戻し: `git revert`（対象: `Yahoo.js` 再認可ブロック／`getYahooAccessToken` の B14 回転／`コード.js` メニュー）
- B14／C14 の値自体はシート上で手戻し
- Property だけ外す: `YAHOO_OAUTH_REDIRECT_URI` 削除可（再認証メニューは再び §2.1bis が必要）

---

## 6. 公式メモ（運用の根拠）

- Access Token: 約1時間
- Refresh Token: 通常4週間。切れ前の再認可で延長可
- 期限を返すAPIは無い → 取得日＋28日で自前計算

---

## 7. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-10-23 | §0 この店の正=`http://localhost/`。トークンとURLの取り違え禁止。50件超UIの setProperty 手順。切れ後は Yahooコースのみ。FATAL握りつぶし注意 |
| 2026-07-31 | 初版 |
