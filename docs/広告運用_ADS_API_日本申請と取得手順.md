# Amazon Ads API — 日本（JP）向け申請・取得手順（正本）

**最終更新**: 2026-10-15  
**対象**: 日本の自社セラー（直取引の広告主）が、自店のスポンサー広告を API で扱う場合  
**公式（日本語）**: [オンボーディング概要](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/overview.html)／[LwA作成](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/create-lwa-app.html)／[アクセス申請](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/apply-for-access.html)／[アクセス割当](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/assign-api-access.html)／[認可グラント](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/get-started/create-authorization-grant.html)／[トークン](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/get-started/retrieve-access-token.html)／[初回呼び出し](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/get-started/first-call.html)

> **PA-API（Product Advertising API）や SP-API とは別物。**  
> 広告レポート・キャンペーン操作は **Amazon Ads API（Advertising API）** のみ。

---

## 0. 結論（海外申請は要るか）

| 疑問 | 答え |
|------|------|
| 米国法人や海外セラー登録が必要か | **不要**（日本セラーのまま直取引広告主として申請可） |
| 申請フォームは日本専用か | **世界共通の Ads API オンボーディング**。日本語ガイドあり |
| 「極東（FE）」とは何か | **APIの地域区分**（日本・豪・シンガポール）。海外移住や米アカウント必須ではない |
| パートナーネットワーク経由が必須か | **自社利用のみなら不要**。代理店・ツール提供ならパートナー経路 |
| 日本データはどのエンドポイントか | **FE**: `advertising-api-fe.amazon.com`／トークン `api.amazon.co.jp`／認可 `apac.account.amazon.com` |

**要するに**: 「海外まで申請がいる」のではなく、**日本店舗でも同じ Ads API 審査を通し、呼び出すときだけ極東（FE）のURLを使う**。

---

## 1. 申請前に揃えるもの（日本・自社）

審査は実質 **「このメール＋会社名に紐づく、有効な広告主アカウントがあるか」** の突合です。

| # | 項目 | 推奨（Octas想定） |
|---|------|-------------------|
| 1 | 使うメール | **Seller Central / Ads / 開発者コンソールで同一**（例: `3358miho@gmail.com`） |
| 2 | 申請区分 | **直取引の広告主様**（自店のみ。パートナーにしない） |
| 3 | 会社／法人名 | **Seller Central・Ads に載っている表記と一字一句揃える**（表示名「オンラインショップ Octas」と法人名が違う場合は、申請欄に何を書いたか控えを残す） |
| 4 | 広告アカウント | Campaign Manager に入れ、**Ads account ID**（`amzn1.ads-account....`）が分かること |
| 5 | アクティブ性 | 直近で配信・費用が発生していること（審査側の「active」判定の実務上の目安） |
| 6 | LwA | **Ads専用の新しいセキュリティプロファイル**（SP-API用があっても流用せず新規作成が公式推奨） |

---

## 2. 日本向け手順（申請〜トークン）

### 手順A — LwA アプリ作成（承認不要）

1. [developer.amazon.com](https://developer.amazon.com/) に **上記メール**でログイン  
2. 開発者コンソール → **Login with Amazon**  
3. **新しいセキュリティプロファイルを作成**  
   - 名前例: `AmazonAdsAPI_Octas_JP`  
   - 説明: 自社スポンサー広告のレポート取得・運用用  
   - 同意プライバシーURL: 自社ポリシーがあればそれ。直取なら任意URL可（公式注記）  
4. **クライアントID／クライアントシークレットを表示**して控える（Git・チャット禁止。Script Properties のみ）

公式: [LwA作成（日本語）](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/create-lwa-app.html)

### 手順B — APIアクセス申請（審査あり・最大72時間）

1. [アクセス申請（日本語）](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/apply-for-access.html) の **「直取引の広告主様」**  
2. **開発者登録と同じメール**でログイン（右上プロフィールを必ず確認。別アカウントだと別フォームになる）  
3. 用途: 自社のスポンサー広告レポート取得・キャンペーン運用の自動化、など正直に  
4. 提出後: 受領メール → 審査完了メール（承認 or 却下＋案内）  
5. `jira@amzn-clicks.atlassian.net` のサポート招待は **承諾**しておく  

> パートナーネットワーク経由は **ツール屋・代理店向け**。自社だけなら直取引でよい。

### 手順C — LwA に API アクセスを割り当て（承認後・一度きり）

1. 承認メールのリンクを開く前に [割当手順](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/onboarding/assign-api-access.html) を読む  
2. **手順Aで作ったプロファイルだけ**を選んで送信  
3. 公式警告: **この紐付けは後から変更できない**

### 手順D — 認可コード → Refresh Token（日本＝FE）

1. 開発者コンソール → 当該プロファイル → **Web設定** → **許可された戻り先URL** に例: `https://www.amazon.co.jp/` を追加  
2. ブラウザで次を開く（**FE用**）:

```text
https://apac.account.amazon.com/ap/oa?client_id=（クライアントID）&scope=advertising::campaign_management&response_type=code&redirect_uri=（許可した戻り先URLと完全一致）
```

3. **Adsに入れるAmazonユーザー**でログイン → 「許可」  
4. リダイレクトURLの `code=` をすぐ控える（**約5分で失効**）  
5. トークン交換（**FE**）: `POST https://api.amazon.co.jp/auth/o2/token`  
   - `grant_type=authorization_code`  
   - `code` / `redirect_uri` / `client_id` / `client_secret`  
6. 返ってきた **refresh_token** を Script Properties に保存（以降はこれで access_token を更新）

公式: [認可グラント](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/get-started/create-authorization-grant.html)／[トークン取得](https://d3a0d0y2hgofx6.cloudfront.net/ja-jp/guides/get-started/retrieve-access-token.html)

### 手順E — 初回API（日本プロファイル）

1. `GET` profiles（地域 **FE**）: ベース `https://advertising-api-fe.amazon.com`  
2. 日本の `profileId` を控える  
3. 以降のヘッダ:
   - `Amazon-Advertising-API-ClientId`
   - `Authorization: Bearer （access_token）`
   - `Amazon-Advertising-API-Scope: （profileId）`

| 用途 | 日本（FE） |
|------|------------|
| API | `https://advertising-api-fe.amazon.com` |
| トークン | `https://api.amazon.co.jp/auth/o2/token` |
| 認可画面 | `https://apac.account.amazon.com/ap/oa` |

NA（`amazon.com` / `advertising-api.amazon.com`）は **米国等向け**。日本店舗だけなら普段は使わない。

---

## 3. 前回なぜ通らなかったか（比較）

### 3.1 サポートが繰り返し言っていること（定型却下）

英語の定型文の要旨:

> 申請した **会社名** と **メール** に紐づく **active な広告主アカウント** が見つからない。

Seller Forums 等でも **同じ定型で何度も却下**される事例が多い。技術不備ではなく **アカウント突合の失敗**が主因。

### 3.2 こちらの状況との対応表

| 観点 | 前回〜今の状況（会話・提出物ベース） | 審査側が見ているもの | ずれやすい点 |
|------|--------------------------------------|----------------------|--------------|
| メール | `3358miho@gmail.com` に統一を推進 | 申請ログインメール＝Adsの持ち主 | 開発者コンソール／申請／Adsで別メールだと不一致 |
| 会社名 | ストア「オンラインショップ Octas」、開発者側「Octas」ロックの可能性 | 申請フォームの **legal / company name** | **表示名と申請名の不一致**が定型却下の定番 |
| アカウント種別 | Seller Central はある／Ads も Campaign Manager 可 | **Advertising の広告主アカウント** | SCだけ見せても足りない（過去サポート指摘どおり） |
| 証拠 | Ads ID・Campaign Manager・メール見えスクショを往復 | メール＋Adsアカウントの一体証明 | 「店名だけ」「SCだけ」だと足りない |
| 経路 | 直取引で申請 → Jira で追加証拠 | 直取引は突合厳しめ | パートナー経路は別審査（自社には向かない） |

### 3.3 「海外申請しなかったから落ちた」か

**ほぼ否。**  
日本は FE エンドポイントを使うだけで、**米アカウント必須の申請フローではない**。  
落ちた理由の本命は **名前／メール／Adsアカウントの突合**側。

### 3.4 いまやる再申請・追記のコツ

1. Jira に **メールが見える Ads 画面**（対応中）を出す  
2. 同じスレに英語で1ブロック:
   - Company name submitted: （申請に書いた文字列をそのまま）
   - Ads login email: `3358miho@gmail.com`
   - Store / Ads display name: `オンラインショップ Octas`
   - Ads account ID: `amzn1.ads-account....`
   - Use case: Direct advertiser, own Sponsored Products only (Japan)
3. Advanced Tools Center の **My Apps** を確認（既に Approved なのに再申請しているケースあり）  
4. 法人名がロックされて直せない場合は、**申請名＝ロック名**で揃え、表示名は「also known as / storefront name」として Jira に書く  

---

## 4. 承認後にやること（本PJ）

| 順 | 内容 | 置き場 |
|----|------|--------|
| 1 | Client ID / Secret / Refresh Token / profileId(JP) | 広告運用GAS の Script Properties |
| 2 | ①② 検索用語レポート取得 → `AMZ_FOLDER_ID` | 既存隔週トリガーの先頭 |
| 3 | （後）⑥ API送信・未広告ASIN自動作成 | 第1スライス（バルク人手UL）のあと |

---

## 5. チェックリスト（再申請前）

- [ ] 開発者・申請・Ads が **同一メール**
- [ ] 区分が **直取引の広告主**（パートナーではない）
- [ ] 申請の会社名と、SC/Ads のどちらかに載る正式名が一致
- [ ] Campaign Manager に入れる／Ads account ID がある
- [ ] Ads用 **新規** LwA プロファイルがある
- [ ] My Apps に既存 Approved が無いことを確認
- [ ] Jira 招待を承諾済み
- [ ] 承認後は FE の認可URL／`api.amazon.co.jp`／`advertising-api-fe.amazon.com` を使う

---

## 6. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-10-15 | 初版。公式日本語オンボーディング＋FE地域＋前回却下比較。海外法人必須ではない旨を明記。 |
