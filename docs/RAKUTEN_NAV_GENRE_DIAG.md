# 楽天ジャンルID NavigationAPI 疎通診断

**文書種別**: スパイク／診断（実装あり・出品自動化の完成ではない）  
**最終更新**: 2026-07-16  
**コード**: `コード.js` の `menuDiagnoseRakutenNavigationGenreGet` 等  
**制約**: `generateRakutenCSV` 非改変。商品マスタ／AI情報取得data／出品CSVへは**書かない**（Logger のみ）。

---

## 1. 目的

既知の楽天 **ジャンルID（半角数字6桁）** で NavigationAPI（genres.get 相当）を呼び、**データ取得できるか**を確認する。

---

## 2. トグル

| Script Property | 既定 | 意味 |
|-----------------|------|------|
| `RAKUTEN_NAV_GENRE_DIAG_ENABLED` | 未設定＝**有効** | `false` にするとメニュー実行時に拒否 |

認証は既存の `RAKUTEN_LICENSE_KEY` / `RAKUTEN_SERVICE_SECRET`（`getRakutenLicenseKey` / `getRakutenServiceSecret`）。**値をログや本 docs に書かない。**

---

## 3. テストID（コード内定数・書込なし）

| genreId | 期待ヒント（目安） |
|---------|-------------------|
| `101888` | ダイエット・健康 > 健康グッズ > その他 |
| `101535` | 食品 > 魚介類・水産加工品 > セット・詰め合わせ |

試行順（IDごと最大2回）:

1. `GET https://api.rms.rakuten.co.jp/es/2.0/navigation/genres/{id}?showAncestors=true`（genres.get 相当）  
2. 失敗時 `GET https://api.rms.rakuten.co.jp/es/2.0/navigation/genres/{id}`  

※ 2026-07-16: 旧パス `.../genres?genreId=` は `GF0002 Not Found`。旧1.0 `.../genre/get` は HTML 404。

---

## 4. 人間向け実行手順

1. リポジトリ反映後、ローカルで **`clasp push`**（Agent は push しない）。  
2. スプレッドシートを開き直す（メニュー再読込）。  
3. メニューいずれか:
   - **Z → 17. 計算式・診断 → 17-⑤ 楽天ジャンル NavigationAPI 疎通（書込なし）**
   - **Z → 99. テストメニュー → 99-⑨ …**
4. GAS エディタ「実行数」→ 該当実行の **ログ** で次を確認:
   - `[RakutenNavGenreDiag] start ...`
   - 各 `genreId=... http=... endpoint=... namePath=... ok=true/false`
   - `[RakutenNavGenreDiag] done ... okCount=…`

---

## 5. 実行結果（人間が追記）

| 日時 | 実行者 | 101888 | 101535 | 使えた endpoint | メモ |
|------|--------|--------|--------|-----------------|------|
| 2026-07-16 0:39 | | http404 GF0002 ok=false | 同上 | 旧 nav2 query / nav1 不可 | URL誤り。認証はJSONエラーまで到達 |
| 2026-07-16 0:43 | | http200 ok=true namePath空 | 同上 | nav2_genres_byId | URL成功。名称パース未対応 |
| 2026-07-16 0:46 | | http200 nameOk=true path一致 | 同上 path一致 | nav2_genres_byId | **名称表示まで成功 okCount=2/2** |

### 合否の目安

- **疎通成功**: 少なくとも一方の ID で `ok=true` かつ `namePath` が空でない（期待ヒントと完全一致は必須ではない）。  
- **認証・権限不足**: `authMissing` または 401/403 連続 → 「現状取得不可」と記録して調査完了扱いでよい。  
- **URL仕様違い**: 両方 endpoint が非200 → 公式マニュアルで正しい 2.0 パスを確認し、定数 URL のみ修正する（CSV経路は触らない）。

---

## 6. 復元

- 診断ブロックとメニュー2項目を削除する、または `RAKUTEN_NAV_GENRE_DIAG_ENABLED=false`。  
- Git: 当該コミットを `git revert`。

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-16 | 名称パース: NavigationAPI 2.0 の `nameJa` / `nameJaPath` に対応。 |
| 2026-07-16 | URL修正: `/es/2.0/navigation/genres/{genreId}`（query 形式を廃止）。 |
| 2026-07-16 | 初版。読取疎通診断のみ。 |
