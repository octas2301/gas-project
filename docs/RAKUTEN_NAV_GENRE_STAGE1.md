# 楽天ジャンルID NavigationAPI — Stage1 要件（テストシート書込隔離）

**文書種別**: 要件定義＋実装済み（コード）  
**最終更新**: 2026-07-17  
**コード**: `コード.js` の `menuDiagnoseRakutenNavigationGenreStage1Write` / `appendRakutenNavGenreStage1Rows_`  
**メニュー**: Z → 17-⑥ / 99-⑩
**前提**: Stage0 疎通完了（[RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)、`okCount=2/2`、endpoint `/es/2.0/navigation/genres/{id}`、名称は `nameJa` / `nameJaPath`）  
**制約（絶対）**:

- `generateRakutenCSV` および楽天CSV／FTP経路は **非改変**
- **▼商品マスタ(人間作業用)**・**AI情報取得data**・出品CSV／TSV への書込 **禁止**
- Agent は **`clasp push` しない**（人間がローカルで実施）
- シークレット（ESA）をログ・docs・セルに **書かない**
- ジャンルの **マスタ自動割当・出品反映は Stage1 スコープ外**

---

## 1. 目的

Stage0 の読取結果を、**専用テストシートにだけ**行追記し、合否をシート上で追跡できるようにする。  
本番マスタを触らずに「API → 表への永続化」まで閉じたループを検証する。

---

## 2. Stage 境界

| Stage | 内容 | 状態 |
|-------|------|------|
| **0** | Logger のみ疎通（既知 genreId → namePath） | **完了** |
| **1**（本要件） | 同一 API 結果を **テストシートへ追記** | **実装済み**（Property 既定オフ） |
| **2**（将来） | テスト用 **別 scriptId / 別スプレッドシート** で clasp 先を隔離 | 未着手 |
| **3**（将来・要別承認） | マスタ列へのジャンルID提案・人間確認ゲート | 未着手・EC重要変更 |

---

## 3. 書込先（隔離ルール）

### 3.1 シート

| 項目 | 値 |
|------|-----|
| シート名（正） | `▼診断(楽天ジャンルNav)` |
| 無ければ | 初回実行時に **ヘッダー付きで新規作成**してよい |
| 禁止シート | `▼商品マスタ(人間作業用)` / `AI情報取得data` / 楽天・Yahoo 出品出力系 / `00_設定マスタ` のマスタ定義行 |

### 3.2 スプレッドシート

- **Stage1 既定**: 現行運用ブック（`.clasp.json` が指すプロジェクトが紐づく SS）内の **上記専用シートのみ** に書く。  
- **より強い隔離が必要なら Stage2**（別 SS + 別 scriptId）。Stage1 では「マスタ列を触れない」ことを隔離の定義とする。

### 3.3 書込モード

- **追記のみ**（既存行の一括クリア・全シート rewrite はしない）
- 1 実行 = ヘッダー下へ **テストID件数分の行**（失敗行も残す）
- 並び: 新しい実行が下（または上）どちらでもよいが、**同一 `runId` でグルーピングできる**こと

---

## 4. 列定義（ヘッダー行1・A列起算）

| 列 | ヘッダー名 | 内容 |
|----|------------|------|
| A | runId | 実行単位ID（例: `Utilities.getUuid()` 先頭8桁＋時刻でも可） |
| B | ranAt | ISO風タイムスタンプ（JST わかりやすければ可） |
| C | genreId | 6桁ジャンルID |
| D | httpStatus | 数値。未到達は空 |
| E | endpointKey | 例: `nav2_genres_byId` / `nav2_genres_byId_showAncestors` |
| F | nameJa | API `nameJa` |
| G | nameJaPath | API `nameJaPath`（表示用パス） |
| H | expectedHint | コード定数の期待ヒント（突合用。空可） |
| I | nameOk | `TRUE`/`FALSE`（期待ヒントと path の一致判定。Stage0 と同ロジック） |
| J | ok | 総合合否（http200 かつ namePath 非空 等。Stage0 の `ok` に合わせる） |
| K | errorSummary | 失敗時の短文（GFコード・例外名。本文巨大JSONは載せない） |
| L | scriptVersion | 任意。git短ハッシュまたは固定文字列 |

ヘッダー名は **完全一致**でコードから参照する（列番号ハードコード禁止を推奨。ヘッダーマップ方式）。

---

## 5. トグル・認証

| Script Property | 既定 | 意味 |
|-----------------|------|------|
| `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED` | 未設定＝**無効（false）** | `true` のときだけシート書込メニューを実行可 |
| `RAKUTEN_NAV_GENRE_DIAG_ENABLED` | Stage0 どおり | 読取診断と独立。Stage1 は読取＋書込 |

- 認証は既存 `getRakutenLicenseKey` / `getRakutenServiceSecret` のみ。  
- Stage1 有効化は **人間が Script Properties で明示**（誤ってマスタ隣接シートに書く事故を減らす）。

---

## 6. テスト入力（Stage0 と同一）

| genreId | expectedHint（目安） |
|---------|----------------------|
| `101888` | ダイエット・健康 > 健康グッズ > その他 |
| `101535` | 食品 > 魚介類・水産加工品 > セット・詰め合わせ |

API:

1. `GET .../es/2.0/navigation/genres/{id}?showAncestors=true`  
2. 失敗時 `GET .../es/2.0/navigation/genres/{id}`  

パース: `nameJa` / `nameJaPath`（[RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md) 準拠）。

---

## 7. メニュー（実装時の案）

- **Z → 17. 計算式・診断 → 17-⑥ 楽天ジャンル Nav Stage1（テストシート追記）**  
- （任意）**Z → 99. テストメニュー** に同様  
- Stage0 の 17-⑤（書込なし）は **残す**（切り分け用）

ログ標準（既存方針に合わせる）: `runId` / `stepName` / `functionName` / `state`（PENDING/RUNNING/DONE/FAILED）を `Logger.log`。

---

## 8. 成功条件（検収）

すべて満たせば Stage1 実装完了:

1. Property 未設定または `false` のとき、メニューは書込せず明示メッセージで終了  
2. `true` のとき、専用シートにヘッダーが存在し、実行ごとにテストID行が追記される  
3. `101888` / `101535` について Stage0 と同様に `ok=TRUE` がシートに残る（API障害時は `ok=FALSE` と `errorSummary`）  
4. 実行前後で **商品マスタ・AI情報取得data の値・書式が変わらない**（目視または差分なし）  
5. `generateRakutenCSV` の差分が当該実装コミットに含まれない  

---

## 9. 人間向け検証手順（実装後）

1. `git pull` → ローカルで `clasp push`（Agent はしない）  
2. Script Properties: `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED=true`  
3. スプレッドシート再読込 → 17-⑥ 実行  
4. `▼診断(楽天ジャンルNav)` で行と `runId` を確認  
5. マスタシートが無変更であることを確認  
6. 終わったら Property を `false` に戻してよい  

---

## 10. 復元

- Property: `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED=false` または削除  
- シート: `▼診断(楽天ジャンルNav)` を削除してよい（本番データではない）  
- Git: Stage1 実装コミットを `git revert`  

---

## 11. 実装時の承認メモ（次チケット用）

Stage1 **コード実装**に入るときは、本要件を満たす範囲でも次を提示して承認を取ること:

- 変更ファイル一覧（想定: `コード.js` の診断近傍のみ・本 docs・CHANGE_LEDGER）  
- シート追記が「テストシートのみ」であることの再確認  
- リスク: シート名誤りによる誤書込（ヘッダー名・シート名の定数化で緩和）

---

## 12. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-17 | 実装: 17-⑥/99-⑩・専用シート追記・Property 既定 false。マスタ/CSV非書込。 |
| 2026-07-17 | 初版。Stage0完了を前提に、テストシート書込隔離の要件を定義。 |
