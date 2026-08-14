# B-③ 競合画像取得・B-①/B-② 実行前確認（要件定義・経緯）

## 1. 経緯

- 競合画像の取得は **毎回の自動化フローに含めると API・実行時間が大きく増える**ため、**必要なときだけ**実行したいという要望があった。
- 既存実装は **Z → 99-⑨ 競合画像取得テスト** で、出力先シート名が「テスト」扱いであり、**本番運用上の位置づけが分かりにくかった**。
- **モール横断セット数判定** シートの有無・**ASIN貼り付け（Keepa用）** の ◎ 行といった前提が複雑なため、**ダイアログで運用側の確認と分岐**を入れたいという議論があった。
- **◎は Keepa 取得後の評価ロジックで主に自動付与**されるが、**内容の正しさは人が精査**する必要がある。**1 メニューですべてを機械保証することはできない**ため、「精査済みルート」「自動実行（免責）」「楽天・Yahoo のみ」を分岐で切る案が採用された。
- **②で「いいえ」**の場合、**Keepa 取得（全件）を案内・実行**し、◎精査後に **再度 B-③** を実行する運用とした（案A：内部で `runKeepaFetchAsinPasteSheet(0)`）。
- **貼り付けシート差し替え後も Amazon 画像が欲しい**ため、Keepa取得_ログ（AI由来JAN＋◎）とマスタ競合ASIN（親子）へのフォールバックを追加（2026-08-09）。

## 2. 判断結果（要件の要点）

| 項目 | 判断 |
|------|------|
| 本番メニュー | **AI出品ツール → B-③ マスタレ点行の競合画像取得（必要時のみ実行）** |
| B-④ | **サブ画像作成（Python／Cursor指示）** — サブ採用CKの JAN を集め指示文表示（GASは合成しない）。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md) |
| 出力シート | **`競合画像取得（必要時B-③実行）`**（右列にサブ採用CKレ点。採用のみは `サブ画像採用ログ`） |
| 99-⑨ | **Z メニューから削除**。関数 `menuTestCompetitorImageUrlsOutput` は**残し**、GAS エディタ等から開発用に実行可能（出力は従来どおり `競合画像取得テスト`）。 |
| モール横断欠け | **レ点 JAN が `モール横断セット数判定` に無い**場合、番号ではなく**処理内容を説明**し、了承後 **`menuTestCrossMallSetCountJudge`** を実行してから画像処理へ。 |
| Amazon 分岐 | **楽天・Yahoo のみ**／**Amazon 含む（精査済み）**／**Amazon 含む（自動・免責）** のダイアログ。 |
| ②「いいえ」時 | **`runKeepaFetchAsinPasteSheet(0)`（全件）** をオファー。実行後は **再度 B-③** を案内。 |
| B-① | 実行前に **ASIN貼り付けの◎・セット数確認**の宣言ダイアログ（はいのみ続行）。 |
| B-② | 実行前に **マスタの送料・梱包・競合価格の確認**の宣言ダイアログ（はいのみ続行）。 |
| Amazon ASIN解決 | **ログ◎ → 貼付◎ → マスタ親子**。各段は全候補取得、画像≥1で確定、0件なら次段。Amazon0でも楽天・Yahoo実行。 |
| ログJAN | **AI情報取得data の JAN**（Keepa EAN不可）。過去行遡及不要。 |

## 3. 実装対応（コード）

- **`menuB3MasterCheckedCompetitorImages`**: B-③ のダイアログと分岐。
- **`menuB4SubImageCursorPrompt`**: B-④。サブ採用CK／採用ログから JAN 集約→Cursor 指示ダイアログ。
- **`menuTestCompetitorImageUrlsOutputImpl_(opt)`**: 画像一覧出力の本体。`outputSheetName`・`skipAmazonBlock`・`menuTag` を指定可能。
- **`resolveAmazonCompetitorAsinsCascade_`**: Amazon ASIN 段0〜2解決。
- **`getCircleAsinsFromKeepaFetchLogForJan_`** / **`getCompetitorAsinsFromMasterForJan_`** / **`getJanFromAiDataForAsinPasteBlock_`**
- **`appendKeepaFetchLogRow`**: `jan`（AI由来）をヘッダー名に合わせて追記。
- **`getUniqueCheckedJansFromMasterForImage_`**, **`ensureCrossMallSheetCoversJansForB3_`**: モール横断の不足検知と更新。
- **`menuUpdateCompetitivePriceOnly`** / **`menuRunPriceProposalOnly`**: 先頭に確認ダイアログ。

## 4. 運用上の注意

- **B-③** は **出品CK（レ点）** および **`モール横断セット数判定` に JAN が載っている行**が中心（従来仕様を踏襲）。
- **Amazon 画像の ASIN 解決順**（JANごと）:
  1. **Keepa取得_ログ**の同一JAN・最新実行の評価◎（全ASIN）→ 画像≥1で確定
  2. なければ **ASIN貼り付け（Keepa用）** の評価◎（全ASIN）→ 画像≥1で確定
  3. なければ **マスタ**同一JANの親・子の `競合店ASINコード`（＋`競合AmazonページURL`由来）→ 画像≥1で確定
  4. いずれも0件なら「Amazon(なし)」を明示。**楽天・Yahooは Amazon 成否に依存せず実行**
- **楽天・Yahoo 画像（案A）**: モール横断の **統合セット数「不明」／画像一致ガード** でも、**URL付き行は採用**（ローカルセット数があればキーに使用。無ければ行ユニークキー）。価格反映用の `excludeImageGuard` 厳格パースとは分離。
- **同一画像の整理**: AmazonはメディアID（`/images/I/…`）で同一判定。楽天は `_ex` 除去後のURL等。同一キー内は **画素数近似スコア（SL² または `_ex=W×H`）が大きいURLを残す**（小さければ除外／大きければ置換）。Keepa列挙時も同一IDは高解像度のみ。
- 楽天・Yahoo のみモードでは **Amazon/Keepa ブロックをスキップ**（従来どおり）。
- **Keepa 全件**はトークン・実行時間制限があるため、**大量ブロック時は分割運用**（メニュー A の 20/50 件）も引き続き検討可能。B-③ の「②いいえ」経路では要件どおり **全件（0 件制限）** を呼ぶ。

※ **メディアIDが違うが見た目が近い別写真**は、現状の同一判定では残り得ます（知覚ハッシュは未実装）。

## 5. 関連定数（コード.js）

- `COMPETITOR_IMAGE_B3_OUTPUT_SHEET_NAME` = `競合画像取得（必要時B-③実行）`
- `COMPETITOR_IMAGE_TEST_OUTPUT_SHEET_NAME` = `競合画像取得テスト`（開発用）
- `COMPETITOR_IMAGE_SUB_ADOPT_LOG_SHEET_NAME` = `サブ画像採用ログ`（採用レ点のみ。B-③再実行時に復元）
- B-③再実行: 破棄前レ点＋採用ログをマージして `サブ採用CK` を復元し、ログは採用行のみ再書込
- `KEEPA_FETCH_LOG_HEADERS` … JAN 列を含む（既存シートは ensure 時に末尾追加）

---

*最終更新: 2026-08-09 — 同一画像は高画素優先。案A（不明セットでも楽天Yahoo）。Amazon段0〜2。*
