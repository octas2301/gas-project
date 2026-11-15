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
| 対象行 | **出品CK付きの親行のみ**（子SKUが空）。子にレ点があっても調査しない |
| 処理単位 | **1親ずつ**（Amazon→楽天店舗→Yahoo店舗）。複数レ点も順に完走 |
| 時間超過 | **途中保存**（シート追記）＋約1分後 **トリガー再開**（`runB3CompetitorImagesFromTrigger`）。予算約22分。上限12回 |
| 切戻し | Script Properties `B3_IMAGE_TIME_SLICE_ENABLED=false` で時間打ち切りオフ（従来どおり1実行内に全親。GAS上限には注意） |
| プレビュー表示 | **行高＝プレビュー列幅**（既定100px・正方形）。`autoResize` はプレビュー列にかけない（潰れるため）。`B3_IMAGE_PREVIEW_CELL_PX` |
| B-④ | **サブ画像作成（Python／Cursor指示）** — サブ採用CKの JAN を集め指示文表示（GASは合成しない）。[HUMAN_RUN](org/D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md) |
| B-⑤ | **楽天MAIN基礎（Python／Cursor指示）** — 出品CK親の正本紐付け＋案F希望テーマ。GASは合成しない。P1コード済（要 clasp）。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md) |
| 出力シート | **`競合画像取得（必要時B-③実行）`**（右列にサブ採用CKレ点。採用のみは `サブ画像採用ログ`） |
| 99-⑨ | **Z メニューから削除**。関数 `menuTestCompetitorImageUrlsOutput` は**残し**、GAS エディタ等から開発用に実行可能（出力は従来どおり `競合画像取得テスト`）。 |
| モール横断欠け | **レ点＋ミックス構成展開JAN** が `モール横断セット数判定` に無い場合、**B-③本線内で自動実行**（`menuTestCrossMallSetCountJudge(targetJans)`。別メニュー不要。toast＋ログ）。切戻しで確認ダイアログにするなら `ensureCrossMallSheetCoversJansForB3_(…, { askConfirm: true })`。 |
| Amazon 分岐 | **既定は確認なし連続**（Amazon込み・Keepa全件なし。ログ◎→貼付→マスタで解決）。旧ダイアログに戻す: Property `B3_IMAGE_ASK_MALL_MODE=true`（楽天Yahooのみ／精査済み／自動免責）。 |
| 旧②「いいえ」時 | askMall=ON のときだけ **`runKeepaFetchAsinPasteSheet(0)`（全件）** をオファー。 |
| B-① | 実行前に **ASIN貼り付けの◎・セット数確認**の宣言ダイアログ（はいのみ続行）。 |
| B-② | 実行前に **マスタの送料・梱包・競合価格の確認**の宣言ダイアログ（はいのみ続行）。 |
| Amazon ASIN解決 | **ログ◎ → 貼付◎ → マスタ親子**。各段は全候補取得、画像≥1で確定、0件なら次段。Amazon0でも楽天・Yahoo実行。**親ごと**に解決（全JAN一括ではない）。 |
| ログJAN | **AI情報取得data の JAN**（Keepa EAN不可）。過去行遡及不要。 |

## 3. 実装対応（コード）

- **`menuB3MasterCheckedCompetitorImages`**: B-③ のダイアログと分岐。未完了ジョブがあれば続き／最初から。
- **`runB3CompetitorImagesFromTrigger`** / **`runB3CompetitorImageJobBody_`**: 時間スライス・途中保存・トリガー再開。
- **`menuTestCompetitorImageUrlsOutputImpl_(opt)`**: 画像ジョブ入口（Lock・状態引継ぎ）。
- **`resolveAmazonCompetitorAsinsCascade_`**: Amazon ASIN 段0〜2解決（親ごと呼び出し）。
- **`getCircleAsinsFromKeepaFetchLogForJan_`** / **`getCompetitorAsinsFromMasterForJan_`** / **`getJanFromAiDataForAsinPasteBlock_`**
- **`appendKeepaFetchLogRow`**: `jan`（AI由来）をヘッダー名に合わせて追記。
- **`getUniqueCheckedJansFromMasterForImage_`**, **`ensureCrossMallSheetCoversJansForB3_`**: 親のみ JAN 収集とモール横断の不足検知。
- **`menuUpdateCompetitivePriceOnly`** / **`menuRunPriceProposalOnly`**: 先頭に確認ダイアログ。
- **`menuB4SubImageCursorPrompt`**: B-④。サブ採用CK／採用ログから JAN 集約→Cursor 指示ダイアログ。
- **`menuB5RakutenMainBaseCursorPrompt`**: B-⑤。出品CK親の正本紐付け＋希望テーマ→Cursor 指示。Property `B5_RAKUTEN_MAIN_BASE_ENABLED=false` でメニュー非表示。

## 4. 運用上の注意

- **B-③** は **出品CK（レ点）付き行**の JAN が中心（画像ループはレ点行＋構成展開）。**子SKU行だけのレ点**も従来どおり対象。
- **ミックス連動（2026-10-29）**: レ点親（またはその親行）の `セット構成1〜5_親SKU` が埋まっていれば、**単品親の JAN も取得対象に含める**（単品を別レ点しなくてよい）。単品親（構成空）は従来どおりレ点 JAN のみ。切戻し: Script Property `B3_EXPAND_COMBO_PARENT_JANS=false`（未設定=ON）。
- **モール横断も B-③ に内蔵（2026-10-30）**: 不足JANは画像取得前に自動で横断更新。メインの B-③ だけでよい（Zの横断単体は任意）。

### B一括への組み込み（オプション・案1・2026-09-04）

- **既定 OFF**。Property `B_INTEGRATED_INCLUDE_B3_IMAGES`（未設定=false）。
- **メニューは増やさない**。`B. 商品構成/情報の一括生成` 起動時に**毎回**確認（キュー空でも先に出す）:
  - 「続けて競合画像（B-③）も取得しますか？」（タイトル: `B. 統合実行｜競合画像`）
  - はい → Property=true＋run state。**全キューパック完了後**に約1分後 `runB3AfterBIntegratedFromTrigger`（toast `B→B-③`）
  - いいえ → Property=false。従来どおり B のみ（toast は出ない）
- チェーン時の B-③: **Amazon 含む・自動**（Keepa全件なし）。モール横断不足は自動補完。単体 **B-③** は再取得用に残す。
- 切戻し: 起動確認で「いいえ」、または Property `false`。
- 時間超過時はシートに追記済みの親／店舗まで残る。約1分後に自動再開。手動ならメニューで「続きから」。
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
- `B3_IMAGE_JOB_STATE_PROP` / `B3_IMAGE_JOB_SAFE_BUDGET_MS`（22分）/ `B3_IMAGE_JOB_MAX_TRIGGER_RUNS`（12）
- `B3_IMAGE_TIME_SLICE_ENABLED`（Script Properties、既定 true）
- `B3_IMAGE_PREVIEW_CELL_PX`（既定100。プレビュー列幅＝データ行の行高）
- B-③再実行: 破棄前レ点＋採用ログをマージして `サブ採用CK` を復元し、ログは採用行のみ再書込
- `KEEPA_FETCH_LOG_HEADERS` … JAN 列を含む（既存シートは ensure 時に末尾追加）

---

*最終更新: 2026-08-25 — 親のみ・1親ずつ・時間スライス途中保存＋トリガー再開。*
