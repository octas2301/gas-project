# D新規 — PT／棚ゲート＋Cursor 手渡し（人間手順）

**状態**: **実装済**（2026-08-02・Drive registry方式）  
**承認**: [LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md](LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)／[D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)／[D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md](D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md)  
**実行**: GAS／SC＝人間。PACKAGED＝Cursor Agent（D完了ダイアログの全文コピペ）。ローカル Python の直接起動は GAS からは行わない。

---

## 0. 要約

| 項目 | 内容 |
|------|------|
| いつ | D で **新規カタログ ON** のときだけ先頭ゲート |
| 何をする | 09 未取込 xlsm を棚へ → PT/Browse 空なら P4b → **Browse 必須** → 棚に PT／Node が無ければ **専用案内** |
| D の後（新規） | 完了ダイアログの枠を **一字残さず** Cursor Agent に貼る → B-T1（新PTは 09 の xlsm から指紋＋列マップ自動登録）→ PACKAGED → 人間が SC UP |
| 相乗りのみ | ゲートなし。Cursor 枠なし。**ASIN空レ点は行スキップ**（有効0件のみ停止） |
| GTIN | カテゴリ初回だけ。確認は `▼Lv4実行ログ(Amazon)` の `EXEMPTION`（[D_ENTRY §1d](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)） |
| 棚参照 | Drive の `shelf_registry.json`（Property `AMAZON_SHELF_REGISTRY_FILE_ID`） |
| 棚なし | E③ が専用案内。案内内の「09フォルダを開く」または `G:\マイドライブ\04.amazonカタログ作成（CSV一括UL）\09.SC純正バルクxlsm保存（人間がDLして保存→Agent確認後06へ）`。SC DL → **09へ保存** → 閉じて **E③をもう一度**。2回目で取込→選択画面（実行は人が押す）。チャット不要 |

---

## 1. 初回セットアップ（Drive棚）

1. リポジトリの `tools/c1_hpc_packaged/shelf_registry.json` を Drive（例: `04.amazonカタログ作成…` 配下）へ置く／同期する  
2. そのファイルの ID を Script Property **`AMAZON_SHELF_REGISTRY_FILE_ID`** に設定  
3. 棚に PT を追加したら **同じ Drive ファイルを更新**（コード push 不要）  
4. 指紋検証は従来どおり PACKAGED（Python）側

---

## 2. 本線手順

1. Property: `APPROVAL_AMAZON_LV4_ENABLED=true`  
   - PT 空の親がある場合のみ `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED=true`（**自動ONしない**）  
   - 相乗り時は PUT 系（`…_ENABLED`／prodなら `…ALLOW_PROD`）  
2. D → Amazon → **新規 ON**（相乗りは任意）  
3. ゲート（E③／新規 ON）:
   - 先に Drive **09** の未取込 xlsm を棚へ取り込む
   - まだ Node／PT が無い → **専用案内**（エラー画面ではない）。09へ保存して案内を閉じ、**E③をもう一度**
   - 取込済み／もともと棚あり → **選択画面**（新規＋相乗り＋マスタ在庫既定）。**実行**で GENERATED
   - Browse 空 → 停止（同じE③再実行／競合ASIN・URL確認）
4. 完了ダイアログ: **「枠をコピー」→ Cursor Agent に貼付**  
5. Agent が Drive 03 に PACKAGED を出すまで待つ  
6. 人間が SC へ xlsm（＋必要なら画像 ZIP）UP  
7. 処理サマリをダウンロードして件数確認  
   - **成功時**: ファイル名そのまま **08.SC処理サマリ監視** へ → `UPLOADED_OK` 自動。副本は 05 可  
   - **失敗時**: **08へ置かない**。**05.SC処理結果・ログ退避**へ保存し、件数・エラーコードを Cursor へ報告。Agent 対策後に再UL。任意で `_NG` 付きを08へ（`UPLOAD_FAILED`）または 21-③／E-5  
8. 作業後トグルを false に戻す（`LV4`／`P4B`／`PUT`／`ALLOW_PROD`）

---

## 3. 検収チェックリスト

- [ ] `AMAZON_SHELF_REGISTRY_FILE_ID` 未設定で新規 ON → 停止メッセージに Property 名
- [ ] 未登録 PT で新規 ON → 停止メッセージに PT・DL／B-T2 指示
- [ ] 棚あり PT（例 SEASONING）でゲート通過
- [ ] PT+Browse 両方非空では P4b Catalog 再実行なし（ログ `ran=0`／`all_pt_browse_filled`）
- [ ] Browse 空のみでも P4b 対象になり、埋まらなければゲート停止
- [ ] 新規成功時: Cursor 全文枠＋冒頭コピペ指示＋コピーボタン
- [ ] 相乗りのみ: Cursor 枠なし（PACKAGED 不要の短文）
- [ ] ダイアログが 21-③必須に見えない（サマリ自動本線）

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-09-18 | **PACKAGED auto-onboard**: E③の 09 取込はルーティング。充填は B-T1 が 09/06 の `{PT}.xlsm` から指紋＋列マップを自動登録して 03 へ出す。止まるのは xlsm 無し／必須列 MISS。要出品 clasp（手渡し文）。 |
| 2026-09-17 | **E③ 09取込**: 案内→09保存→E③再実行で選択画面。自動実行しない。要 `clasp push`。切戻し `AMAZON_SHELF_INGEST_09_ENABLED=false`。 |
| 2026-08-05 | **Browse必須**: P4b=PT空orBrowse空。ゲートでBrowse空／Node未登録は停止。要 `clasp push`。 |
| 2026-08-02 | **実装済**: Drive registry＋D先頭ゲート＋Cursor手渡しダイアログ。3者スキップ。 |
| 2026-08-01 | ASIN空 soft skip（相乗りのみ含む）を要約に追記。 |
| 2026-08-01 | 初版スケルトン（実装待ち）。 |
