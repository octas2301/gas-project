# D新規 — PT／棚ゲート＋Cursor 手渡し（承認パッケージ）

**日付**: 2026-08-01（実装 2026-08-02）  
**状態**: **実装済**（Drive `AMAZON_SHELF_REGISTRY_FILE_ID`／3者スキップ）  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)／[LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)（**P4b-c**）／[LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md)／[LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md)  
**手順**: [D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md](D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md)  
**三者レビュー**: **スキップ**（2026-08-02・実装承認時）

---

## 1. 目的

新規カタログを D から出すとき、**Product Type（PT）未確定／純正テンプレ未整備のまま GENERATED に進む**ことを防ぐ。  
D 完了後は **Cursor Agent への全文コピペ**で PACKAGED を依頼し、GAS→ローカル Python の自動起動は行わない。

会話合意（2026-08-01）を実装可能な要件として固定する。

---

## 2. 二線の完了定義

| 線 | D 完了の意味 | D の後 |
|----|--------------|--------|
| **相乗り（既存）** | SP-API PUT まで（在庫選択どおり） | SC 目視。PACKAGED／Cursor 不要 |
| **新規カタログ** | GENERATED まで＝**準備完了**（掲載完了ではない） | Cursor で PACKAGED → 人間が SC UP → 処理サマリを監視フォルダへ（`UPLOADED_OK` 自動） |

新規＋相乗り同時実行時:

- ゲート成功後、従来どおり新規 GENERATED と相乗り PUT を進めてよい（順序の厳密分離は必須としない）

### 2.1 相乗り・N列 ASIN 空（soft skip）

| ケース | 挙動 |
|--------|------|
| 新規＋相乗り／**相乗りのみ** | ASIN 空の行は **skipped のみ**。blocking にしない（全体は止めない） |
| ASIN 付きが1件以上 | その行だけ相乗り PUT |
| 有効候補が0件（相乗りのみ） | 「相乗りの有効候補が0件」で停止 |
| 有効候補が0件（新規同時） | 相乗りだけ soft 無効化して新規は続行（既存） |

**実装**: `コード.js` `runBatchExportAmazonFacade`（`ASINコード空` で始まる skip 理由を常に soft）。他理由（SKU空・qty不正等）は従来どおり blocking。

---

## 3. GTIN免除（21-⑭）— 運用のみ（本包コード対象外）

| 項目 | 仕様 |
|------|------|
| 単位 | マスタの **Amazon カテゴリ文字列ごと**（SKUごと・実行ごとではない） |
| 確認場所 | シート `▼Lv4実行ログ(Amazon)` の `recordType=EXEMPTION`（または 21-⑭ の「登録済み」） |
| ルール | 新規前に「この親カテゴリは証跡済みか」だけ見る。済みなら再記録不要 |
| 本包 | **コード変更しない**（[D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md) §1d が正） |

---

## 4. D先頭ゲート（実装コア・新規 ON 時のみ）

### 4.1 フロー

```text
D Amazon かつ includeNew
  → PT または Browse が空の親だけ P4b（空セル書込）
  → P4b後も Browse 空 → ハード停止（手入力／競合ASIN・URL確認）
  → Browse Node が SHELF browseIndex に無い → ハード停止＋DL／SHELF同期指示
  → 解決済み PT が棚 registry に無い → ハード停止＋DL／B-T2 指示
  → ある → 新規（および選択していれば相乗り）へ続行
```

### 4.2 固定仕様

| 項目 | 仕様 |
|------|------|
| 発火条件 | `includeNew=true` かつ新規レ点あり。**相乗りのみでは走らない** |
| 挿入位置 | `runBatchExportAmazonFacade` 内の **新規プレフライト**（GENERATED／相乗り PUT より前）。Drive 02 MAIN ゲートと併用可 |
| 失敗時 | **ファサード全体停止**（相乗りも進まない） |
| P4b | PT **または Browse** が空の親だけ既存 `menuAmazonP4bSuggestProductTypeBrowse_` 相当。**両方非空なら Catalog 再実行しない** |
| Browse | **必須**。P4b後も空なら停止（PACKAGED前に人間作業を終わらせる） |
| Property | `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED` 等は **自動 ON しない**。未設定なら停止メッセージで ON を指示 |
| 棚判定（GAS v1） | GAS 可読な `shelf_registry.json` 相当を参照し、解決済み PT が `entries[].productType` に無い → 停止。Browse Node が `browseIndex` に無い → 停止。**指紋再検証は PACKAGED（Python）側**（GAS では指紋まで見ない） |
| registry 同期 | **Drive ファイル ID**（Property `AMAZON_SHELF_REGISTRY_FILE_ID`）。リポジトリの `tools/c1_hpc_packaged/shelf_registry.json` と同内容を Drive に置き GAS が読む |
| 停止時 UI | PT 名、Browse、想定ファイル名（registry にあれば）、**SC から該当テンプレ DL → 06 配置 → 棚登録**（未登録系統は B-T2）を日本語で明示 |
| C末尾 | **本包スコープ外**。後追いで予告のみ可。**フル P4b の C＋D 二重実行は禁止** |

### 4.3 既存ゲートとの関係

- Drive 02 MAIN・GTIN 証跡・`APPROVAL_AMAZON_LV4_ENABLED` 等の既存 fail-closed は維持
- 本ゲートはそれらに **追加**する新規専用プレフライト

---

## 5. D完了後の Cursor 手渡し（方式 B・全文コピペ）

| 項目 | 仕様 |
|------|------|
| GAS→ローカル起動 | **対象外**（PowerShell／Cursor 自動キックは作らない） |
| 対象ダイアログ | `showBatchExportAmazonDaDialog_`（新規 GENERATED 成功時・HtmlService） |
| 冒頭 | **「次の枠を一字残さずコピーし、Cursor の Agent モードに貼れ」** を明示 |
| 枠内 | **自己完結の長い依頼文**（人間が中身を理解しなくてもよい） |
| 埋め込み必須 | 親SKU／subBatchId／GENERATED 手がかり／PT／「SC UP しない」「未対応 PT は止めて報告」／`D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN`（または T1）・`CURRENT_PHASE` を読め |
| 相乗りのみ | 当該枠を出さない（または「PACKAGED 不要」1行） |
| `UPLOADED_OK` | **SC サマリ自動（21-⑮系）を本線**。21-③／E-5 は逃げ道。ダイアログの「必ず Z から 21-③」は削除または「自動失敗時のみ」 |

依頼文の趣旨（実装時に変数埋め）: Agent が B-T1（必要なら dry_run→prod）で PACKAGED を Drive 03 に出し、SC は人間、未対応 PT／棚なしは止めて報告、コミットは指示までしない。

---

## 6. 含まないもの

- GAS メニューからのローカル Python 起動橋
- C末尾 P4b の本実装
- B-T2（新複合の棚登録）本体（ゲートが「無い」と止めた先の別チケット）
- 楽天聖域・Yahoo 本体・B 統合境界変更
- GTIN⑭の自動化／毎回強制 UI
- GENERATED 本体の D 内再実装（薄いファサード維持）

---

## 7. 実装見積り（実装承認後）

| ファイル | 概要 |
|----------|------|
| `コード.js` | D 新規プレフライト（P4b＋棚）、完了ダイアログの Cursor 全文 |
| `AmazonCategoryPt.js` | D から呼ぶ非 UI エントリ（既存 `_` 系流用） |
| 棚読取 | `shelf_registry.json` を GAS 可読に同期する 1 方式 |
| docs | 本承認状態更新・HUMAN_RUN 検収・Facade／PHASE／HANDOVER／LEDGER |

着手前に **変更予定ファイル一覧／概要／リスク** を再提示し、社長の **実装承認** を得る。

### 想定リスクと緩和

| リスク | 緩和 |
|--------|------|
| 棚なしで新規が進む | D 先頭ハード停止 |
| P4b トグル誤常時 ON | 自動 ON 禁止・作業後 false |
| registry 古い／リポジトリと不一致 | 同期方式を1つに固定＋ログに参照版 |
| Cursor 文を貼り忘れる | ダイアログ冒頭でコピペ必須を強調 |
| 指紋不一致を GAS が見逃す | PACKAGED 側で既存どおり停止 |

---

## 8. 検収観点（実装後・HUMAN_RUN）

- [ ] 新規 ON・未登録 PT → D が停止し DL／B-T2 指示が出る
- [ ] 新規 ON・棚あり PT（例 SEASONING）→ ゲート通過して GENERATED まで進む
- [ ] PT 済みでも **Browse 空なら** P4b 再実行／ゲート停止（`all_pt_browse_filled` は両方非空時のみ）
- [ ] PT+Browse 両方非空の親では Catalog 再ヒットしない（ログで確認）
- [ ] 新規成功ダイアログに Cursor 全文枠があり、コピペ指示が冒頭にある
- [ ] 相乗りのみでは Cursor 枠なし（または PACKAGED 不要）
- [ ] ダイアログが SC サマリ自動を本線とし、21-③必須に見えない
- [x] 相乗りのみでも N列ASIN空は行スキップ（全体継続・有効0件のみ停止）
- [x] 棚方式=Drive File ID（`AMAZON_SHELF_REGISTRY_FILE_ID`）を固定（2026-08-02）

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-05 | **Browse 必須**: P4b対象=PT空またはBrowse空。ゲートでBrowse空ハード停止。browseIndex未登録は従来どおりDL指示。 |
| 2026-08-02 | **実装**: Drive棚＋D先頭ゲート＋Cursor手渡し。3者スキップ。 |
| 2026-08-01 | §2.1: 相乗りのみでも N列ASIN空は soft skip（実装反映）。 |
| 2026-08-01 | 初版。会話合意を方針ロック（docs のみ・実装未）。 |
