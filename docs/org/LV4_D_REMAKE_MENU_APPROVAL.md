# D内レ点 — 失敗後の再GENERATED（冪等解除）（承認パッケージ）

**状態**: **実装済＋clasp push済**（2026-08-02・社長実装承認）。実機スモークは任意。  
**実装**: `コード.js`（Dダイアログ `amazonRemake`／`runBatchExportAmazonFacade`）＋ `AmazonApprovalExport.js`（`amazonApprovalLv4RemakeUnlockParents_`）  

**ファイル名**: 歴史的に `LV4_D_REMAKE_MENU_APPROVAL.md`（旧「D修正メニュー」）。**専用メニューは作らない**。  
**親**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)／[LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md](LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md) §冪等  
**三者**: **スキップ**（2026-08-02・社長実装承認・方針ロック時合意どおり）

---

## 1. 目的

SC失敗・属性修正後などに **同じ親で GENERATED を作り直す**とき、手動の 21-④ 往復や別メニューを増やさず、**Dダイアログのレ点1つ**で「冪等解除 → 続けて通常D」を安全に行う。

| やること | やらないこと |
|----------|--------------|
| 対象親の **最新 open subBatch** を `UPLOAD_FAILED` 相当でクローズ | 専用「D修正」メニュー |
| 続けて通常D（棚ゲート・GENERATED・Cursor手渡し文） | xlsm編集・SC自動UL |
| Z の 21-④／E-5 は上級・逃げ道として残す | `UPLOADED_OK` 済み subBatch の自動解除 |
| | 相乗りPUTの強制（相乗りは従来どおり別レ点） |
| | MAP／SHELF同期（Python・人間／Agent） |

---

## 2. UI（Dダイアログ）

**文言案**（実装時に微調整可）:

> **失敗後の再GENERATED（該当親の冪等を解除してから続行）**

| 項目 | 仕様 |
|------|------|
| 既定 | **OFF**（初回出品で誤解除しない） |
| 有効条件 | **新規カタログ ON** のときのみ（相乗りのみでは無視／非表示） |
| 確認 | ON時に確認ダイアログ: 「対象親の最新 open subBatch を冪等解除します。SC成功済み（`UPLOADED_OK`）は対象外」 |
| 解除0件 | ログして **通常Dへ続行**（停止しない） |

browse／PTの修正は本レ点の仕事ではない（D新規先頭の P4b＋SHELF）。  
**強制用**に残す。PT/Browse が変わっている再実行は **§2.1 案Bで自動解除**されるため、通常は OFF でよい。

### 2.1 条件付き自動（案B・2026-08-05）

| 項目 | 仕様 |
|------|------|
| 発火 | 新規 GENERATED が冪等で0件（`idempotentBlocked`） |
| 条件 | 最新が GENERATED/PACKAGED かつ、マスタ `PT\|BrowseNodeId` が Property 指紋と**異なる**（または指紋未保存＝旧GENERATED） |
| 動作 | 該当親だけ `UPLOAD_FAILED` → **同じD内で** Lv4 を1回再実行 |
| 非発火 | 指紋同一／`UPLOADED_OK`／マスタ未変更の二重押し |
| 指紋 | GENERATED 成功時に `APPROVAL_AMAZON_LV4_PARENT_PT_BROWSE_FP` と meta.json `parentPtBrowseFp` |

---

## 3. 解除対象の粒度（ロック・案A）

**レ点付き親ごと**に、その親に紐づく **最新の open subBatch のみ** を解除する。

| 用語 | 意味 |
|------|------|
| open | `UPLOADED_OK`／`UPLOAD_FAILED` 未確定で、冪等ブロックの原因になり得る状態（例: GENERATED 済み・待ちリスト上） |
| 最新 | 当該親について時刻／作成順が最も新しい open 1件 |
| レ点付き親 | マスタで今回のD対象となる新規レ点の親SKU |

**禁止**: 親に紐づく未クローズ全部の一括解除（案B・却下）。誤解除リスクが高い。

`UPLOADED_OK` が付いている subBatch は **触らない**。誤って対象に入った場合は確認ダイアログで止め、実行しない。

---

## 4. 実行順序（実装コア）

```text
D Amazon かつ includeNew かつ remakeCheckbox=ON
  → 確認ダイアログ
  → レ点付き親ごとに最新 open subBatch を特定
  → 各件を UPLOAD_FAILED 相当で記録（21-④／E-5 と同義の既存経路を呼ぶ）
  → 通常の runBatchExportAmazonFacade 続行
       （PT空ならP4b→棚ゲート→GENERATED→完了ダイアログ Cursor文）
```

- remakeCheckbox=OFF → 従来どおり（冪等ブロック時は人間が Z で 21-④）。
- 21-⑱（P4b）は **本レ点に常設必須としない**（PT空時のみ既存ゲート）。

---

## 5. 検収（実装時）

- [x] コード: Dダイアログレ点・既定OFF・新規OFF時は無効化／引数無視
- [x] コード: 案Aヘルパ＋確認ダイアログ2段（解除確認→本線確認）→確認後に mark
- [ ] 実機: レ点親の最新 open 1件だけが `UPLOAD_FAILED` 相当になる
- [ ] 実機: `UPLOADED_OK` は変更されない（全親が OK のみなら停止メッセージ）
- [ ] 実機: 解除後に同一親で GENERATED が再出され、完了ダイアログに Cursor 全文が出る
- [ ] 実機: 相乗りのみ実行では本レ点が効かない
- [ ] 実機: Z の 21-④ が従来どおり使える

---

## 6. 属性MAP／SHELF（本包外・参照）

属性MAPの正本運用は [LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md](LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md)／[D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md](D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)。  
SHELF差分は別ツール（破壊的 recreate 禁止）。本レ点は **冪等＋D再実行** のみ。

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-05 | **条件付き自動再GENERATED（案B）**: 冪等0件時、PT/Browse指紋差分（または未保存）なら自動解除→再実行。同一指紋はブロック維持。手動レ点は強制用に残す。 |
| 2026-08-02 | 後続メモ起草（専用D修正メニュー案）。 |
| 2026-08-02 | **方針ロック**: 専用メニュー却下 → **D内レ点**。解除粒度=**案A**（レ点親の最新 open subBatch のみ）。コード未。 |
| 2026-08-02 | **実装済**（GAS）＋**clasp push済**。実機スモーク任意。 |
