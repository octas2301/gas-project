# Amazon MAP sheet ↔ JSON 同期 — 人間／Agent 手順

**承認**: [LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md](LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md)  
**正本**: **sheet**／派生=**MD**／実行=**JSON（sheetから生成）**  
**SHELF（別系統）**: [D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md](D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md)

---

## 0. 役割

| 誰 | 何をする |
|----|----------|
| 人間 | `▼設定(Amazonマッピング)` の MAP／ERRORS を直す。SC UL |
| Cursor | SC結果を [MAP_SC_ERROR_LEDGER.md](MAP_SC_ERROR_LEDGER.md) に記録し、結論を sheet に反映 |
| Agent | **PACKAGED 前に** `sync_map_sheet_to_column_json.py` → 充填 |

---

## 0b. いつ何をするか（チェックリスト）

### PACKAGED 前（毎回・必須）

- [ ] sheet MAP が意図どおりか確認（直した直後なら必須）
- [ ] `python sync_map_sheet_to_column_json.py --profile <grocery|seasoning> --dry-run`
- [ ] 差分が想定どおりなら apply（`--dry-run` なし）
- [ ] 続けて `c1_bulk_fill_by_name.py`（または D完了ダイアログの Cursor 手順）

### SC 失敗後（Cursor → 人間 → 次回 PACKAGED）

- [ ] 処理サマリを **05** へ退避（成功時だけ通常名で 08）。失敗ファイルを 08 に通常名で置かない
- [ ] `MAP_SC_ERROR_LEDGER.md` に症状・コード・結論を追記
- [ ] 実行に効く結論だけ **sheet MAP／ERRORS** を更新（MDだけで終わらない）
- [ ] 任意: `append_map_sheet_error.py`
- [ ] 再PACKAGED 前に必ず上記「PACKAGED 前」を再実行

### SHELF（Browse）を触ったとき（属性MAPとは別）

- [ ] `sync_shelf_browse_to_map_sheet.py` 系で **差分更新のみ**（破壊的 recreate 禁止）
- [ ] 属性の `sync_map_sheet_to_column_json.py` で SHELF 列を消さない・上書きしない
- [ ] Drive の `shelf_registry.json`／`AMAZON_SHELF_REGISTRY_FILE_ID` がリポジトリと一致するか確認（Dゲート用）

### 初回／回帰（コード側ルールを sheet に戻す）

- [ ] `python push_map_attr_patches_to_sheet.py`（必要なときだけ）
- [ ] 続けて sheet→JSON（§1）

---

## 1. PACKAGED 前（必須・コマンド）

```text
cd tools\c1_hpc_packaged
python sync_map_sheet_to_column_json.py --profile grocery --dry-run
python sync_map_sheet_to_column_json.py --profile grocery
python c1_bulk_fill_by_name.py --mode prod --product-type GROCERY ...
```

SEASONING のときは `--profile seasoning`。

**確認**: dry-run で想定外の大量削除・defaults 消滅が出たら apply しない。原因を sheet 側で直す。

---

## 2. SCエラー後（Cursor）

1. 処理サマリ／画面の症状を `MAP_SC_ERROR_LEDGER.md` に追記  
2. 実行に効く結論だけ MAP 行を更新（例: 既定・transform・doNotUse・notes）  
3. 任意: ERRORS 行追記  

```text
python append_map_sheet_error.py --symptom "..." --cause "..." --map-fix "..." --status OPEN
```

4. 次回 PACKAGED 前に §1 の sheet→JSON  

---

## 3. sheet をコード側の最新ルールに揃える（初回／回帰時）

```text
python push_map_attr_patches_to_sheet.py
python sync_map_sheet_to_column_json.py --profile grocery
```

---

## 4. やってはいけないこと

- MD だけ直して JSON を手編集して PACKAGED  
- sheet を直さず JSON だけ直す（正本が割れる）  
- SHELF 列を属性 sync で消す（SHELF は別ツール）  
- MAP／SHELF の **破壊的 recreate**（差分更新のみ）  
- SC失敗の processing-summary を成功扱いのファイル名で 08 に置く  

---

## 5. 関連

| 用途 | 正本 |
|------|------|
| 属性MAP承認 | [LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md](LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md) |
| エラー台帳（属性） | [MAP_SC_ERROR_LEDGER.md](MAP_SC_ERROR_LEDGER.md) |
| レーンB台帳 | [LANE_B_SC_ERROR_LEDGER.md](LANE_B_SC_ERROR_LEDGER.md) |
| 失敗後の再GENERATED | [LV4_D_REMAKE_MENU_APPROVAL.md](LV4_D_REMAKE_MENU_APPROVAL.md)（D内レ点・実装別承認） |

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | 起草。 |
| 2026-08-02 | **固め**: §0b チェックリスト（PACKAGED前／SC失敗後／SHELF／禁止）。 |
