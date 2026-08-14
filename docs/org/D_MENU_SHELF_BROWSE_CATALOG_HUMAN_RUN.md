# SHELF Browse 網羅カタログ — 人間手順

**状態**: **実装済**（2026-08-02・要 clasp push／D〜SCは人間）  
**承認**: [LV4_SHELF_BROWSE_CATALOG_APPROVAL.md](LV4_SHELF_BROWSE_CATALOG_APPROVAL.md)  
**実行**: 抽出・MAP同期・PACKAGED＝**Agent**。SC DL／UL・Dメニュー＝**人間**。

---

## 0. 要約

| 項目 | 内容 |
|------|------|
| 何をする | 06純正の `データを閲覧する` を全行抽出し、MAP右の SHELF と json に蓄積 |
| 照合 | マスタ／P4b の Browse Node ID → テンプレ有無＋採用PT |
| 缶飯 | テンプレ `FOOD_FISH_GROCERY.xlsm`／採用PT **GROCERY** |

---

## 1. 人間（テンプレ追加時）

1. SC から純正 xlsm を DL → **09**  
2. Agent に B-T0 指紋＋ **SHELF browse 抽出** を依頼  
3. 問題なければ **09 → 06**  
4. Drive の `shelf_registry.json`／`shelf_browse_catalog.json` が更新されていることを確認  
5. スプレッドシート `▼設定(Amazonマッピング)` 右の SHELF に行が増えていることを確認  

---

## 2. Agent コマンド（抽出）

```text
cd tools\c1_hpc_packaged
python c1_shelf_browse_extract.py --template-dir "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/06.純正テンプレ原本（読取専用・触らない）"
python sync_shelf_browse_to_map_sheet.py
```

特定ファイルのみ:

```text
python c1_shelf_browse_extract.py --xlsm "G:/マイドライブ/.../06/.../FOOD_FISH_GROCERY.xlsm" --allowed-pts FOOD,FISH,GROCERY --preferred-pt GROCERY
```

---

## 3. 缶飯通し（運用）— Phase D

**Agent側準備（済）**: SHELF 194行・Drive04 の `shelf_registry.json`（browseIndex 込み）同期・grocery C1マップ・GASコード。  
**人間必須**:

1. `clasp push`（`AmazonCategoryPt.js`／`コード.js`）  
2. Property `AMAZON_SHELF_REGISTRY_FILE_ID` が Drive04 の **更新済み** `shelf_registry.json` を指すこと（中身に `browseIndex` がある）  
3. `APPROVAL_AMAZON_LV4_ENABLED=true`。PT空なら `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED=true`（作業後 false）  
4. 冪等ブロック時は **D** で「失敗後の再GENERATED」レ点ON（または先に **21-④**）→ Amazonのみ・新規・在庫 ZERO  
5. 完了ダイアログ全文を Cursor に貼付 → Agent が `food_fish_grocery_column_map.json` で PACKAGED → 03  
6. SC UL。トグル戻し  

Agent への貼付後コマンド例（subBatch はダイアログの値）:

```text
cd tools\c1_hpc_packaged
python c1_fetch_inputs.py ...
python c1_packaged.py --mode prod --sub-batch （subBatchId）
```

（config の `column_map_path` を `food_fish_grocery_column_map.json` にするか、棚引き B-T1 経由）

---

## 4. 検収欄

| 項目 | 結果 | メモ |
|------|------|------|
| GROCERY 全行抽出 | **OK** | 160行＋SEASONING34＝計194 |
| MAP SHELF 差分反映 | **OK** | 列**AK**〜（RULES=AD／共通認識=AG の右。AA誤消去は修復済）。ヘッダは **行6英／行7日／行8説明（frozen）／行9〜データ** |
| Node→GROCERY 解決 | □ | clasp push 後に D／P4b で確認 |
| D 通し GENERATED | □ | 人間 |
| PACKAGED／SC | □ | Agent→人間 |

---

## 5. 後続（別承認・本手順外）

- **専用D修正メニュー**（作らない。失敗後再GENERATEDは [D内レ点承認](LV4_D_REMAKE_MENU_APPROVAL.md)）  
- 属性MAP（ハイライトB／HJ等）のライブ差分sync完了確認  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | 後続を「専用D修正メニュー」却下→D内レ点へ更新。 |
| 2026-08-02 | 起草。 |
