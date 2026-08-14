# SHELF Browse 網羅カタログ — 承認パッケージ

**状態**: **方針ロック／実装済**（2026-08-02・抽出194行・MAP AA列・P4b/D Node ルーティング）  
**三点**: **スキップ**（社長明示・既存レーンB方針に準拠）  
**手順**: [D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md](D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md)  
**関連**: [LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md)／[LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md](LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／C1 grocery・seasoning 列マップ

---

## 1. 目的

純正 xlsm の **`データを閲覧する`** にある Browse Node 一覧を網羅蓄積し、

1. P4b／Catalog が返した **Browse Node ID** が、取得済みテンプレのどれかに載るか確定する  
2. 載っていれば **テンプレファイル＋選べる商品タイプ（allowed）＋採用PT（preferred）** を紐付ける  
3. 機械は **browse でルーティング → 採用PT＋テンプレ指紋で充填可否を検証** する二段キーにする  

缶飯事例: Catalog PT=`MEAT`（xlsmに無し）／browse=`肉の缶詰・瓶詰 (71192051)`（`FOOD_FISH_GROCERY` の一覧に有り）→ 採用PT=`GROCERY`。

---

## 2. 矛盾解消（却下した案）

| ID | 却下 | 正（ロック） |
|----|------|--------------|
| X1 | 缶飯本線PT=`FOOD` 固定 | **採用PT=`GROCERY`**（同一テンプレに FOOD/FISH/GROCERY。MEATなし） |
| X2 | SHELFに代表ノードだけ点登録 | **`データを閲覧する` 全行網羅**（点登録禁止） |
| X3 | browse 名の曖昧一致だけでゲートOK | 照合は **browseNodeId**。最終検証は **PT＋指紋** |
| X4 | マスタPTを人間が常時手修正 | PT/browseは **P4b／Dゲート／メニュー**。人間手編集は商品名等のみ |
| X5 | 破壊的 recreate でMAP全書換 | **差分更新のみ** |
| X6 | MEAT→GROCERY 固定を永続本線 | **過渡フォールバック**。SHELF Node 引上が本線 |

---

## 3. 社長確定方針（ロック 2026-08-02）

| # | 論点 | 決定 |
|---|------|------|
| 1 | ルーティング正 | **Browse Node ID** が SHELF 網羅表に存在するか |
| 2 | 充填・指紋正 | 選んだ **productType ＋ fingerprintSha256**（従来 registry entries） |
| 3 | 抽出元 | 各06純正のシート **`データを閲覧する`**（列 `Browse Node` / `BrowsePath`） |
| 4 | MAP置き場 | `▼設定(Amazonマッピング)` 内必須。共通認識の右（列AK〜）。`=== SHELF ===` は行4で MAP と揃え |
| 4b | SHELFヘッダ | **行6=英語項目名／行7=日本語訳／行8=項目説明（リストheader固定・frozen）／行9〜=データ**。目的・説明の別ブロックは置かない |
| 5 | 列（明細） | `templateFile` / `templateUrl` / `allowedProductTypes` / `preferredProductType` / `browseNodeId` / `browsePath` / `fingerprintSha` / `columnMapPath` / `extractedAt` / `sourceSheet` |
| 6 | allowedProductTypes | その xlsm の商品タイプ候補（例: `FOOD,FISH,GROCERY`） |
| 7 | preferredProductType | 候補から1つ。未設定時はテンプレ既定（缶飯複合の既定=`GROCERY`） |
| 8 | 機械正本 | `shelf_browse_catalog.json`（リポジトリ）＋ Drive 同期。シートは鏡 |
| 9 | registry | `shelf_registry.json` の PT＋指紋エントリは維持。browse 索引は catalog から引く |
| 10 | P4b | Catalog の browse から Node ID 抽出 → catalog ヒットなら **preferredPT** をマスタへ（Catalog PT 名より優先）。ミスなら書込せず停止理由 |
| 11 | Dゲート | browse Node が catalog 未登録 → **停止＋DL指示**。登録済なら採用PTをマスタへ寄せ、従来どおり PT∈registry を確認 |
| 12 | 過渡エイリアス | `MEAT→GROCERY` は catalog ミス時のフォールバックのみ |
| 13 | C1 | grocery: `food_fish_grocery_column_map.json`。seasoning: 既存。PT/browse 必須・既定browse禁止・ハイライトB・HJ型番 |
| 14 | 聖域 | 楽天／Yahoo／B統合非触。06破壊禁止。SC自動ULなし。GASは xlsm 非編集 |
| 15 | 後続 | **D内レ点**（失敗後再GENERATED・案A）は [LV4_D_REMAKE_MENU_APPROVAL.md](LV4_D_REMAKE_MENU_APPROVAL.md)。属性MAPは [MAP sync](LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md) |
| 16 | 三点 | スキップ |

---

## 4. C1／缶飯ロック（会話統合）

| 項目 | 値 |
|------|-----|
| 純正 | `FOOD_FISH_GROCERY.xlsm` |
| 指紋 | `74ccdcf96c22879dc80cbe87e8b41aa615e923529f151f091552ccbe3cefb010`（行3–5・maxCol310） |
| 採用PT | **GROCERY** |
| C1マップ | `tools/c1_hpc_packaged/food_fish_grocery_column_map.json` |
| PT/browse | マスタ両方必須。空＝親除外 |
| ハイライト | タイトル≤75: 楽天キャッチ→Yahooキャッチ→箇条書き① |
| 型番 | HJ `メーカー型番` → 品番 → GENERATED → 子SKU |
| 既定 | テーマ`サイズ`（GROCERY純正プルダウンのみ。`SET_NAME`不可）／入数・ユニット＝セット缶数／重量＝サイズからg／色・形態・温度は未出力／KW1枠 |

役割: 人間＝スプシ専用項目＋メニュー＋SC。Agent＝抽出／PACKAGED／MAP差分sync。GAS＝xlsm非編集。

---

## 5. 変更ファイル（実装）

| パス | 内容 |
|------|------|
| `tools/c1_hpc_packaged/c1_shelf_browse_extract.py` | 06 xlsm → browse 全行抽出 |
| `tools/c1_hpc_packaged/shelf_browse_catalog.json` | 機械正本（網羅） |
| `tools/c1_hpc_packaged/sync_shelf_browse_to_map_sheet.py` | MAP SHELF 差分書込 |
| `tools/c1_hpc_packaged/shelf_registry.json` | メタ参照（既存 PT エントリ維持） |
| `AmazonCategoryPt.js` | Node ID → catalog → preferredPT |
| `コード.js` | Dゲート browse 未登録停止／採用PT寄せ |
| 本承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／T2／P4b | docs |

**戻し**: git revert。Drive の catalog／registry を旧版に戻す。Property 変更なし（既存 `AMAZON_SHELF_REGISTRY_FILE_ID` 流用可。catalog はリポジトリ＋Drive同梱または registry 隣接ファイル）。

---

## 6. 検収

- [x] FOOD_FISH_GROCERY 全行抽出（160）＋ MAP 列AA  
- [x] SEASONING 複合抽出（34）・計194  
- [x] Node `71192051` → preferred GROCERY（catalog）  
- [ ] clasp push 後: Catalog MEAT でもマスタが GROCERY になる（D／P4b）  
- [ ] catalog ミスで D 停止＋DL指示（実機）  
- [x] C1 grocery: PT/browse require（コード済）  
- [x] 楽天聖域非触  

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | 起草・方針ロック（網羅SHELF・browseルーティング・缶飯GROCERY・過渡MEATエイリアス）。 |
| 2026-08-02 | 実装: 抽出194行・MAP AA・registry browseIndex・P4b/D Node ルーティング。要 clasp push。 |
