# 楽天マトリクス — サブ昇順詰め

**文書種別**: requirements  
**状態**: 2026-09-06 チャット確定・実装済  
**最終更新**: 2026-09-06  

**対象**: `rakutenMatrixAutobindByFilename_`（C「楽天のみ」）  
**関連**: [D_MENU_RAKUTEN_MATRIX_WORK_MERGE_SCOPE_REQUIREMENTS.md](D_MENU_RAKUTEN_MATRIX_WORK_MERGE_SCOPE_REQUIREMENTS.md)

---

## 1. 目的

サブ枠を欠番で空けず、**手元の `_sub*` を昇順に並べて列1〜10へ詰める**。「列番号＝ファイルの `_subN`」は必須にしない。

---

## 2. 確定仕様（`RAKUTEN_MATRIX_SUB_PACK_ASC_ENABLED`）

| 項目 | 内容 |
|------|------|
| 未設定 | **ON** |
| 対象 | 親SKU／JAN／子SKUで今回レ点親に紐づく `_subN`（または `_pN`）ファイル |
| ソート | **N昇順** → 同一Nは **ファイル名昇順**（同一Nは圧縮しない） |
| 配置 | 親行のサブ1→10へ連続詰め。L2で全子へ共有 |
| 11枚目以降 | 親の候補 |
| 番号なし | 親/JAN一致のみ。番号付き詰めの**あと**、空きサブをファイル名昇順で埋め、余りは候補。子SKU一致かつ番号なしは従来どおり MAIN1 |
| OFF | 旧仕様（`_subN` → サブN列固定） |

### 例

ファイルのサブ番号が `1,2,3,4,4,4,7,7,9,10` → 列1〜10にその順で入る。

---

## 3. 戻し

`RAKUTEN_MATRIX_SUB_PACK_ASC_ENABLED=false`

---

## 4. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-09-06 | 初版。詰め本線＋番号なし余り枠／候補。 |
