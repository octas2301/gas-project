# レーンB — SCエラー／成功台帳

**正本**: 本ファイル  
**承認**: [LV4_LANE_B_ERROR_LEDGER_APPROVAL.md](LV4_LANE_B_ERROR_LEDGER_APPROVAL.md)  
**手順**: [D_MENU_LANE_B_LEDGER_HUMAN_RUN.md](D_MENU_LANE_B_LEDGER_HUMAN_RUN.md)  
**最終更新**: 2026-08-10（七味SUCCESS／缶飯SUCCESS・人間SC確認／常時ONセット）

記入ルール: 1行＝1回のSC結果（または明確な中間失敗）。フルサマリ本文・秘密は載せない。新しい行は**表の先頭（直近）**に追加。

---

## 台帳

| 日付 | PT／ブラウズ | 親SKU／子SKU／ASIN | xlsm・サマリ（ファイル名） | 結果 | コード | 要約 | 対策・再発防止 | 参照 |
|------|-------------|-------------------|---------------------------|------|--------|------|----------------|------|
| 2026-08-10 | GROCERY／缶飯 | 複数親（例 `…2288018-oya`／`…2202019-oya`／`…2285127-oya`）＋相乗り | （SC在庫管理） | **SUCCESS**（ライブ） | — | 人間確認: 出品中。A1クローズ | 再xlsm不要。台帳クローズ | CURRENT_PHASE §0 |
| 2026-08-02 | GROCERY／缶飯 | 親`sanky-4538872281013-oya`＋子12 | `…052208.xlsm` → `…052208-processing-summary`（05） | **PARTIAL**（その他エラーで完了） | **100521**のみ | 処理13／成功0／その他13／失敗0。browse・数量OK。審査中（最大48h） | → 2026-08-10 SUCCESS | [MAP_SC_ERROR_LEDGER](MAP_SC_ERROR_LEDGER.md) |
| 2026-08-02 | SEASONING／唐辛子 | 親`sanky-4538872180149-oya`／子`…19as13`／ASIN`B01N5A6ESU`・親`B0HC0S6PRN`系 | （UL#再送済） | **SUCCESS**（ライブ） | — | SC在庫で**出品中**確認（在庫50・自己発）。100521クリア | KW1枠維持。台帳クローズ | [C1 §0b](D_MENU_C1_HUMAN_RUN.md) |
| 2026-08-02 | GROCERY／缶飯 | 同上 | `…044503.xlsm` → `…044503-processing-summary`（05） | **FAIL** | **90194/90225** | 処理13／成功0。browseが短縮名 | プルダウン一致の `BrowsePath (NodeId)` →`…052208` で解消確認 | [MAP_SC_ERROR_LEDGER](MAP_SC_ERROR_LEDGER.md) |
| 2026-07-31 | SEASONING／唐辛子 | 親`sanky-4538872180149-oya`／`B0HC9S8PRN`・子`sanky-B01N5A6ESU-19s13`／`B0HC9RRCBP` | `CK_…_032459.xlsm` → `…032459-processing-summary.xlsm`（Downloads） | **PARTIAL**（その他エラーで完了） | **100521**のみ | 処理2／成功0／その他エラー2／失敗0。**99016なし**。審査中だった | → 2026-08-02 SUCCESS | [C1 §0b](D_MENU_C1_HUMAN_RUN.md)／[D_ENTRY §1f](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md) |
| 2026-07-31 | SEASONING／唐辛子 | 同上 | 初回PACKAGED（`…235409`系）→サマリ | ERROR | **99016**／**100521** | KW5枠＋審査。処理2／成功0／その他エラー2 | KWは**1枠**。粉末／グラム既定。列マップ更新済 | C1 §0b／[SEASONING列マップ](D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md) |
| 2026-07-27 | HPC（安眠） | 親`lifec-4560151300924-oya`／子`…19s124` | Downloads `…-processing-summary.xlsm` | **SUCCESS** | — | 処理2／成功2／エラー0。E-5済 | 成功パターンとして再現可 | [C1 §7](D_MENU_C1_HUMAN_RUN.md) |
| 2026-07-26 | HPC（relax） | 親`lifec-4560151300405-oya`／子`…16s184` | Drive05 `…-processing-summary.xlsm` | **SUCCESS** | — | 処理2／成功2／エラー0 | 同上 | C1 §7 |

---

## よくあるコード（索引）

| コード | 意味（社内要約） | 典型対策 |
|--------|------------------|----------|
| **99016** | generic_keyword 等の枠超過 | SEASONINGは検索KW **1枠**（空白結合）。七味再送で**解消確認済** |
| **100521** | 審査／その他（最大48h・追加情報の要否判断） | 再xlsmよりSC側の審査・停止理由確認。公開されれば台帳に SUCCESS 追記 |

---

## 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-10 | **缶飯 SUCCESS**（人間SC出品中確認）。A1クローズ。 |
| 2026-08-02 | **七味 SUCCESS**（出品中・48h待ち解消）。缶飯 UL#3＝100521のみ。 |
| 2026-08-02 | **缶飯 UL#3**: `…052208`＝100521のみ（browse解消・再xlsm不要）。 |
| 2026-08-01 | 初版。七味エラー＋HPC成功2件をシード。 |
| 2026-08-01 | **七味再送サマリ確定**: `…032459`＝100521のみ（99016解消）。 |
