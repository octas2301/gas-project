# LV4 — ④⑤⑥ SP-APIリサーチ／3D箱フィット 実装承認

**状態**: **④自己発3D 実装済**／**FBA P1a 診断実装済（要 clasp push）**／P1b本線・⑤⑥は後続  
**三点**: **スキップ**（社長方針合意済）  
**正本**: [B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md](B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md)  
**手順**: [B_SPAPI_RESEARCH_P0_P2_HUMAN_RUN.md](B_SPAPI_RESEARCH_P0_P2_HUMAN_RUN.md)

---

## 1. 目的

④自己発箱（3Dフィット）を先に本線化。FBAは **P1a＝診断のみ** → 合格後 P1b。⑤競合ASIN・⑥メーカー品番は後続。

---

## 2. 社長確定方針（ロック）

| # | 決定 |
|---|------|
| 1 | 実装順 **④自己発3D → FBA(Catalog) → ⑤ → ⑥** |
| 2 | 非定型箱は **exclude**（選定しない） |
| 3 | 向き自由・N≤40・1箱 |
| 4 | ⑤は人間◎が上限（後続） |
| 5 | 3者レビュー不要 |
| 6 | FBA P1a はマスタ `サイズ＆自己発/FBA` を書かない |

---

## 3. 変更（④）

| パス | 内容 |
|------|------|
| `00_設定マスタ` E〜H | 内寸・種別。exclude 2行 |
| `コード.js` | `getMaterialTableFromSettings` 拡張／3Dフィット／Step3.1接続 |
| `tools/c1_hpc_packaged/write_box_inner_dims.py` | 内寸書込 |
| `tools/c1_hpc_packaged/mark_box_exclude.py` | exclude マーク |
| docs | 要件／HUMAN_RUN／本承認 |

**戻し**: `B_LOGISTICS_USE_3D_FIT=false`／git revert

---

## 3b. 変更（FBA P1a）

| パス | 内容 |
|------|------|
| `コード.js` | Z 15-⑰／`menuFbaTierCatalogDiagnoseP1aForCheckedParents`／仮ティア／診断シート |
| docs | 要件 §2.4・HUMAN_RUN U4・本承認 |

**戻し**: `B_FBA_P1A_DIAG_ENABLED=false`／git revert  
**リスク**: Catalog読取のみ（書込なし）。誤ティアは診断シート限定。

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-13 | FBA P1a 診断実装（15-⑰・マスタ非書込） |
| 2026-08-13 | ④3D実装。exclude。優先順明記 |
| 2026-08-13 | 初版要件 |
