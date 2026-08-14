# サブ画像コース（楽天出口）— 承認パッケージ

**状態**: **実装承認**（プラン実装・2026-08-09）。実機は HUMAN_RUN（clasp push は人間）。  
**要件**: [D_MENU_SUB_IMAGE_RAKUTEN_COURSE_REQUIREMENTS.md](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_REQUIREMENTS.md)  
**手順**: [D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)  
**投入レイヤ**: [LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md](LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md)  
**三者**: **スキップ**

---

## 1. 目的

B-③採用→AI合成v3→`{子SKU}_subN`→楽天マトリクス自動→R-Cabinet まで本線化。  
人手は **サブ採用CK＋短い目視**のみ。Amazon は既存 **REUSE_RAKUTEN（U4）**。Yahoo は別チケット。

---

## 2. 変更予定ファイル

| ファイル | 概要 |
|----------|------|
| `tools/set_main_image/sub_image_b3_curate.py` | 参照themeId／phaseOrder／cluster 自動列 |
| `tools/set_main_image/export_sub_images_for_rakuten_matrix.py` | compose meta→subN 本線 export |
| `コード.js` | B-③右列ヘッダー拡張（レ点復元非破壊） |
| `docs/org/*`（本承認・要件・E2E HUMAN_RUN） | 正本 |
| `docs/AGENT_HANDOVER.md` / `CHANGE_LEDGER.md` / `CURRENT_PHASE.md` | 台帳 |
| PoC HUMAN_RUN | 本線への誘導リンク |

**触らない**: `generateRakutenCSV`／Yahoo API／Amazon PT 新規生成／Drive 破壊的整理。

---

## 3. 変更概要

1. B-③に LPテーマ分類由来の **自動参照列**（人手編集対象外）
2. compose `run_meta.json` から phaseOrder 順に `{子SKU}_subN.jpg` を export（`--pick a|b`）
3. 既存ファイル名自動投入＋R-Cabinet 経路で楽天へ。Amazon は U4 手順のみ

---

## 4. リスクと戻し方

| リスク | 緩和 |
|--------|------|
| B-③列増で GAS 復元ずれ | ヘッダー名で CK 列解決。余分列は空埋め |
| 誤 export | dry-run／目視後に Drive 配置 |
| 誤自動投入 | `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false` |

- Git revert → 必要なら clasp push
- Property OFF（マトリクス自動）

---

## 5. 必須3点セット

| 項目 | 内容 |
|------|------|
| docs | 要件・承認・E2E HUMAN_RUN・台帳 |
| ログ | curate／export／autobind の件数 |
| 復元 | Property OFF ＋ Git revert |
