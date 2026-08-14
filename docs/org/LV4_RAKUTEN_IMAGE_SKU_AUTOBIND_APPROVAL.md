# 楽天画像 SKU 自動紐付け — 承認パッケージ

**状態**: **実装承認**（プラン実装・2026-08-09）。実機検収は HUMAN_RUN。  
**要件**: [D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)  
**手順**: [D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)（楽天 MAIN/サブ自動）／[D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)  
**三者**: **スキップ**（Amazon U2-ε と同型の転用）

---

## 1. 目的

Amazon U2-ε（ファイル名に子SKU → MAIN 自動投入）と同型を楽天マッチング sheet に転用する。  
**セット数はマスタ列で子SKUを決めてからファイル名に載せる**。画像からセット数を Vision 推定しない（本フェーズ）。

---

## 2. 変更予定ファイル

| ファイル | 概要 |
|----------|------|
| `コード.js` | `generateAiImageMatrix` にファイル名一致の楽天 MAIN/サブ自動投入。Property トグル |
| `tools/set_main_image/master_sets.py` | セット数→子SKU解決ヘルパ |
| `tools/set_main_image/rakuten_image_names.py` | 命名契約（MAIN/サブ） |
| `tools/set_main_image/export_sub_images_for_rakuten_matrix.py` | サブ出力を `{子SKU}_subN.jpg` へコピー |
| `docs/org/*`（本承認・要件・HUMAN_RUN・C追記） | 正本・手順 |
| `docs/AGENT_HANDOVER.md` / `CHANGE_LEDGER.md` / `CURRENT_PHASE.md` | 台帳 |

**触らない**: `generateRakutenCSV`／`Yahoo.js`／Amazon U2 本線の破壊的変更／アップ時リネーム規約の作り直し。

---

## 3. 変更概要

1. Drive 楽天ソースの画像をマトリクス生成時、**ファイル名に子SKU**があれば当該子行の空メイン1（または `_subN` → サブN）へ自動投入。既存非上書き。親SKUのみ一致は候補のみ。
2. Script Property `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED`（未設定=true）。`false` で旧・AI仕分け／手ドラッグ中心。
3. 生成ツールは既存 `{子SKU}_rakuten.jpg` を本線とし、サブは `{子SKU}_subN.jpg` 契約を明示・エクスポート補助。

---

## 4. 想定リスク

| リスク | 緩和 |
|--------|------|
| 短いSKU断片の誤一致 | 最長一致優先（Amazon ε 同） |
| 複数ファイルが同一子 | MAIN1は1枚・余りは候補。ログ |
| 既存 Vision 仕分けとの競合 | **ファイル名パスを先**に実行。使用済みファイルは AI ループから除外 |
| 誤自動投入 | 既存枠は非上書き。Property OFF |

---

## 5. 戻し方

- Property: `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false`
- Git: 当該コミット／ファイル revert → `clasp push`

---

## 6. 必須3点セット

| 項目 | 内容 |
|------|------|
| docs | 要件・本承認・HUMAN_RUN・HANDOVER／LEDGER |
| 調査ログ | `runId`・autobind 件数・skip 理由 |
| 復元 | Property OFF ＋ Git revert |
