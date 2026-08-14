# レーンB — B-T1 prod 第2段（項目名充填の本番 xlsm 出力）承認パッケージ

**状態**: **実装済／スモーク合格**（2026-08-01）。三点スキップ。SC UL は人間  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)（レーンB）  
**親**: [LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md)（**dry_run 実装済**）  
**前提済**: B-T0／B-T1 dry_run  
**手順**: [D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md](D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md)  
**三者レビュー**: **スキップ**（社長明示 2026-08-01）

---

## 1. 目的

B-T1 dry_run と同じ経路で、**本番向け PACKAGED xlsm**（`_DRYRUN` なし）を **03** に新規出力する。

| 段階 | 内容 | 状態 |
|------|------|------|
| B-T0 | 指紋／差分 | **済** |
| B-T1 dry_run | 棚引き＋項目名＋`*_DRYRUN.xlsm` | **済** |
| **B-T1 prod（本包）** | 同経路で **prod xlsm** | **実装済／スモークOK** |
| B-T2 | 新PT／別複合／単体 | **方針ロック済・実装は実需後** — [T2承認](LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md) |

---

## 2〜3. 背景・置き場

B-T1 継承。根 `04`。06読取専用／03新規のみ／05ログ／SC ULは人間。

---

## 4. 社長確定方針（**ロック済 2026-08-01**）

| # | 論点 | **決定** |
|---|------|----------|
| 1 | スコープ | B-T1 入口の prod 出力のみ |
| 2 | PT／テンプレ | 当面 **SEASONING＋複合** |
| 3 | 入口 | `c1_bulk_fill_by_name.py --mode prod`（既定 dry_run） |
| 4 | 棚 | prod は **`DL_NOT_NEEDED` 必須**。`--allow-dl-required` は prod 拒否 |
| 5 | 指紋 | **一致必須**（不一致→非作成） |
| 6 | 項目名 | 必須 MISS → 中止 |
| 7 | accepted=0 | prod 中止 |
| 8 | 出力名 | `_DRYRUN` **なし** |
| 9 | 03 | 上書き禁止（衝突時タイムスタンプ） |
| 10 | 聖域 | 06非破壊・マスタ非書込・SC非自動UL・楽天／Yahoo／B統合非触 |
| 11 | 既存 C1 | 併存可 |
| 12 | 検収 | prod 1親＋ゲート拒否スモーク |
| 13 | 三点 | **スキップ** |

---

## 5. 実装ファイル

| パス | 内容 |
|------|------|
| `tools/c1_hpc_packaged/c1_bulk_fill_by_name.py` | `--mode dry_run\|prod`・prod 棚ゲート |
| docs | 本承認／HUMAN_RUN／PHASE 等 |

**戻し**: git revert または `--mode` 既定のみ使用。

---

## 6. リスク（要約）

誤 prod → 明示フラグ＋棚／指紋ハードゲート。詳細は起草時§6。

---

## 7. 検収

- [x] §4 ロック… **2026-08-01**  
- [x] 三点スキップ… **2026-08-01**  
- [x] 実装承認… **2026-08-01**  
- [x] prod 棚あり・指紋一致 → `_DRYRUN` なし… `…_oya_20260801_085442.xlsm`／`B_T1_FILL_PROD_SEASONING_20260801_085431`  
- [x] 指紋不一致 → 非作成… exit 3／`B_T1_FILL_PROD_SEASONING_20260801_085444`  
- [x] 棚なし＋prod → 拒否… HERB exit 2  
- [x] 06／マスタ非破壊  
- [x] docs  

---

## 8. 社長確認

- [x] §4  
- [x] 三点スキップ  
- [x] 実装承認  

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。 |
| 2026-08-01 | **§4ロック／三点スキップ／実装＋スモーク**。 |
