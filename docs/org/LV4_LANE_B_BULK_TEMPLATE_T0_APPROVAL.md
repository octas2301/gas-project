# レーンB / C1-次 — 純正バルク即活用 B-T0（取込＋指紋／差分）承認パッケージ

**状態**: **コード実装済／スモーク合格**（2026-08-01）。三点スキップ。B-T1は別承認  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)  
**親**: [D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)／[LV4_P2_DC_123_INVESTIGATION_APPROVAL.md](LV4_P2_DC_123_INVESTIGATION_APPROVAL.md) §7.4／[LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) §2  
**前提済**: C1 HPC／SEASONING／[P4b](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)  
**手順**: [D_MENU_LANE_B_BULK_T0_HUMAN_RUN.md](D_MENU_LANE_B_BULK_T0_HUMAN_RUN.md)  
**ツール**: `tools/c1_hpc_packaged/c1_bulk_fingerprint.py`  
**三者レビュー**: **スキップ**（社長明示 2026-08-01）

---

## 1. 目的

SC から人間がダウンロードした **新しい純正バルク（xlsm）** をパイプラインに載せる第一歩（指紋／差分のみ）。

| 段階 | 内容 | 状態 |
|------|------|------|
| **B-T0** | 取込＋指紋＋差分レポート（書込なし） | **実装済／スモークOK** |
| **B-T1** | 項目名マッピング充填＋不足一覧 | **実装済** — [T1承認](LV4_LANE_B_BULK_TEMPLATE_T1_APPROVAL.md) |
| **B-T2** | 新PT／別複合／単体本線化 | **方針ロック済・実装は実需後** — [T2承認](LV4_LANE_B_BULK_TEMPLATE_T2_APPROVAL.md) |

---

## 2〜3. Drive 置き場

**根**: `G:\マイドライブ\04.amazonカタログ作成（CSV一括UL）\`

| サブ | 役割 |
|------|------|
| **09.SC純正バルクxlsm保存（人間がDLして保存→Agent確認後06へ）** | **入口**（未検証） |
| **06.純正テンプレ原本（読取専用・触らない）** | 確認後の運用原本 |
| **05.SC処理結果・ログ退避（人間）** | T0 レポート出口 |
| **03.…** | PACKAGED（B-T1以降） |

---

## 4. 社長確定方針

| # | 論点 | 決定 |
|---|------|------|
| 1 | スコープ | **B-T0＝指紋＋差分のみ** |
| 2 | 置き場根 | **`04` 集約** |
| 3 | 入口 | **`09.SC純正バルクxlsm保存（人間がDLして保存→Agent確認後06へ）`** |
| 4 | 合格後 | **09 → 06** |
| 5 | レポート | **`05`** |
| 6〜11 | 比較先・Agent実行・GAS非触・マスタ非読・API非DL | ロックどおり |
| 12 | 三点 | **スキップ**（2026-08-01） |

---

## 5. 実装ファイル

| パス | 内容 |
|------|------|
| `tools/c1_hpc_packaged/c1_bulk_fingerprint.py` | 09→指紋→05 レポート |
| docs | 本承認／HUMAN_RUN／PHASE 等 |

**戻し**: ツール削除／git revert。

---

## 6. 仕様（実装）

```text
python c1_bulk_fingerprint.py
# または
python c1_bulk_fingerprint.py --inbox "…/09…/file.xlsm" --report-dir "…/05…"
```

- 既定入口＝09／既定出口＝05／比較＝`fingerprints/*.json`  
- status: `match`／`unknown_template`／`FAILED`  
- **データ行・マスタ・06 を書かない**

---

## 8. 検収

- [x] §4 ロック（置き場・スコープ・三点スキップ・実装承認）… **2026-08-01**  
- [x] `09` 作成… **人間済**  
- [x] 実装… `c1_bulk_fingerprint.py`  
- [x] スモーク… SEASONING 一致 `status=match` sha=`57190dbc…` → `05\B_T0_20260801_081134_…`  
- [x] docs  

---

## 9. 社長確認

- [x] 09 名… **2026-08-01**  
- [x] B-T0＝指紋のみ… **2026-08-01**  
- [x] 三点スキップ… **2026-08-01**  
- [x] 実装承認… **2026-08-01**  
- [x] スモーク… **2026-08-01**  

---

## 10. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草→09確定。 |
| 2026-08-01 | **実装承認＋三点スキップ＋コード**。スモーク match（SEASONING）。 |
