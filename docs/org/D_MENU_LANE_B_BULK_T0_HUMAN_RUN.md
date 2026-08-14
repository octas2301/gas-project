# レーンB — 純正バルク B-T0（指紋／差分）人間手順

**状態**: **実装済／スモーク合格**（2026-08-01）  
**承認**: [LV4_LANE_B_BULK_TEMPLATE_T0_APPROVAL.md](LV4_LANE_B_BULK_TEMPLATE_T0_APPROVAL.md)  
**実行**: **ローカル Python は Agent モード**

---

## 0. 置き場

根: `G:\マイドライブ\04.amazonカタログ作成（CSV一括UL）\`

| 役割 | パス |
|------|------|
| **入口** | `…\09.SC純正バルクxlsm保存（人間がDLして保存→Agent確認後06へ）\` |
| **原本** | `…\06.純正テンプレ原本（読取専用・触らない）\` |
| **レポート** | `…\05.SC処理結果・ログ退避（人間）\` |

---

## 1. 人間

1. SC から純正 xlsm を DL → **09** へ保存  
2. Agent に「B-T0 指紋」を依頼  
3. **05** の `B_T0_*_SUMMARY.txt`／`*_FINGERPRINT.json` を確認  
4. 問題なければ **09 → 06** へコピー／移動  

---

## 2. Agent

```text
cd tools\c1_hpc_packaged
python c1_bulk_fingerprint.py
```

特定ファイル:

```text
python c1_bulk_fingerprint.py --inbox "G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/09.SC純正バルクxlsm保存（人間がDLして保存→Agent確認後06へ）/（ファイル）.xlsm"
```

---

## 3. 合格記録

| 段階 | 結果 | メモ |
|------|------|------|
| 09作成 | **OK** | 2026-08-01 |
| 三点スキップ／実装承認 | **OK** | 2026-08-01 |
| スモーク | **OK** | `status=match` SEASONING／sha `57190dbc…`／`B_T0_20260801_081134_…`（05） |

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草→09確定→**実装＋スモーク**。 |
