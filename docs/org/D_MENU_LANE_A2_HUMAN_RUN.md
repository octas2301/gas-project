# レーンA2 — 運用固め（人間手順）

**状態**: **検収OK**（2026-08-01）  
**承認**: [LV4_LANE_A2_OPS_HARDENING_APPROVAL.md](LV4_LANE_A2_OPS_HARDENING_APPROVAL.md)  
**親**: [D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)

---

## 0. 目的

1. 確認キャンセルで PUT が走らないこと  
2. 作業後トグルを安全側に戻せること  
3. D 本線以外の逃げ道メニューに届くこと  

---

## 1〜3. 手順要約

詳細は承認包。要点: dry_run 確認でキャンセル／Property false／D補助または Z-21 到達。

---

## 4. 合格記録（2026-08-01）

| 段階 | 結果 | 記録 |
|------|------|------|
| A2-a キャンセル | **OK** | 12:07:21 `cancelled_by_user`。PUTなし。SKU例 `sanky-B01N5A6ESU-19as13` |
| A2-b トグル | **OK** | 社長宣言「A2実機OK」（ENABLED／ALLOW_PROD／MASTER_QTY を false に戻す） |
| A2-c 逃げ道 | **OK** | `menuAmazonCourseE0Precheck` DONE checked=1 warn=0（12:08:42）。UIキャンセル＝書込なし |

---

## 5. やってはいけない

- キャンセル試験を `ALLOW_PROD=true`＋prod 選択のまま「運任せ」でやる  
- A2 スコープでマスタ在庫（A3）を試す  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。実機待ち。 |
| 2026-08-01 | **検収OK**。 |
