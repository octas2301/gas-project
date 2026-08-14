# レーンA3 — 自己発マスタ在庫>0（人間手順）

**状態**: **dry／prod 検収OK**（2026-08-01）。トグル戻しは作業後必須  
**承認**: [LV4_LANE_A3_MASTER_QTY_APPROVAL.md](LV4_LANE_A3_MASTER_QTY_APPROVAL.md)  
**親**: [D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md) §1  

---

## 0. 目的

相乗り **自己発**で、マスタ「在庫数」生値を quantity として dry_run → prod する（承認②相当）。

---

## 1〜4. 手順要約

自己発＋マスタ在庫＋1SKU → dry_run → prod → トグル false。詳細は承認包。

---

## 5. 合格記録（2026-08-01）

| 段階 | 結果 | 記録 |
|------|------|------|
| A3-a dry | **OK** | `SPAPI_PUT_OFFER_CK_DRY_20260801_121701_49a49e` VALID／MASTER／`…19as13`／行503 |
| A3-b prod | **OK** | `SPAPI_PUT_OFFER_CK_PROD_20260801_121813_f677a3` ACCEPTED |
| A3-c トグル | □ | `ENABLED`／`ALLOW_PROD`／`ALLOW_MASTER_QTY` → **false**（作業後） |
| マスタ在庫数不変 | □ | 人間目視 |

---

## 6. やってはいけない

- FBAで quantity 送信を期待する  
- `ALLOW_MASTER_QTY` を true のまま放置  

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草。実機待ち。 |
| 2026-08-01 | dry／prod OK。トグルは人間。 |
