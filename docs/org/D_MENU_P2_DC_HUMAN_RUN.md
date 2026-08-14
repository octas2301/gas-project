# Amazon P2 — Dc ①②③（人間手順・調査結論）

**状態**: **調査完了＋方針ロック**（2026-08-01。PoCコードは未）  
**承認**: [LV4_P2_DC_123_INVESTIGATION_APPROVAL.md](LV4_P2_DC_123_INVESTIGATION_APPROVAL.md)（**§7.4**）  
**過渡（現行本番）**: [D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md) §1e（21-⑮〜⑰）

---

## 0. 当面の正（ハイブリッド＋レーン）

```text
【レーンB・新規】
①【人手】Seller Central へ PACKAGED xlsm／画像ZIP UP
②【人手】processing-summary を Drive 08（監視）へ置く
③【GAS】21-⑮（ファイル名）※中身パースは低優先
→【人手】店頭ライブ・画像の最終確認（断定しない）
【レーンA・既存・本開発】
相乗り／更新は SP-API PUT 等を主に伸ばす（別チケット）
【レーンC】
新規JSONはゲート後のみ（承認§7.4）
```

**公式 API ではできないこと（現行 xlsm＝レーンB）**

- xlsm を SP-API で SC と同じ経路に自動 UP（①）  
- SC の `*-processing-summary.xlsm` を API で自動 DL（②）  

**やらない（今）**: レーンC本実装、②単独ポーリング、xlsm①②のAPI追従。

---

## 1. 調査記録（要約）

| Dc | 結論 |
|----|------|
| ① | PACKAGED xlsm API UP＝**不可**。人手継続（レーンB） |
| ② | SC processing-summary 自動DL＝**不可**。08手置き継続 |
| ③ | 21-⑮**温存**。強化＝任意・低優先 |
| 方針 | **A本線／B取り貯め／Cゲート後**（§7.4） |

詳細は承認包 §7・§7.4。

---

## 2. やってはいけない

- 非公式の SC ログイン自動化  
- 21-⑮の削除  
- ④⑤・自動ループの実装（P3）  
- レーンCのゲート前本実装／xlsm①②のAPI追従  
- 「UPLOADED_OK＝掲載完了」と断定  

---

## 3. 次

1. ~~レーンA拡大~~（A1〜A3・デュアル **済**）／七味ライブ（運用継続）  
2. ~~レーンBのエラー台帳~~ **初版済** — [台帳](LANE_B_SC_ERROR_LEDGER.md)／[手順](D_MENU_LANE_B_LEDGER_HUMAN_RUN.md)  
3. （低優先）③強化 PoC — 実装承認後に手順追記  

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 下書き。 |
| 2026-08-01 | 調査結論反映（ハイブリッド当面の正）。 |
| 2026-08-01 | レーンB台帳へリンク（§3）。 |
| 2026-08-01 | 方針ロック（A/B/C）反映。 |
