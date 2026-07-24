# Dメニュー Amazon 統合 — 三点レビュー多数決メモ

**文書種別**: 検証記録（多数決）  
**日付**: 2026-07-24（反映 2026-07-25）  
**対象正本**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)  
**追加必読**: [AI_ORG_CHARTER.md](AI_ORG_CHARTER.md) ・ [AI_APPROVAL_MATRIX.md](AI_APPROVAL_MATRIX.md) ・ [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §10・§11.0 ・ [LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md) §2.1・§7 ・ [CURRENT_PHASE.md](../CURRENT_PHASE.md) ・ `AGENTS.md` §2/§5/§8  
**関連**: [THREE_REVIEW_RUNBOOK.md](THREE_REVIEW_RUNBOOK.md)  
**ルール**: 2/3以上＝採用、1票＝未決、社長が最終ゲート  
**mode**: 正式（親1＋並列サブ3・別モデル）  
**docs書き込み**: 本メモ＋要件反映（コード実装なし）

---

## 1. レビューア

| ID | モデル | 結論 |
|----|--------|------|
| Reviewer-1 | Claude Opus 4.8 | 条件付き YES |
| Reviewer-2 | GPT-5.6 Terra | 条件付き |
| Reviewer-3 | Composer | 条件付き YES |

総合: **条件付き YES**（採用項を正本反映後、U0クローズ。U3は別実装承認）。

---

## 2. スコア（参考）

| 観点 | R1 | R2 | R3 |
|------|----|----|-----|
| 実現性 | 4 | 4 | 4 |
| 要件漏れ | 3 | 3 | 3 |
| 矛盾 | 3 | 2 | 4 |
| 聖域 | 4 | 5 | 5 |
| 逃げ漏れ | 3 | 3 | 3 |

---

## 3. 採用（2/3以上）→ 正本反映済

| # | 内容 | 票 |
|---|------|----|
| A | Daフローに **PACKAGED** を明示。「SCの2手」＝ xlsm UP ＋ ZIP UP のみ。承認①・C・21-③・Propertyは別チェック | 3/3 |
| B | **T3**: 手ZIP＝当面の正／T3＝自動化の実装待ち。§11.0 の R2 URL→18320 を **T2だけでは足りない既存証跡**とする（社長確認） | 2/3＋社長 |
| C | **U3薄いファサード契約**: Dは既存21呼出＋表示のみ。GENERATED本体再実装禁止 | 2/3 |
| D | Da完了ダイアログ最低仕様（PACKAGED場所・手ZIP・21-③＝Z 等） | 2/3 |

### 社長確定（2026-07-25）

1. T3文言: **「T2で対応できなくなったことが確定したことをもって」** という前提なら可 → §11.0 18320をその既存証跡とする。  
2. Da人間作業: **当面は「PACKAGED作成＋SCの2手」明示可**。将来は **API連携必須**（Dc強化）。  
3. 本多数決メモの docs 配置: **可**。

---

## 4. 未決（残す）

- REUSE／AMAZON_ONLY の選択単位と0枚フォールバック（U2）  
- Dから21-③を許す時期  
- 複合モール（楽天＋Amazon）の詳細UX  
- T3実装着手の最終確認（証跡は既存。実装承認は別）  

---

## 5. 総合

**条件付き YES → 採用反映で U0 クローズ可。**  
次: U2（任意）／U3 実装は変更ファイル一覧を出して別承認。コードなし。
