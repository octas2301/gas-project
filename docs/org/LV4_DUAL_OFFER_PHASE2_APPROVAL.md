# デュアルオファー Phase2 — 1実行で自己発＋FBA（承認・要件）

**日付**: 2026-08-01  
**状態**: **検収OK**（両系統 dry／prod・2026-08-01）。三点スキップ  
**親**: [LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md](LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)（Phase1検収OK）  
**手順**: [D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md](D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md)  
**三者レビュー**: **スキップ**（条件＝Dのみ／順次自己発→FBA／系統独立保存／MAX＝行数）

---

## 1. 目的

D の1回の実行で、同一レ点行に対し **自己発PUTと FBA PUT を最大2本**出す。列契約は Phase1 のまま。

---

## 2. 社長確定方針（2026-08-01）

| # | 論点 | 決定 |
|---|------|------|
| 1 | UI | 自己発／FBAを **チェック複数選択**（1つ以上必須） |
| 2 | 実行順 | **自己発 → FBA**（固定） |
| 3 | 実装 | **最小**: ファサードが系統ごとに既存 PUT を順次呼出 |
| 4 | 部分成功 | 可。全体 ok＝選んだ系統がすべて成功のときのみ |
| 5 | dry_run保存 | 系統独立（VALIDな系統の列だけ保存） |
| 6 | MAX_ITEMS | **レ点行数**。両系統時は API 最大2×を確認文言に明記 |
| 7 | MASTER | 自己発のみ quantity。FBA非送信 |
| 8 | 範囲 | **Dのみ** |
| 9 | 三点 | **スキップ** |

---

## 3. 変更ファイル

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `コード.js` | チェックUI・`mfn,fba` 引数・順次呼出・部分成功ダイアログ |
| 改修 | `AmazonSpapiPut.js` | fulfillments 解析ヘルパ（任意） |
| 新規 | 本ファイル／HUMAN_RUN | 正本・手順 |
| 更新 | Phase1§7・D_ENTRY・PHASE／HANDOVER／LEDGER／ROADMAP | 誘導 |

**やらない**: Z二列、系統別価格、Yahoo／楽天／B統合、ALLOW_* 自動ON。

---

## 4. リスクと緩和

| リスク | 緩和 |
|--------|------|
| API件数増 | 確認に系統数×SKUを明示 |
| 部分成功の誤読 | 系統別 runId／ok を完了ダイアログに出す |
| 列破壊 | Phase1どおり系統別のみ書込 |

---

## 5. 検収

- [x] 方針承認… **2026-08-01**  
- [x] コード実装… **2026-08-01**（チェックUI・順次PUT・部分成功）  
- [ ] 片方のみ＝Phase1相当（任意・未再実施。Phase1単系統検収済）  
- [x] 両方 dry_run → 両列… **2026-08-01**（`…f372b8`／`…40d85e`・VALID）  
- [x] 両方 prod 1SKU… **2026-08-01**（`…8fcbc2`／`…cc7c72`・ACCEPTED）  
- [ ] キャンセルで PUT なし（任意・未再実施。A2でキャンセル経路検収済）  
- [x] docs（承認／HUMAN_RUN／PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY）  

詳細: [D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md](D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md) §3

---

## 6. 社長確認

- [x] §2… **2026-08-01**（三点スキップ・ファサード2回・部分成功可）  
- [x] 実機… **2026-08-01**（両系統 dry＋prod）  

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草・承認反映。実装着手。 |
| 2026-08-01 | コード実装完了。実機（両系統 dry→prod）待ち。 |
| 2026-08-01 | **検収OK**: dry `…f372b8`／`…40d85e`・prod `…8fcbc2`／`…cc7c72`。 |
