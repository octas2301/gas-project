# レーンA2 — 運用固め（確認キャンセル・トグル・逃げ道）承認パッケージ

**日付**: 2026-08-01  
**状態**: **検収OK**（2026-08-01）。三点スキップ。原則コードなし  
**レーン方針**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)／[LV4_LANE_A1_FBA_OFFER_APPROVAL.md](LV4_LANE_A1_FBA_OFFER_APPROVAL.md) §8  
**前提**: A1・デュアル Phase1 **検収OK**  
**手順**: [D_MENU_LANE_A2_HUMAN_RUN.md](D_MENU_LANE_A2_HUMAN_RUN.md)  
**三者レビュー**: **スキップ**（運用検収＋docs。契約・列・EC一括変更なし）

---

## 1. 目的

レーンAの安全運用を実機で閉じる（機能追加ではない）。A1 任意の A1-c を独立チケット化。

| 段階 | 内容 | 成功定義 |
|------|------|----------|
| **A2-a** | D 相乗りの開始前確認で **キャンセル** | PUT なし。ログ `cancelled_by_user`／理由 `ユーザー取消` |
| **A2-b** | 作業後 **Property トグル戻し** | `ENABLED`／`ALLOW_PROD`／`ALLOW_MASTER_QTY` を false にできる手順が明記され実施可能 |
| **A2-c** | **逃げ道疎通**（スモーク） | `D補助. Amazon段階実行（旧E）` または `Z → 21-⑩` 等にメニュー到達できる（本線 D を壊さない） |

**含まない**: A3（マスタ在庫>0）、Phase2 同時2PUT、楽天／Yahoo／B統合、新規機能本実装。

---

## 2. 社長確定方針（2026-08-01）

| # | 論点 | 決定 |
|---|------|------|
| 1 | コード | **原則なし**。実機で穴が出たときだけ最小修正 |
| 2 | 成果物 | 本承認包＋HUMAN_RUN＋PHASE／HANDOVER／LEDGER／ROADMAP／D_ENTRY 参照更新 |
| 3 | 実機 | キャンセル1回＋トグル戻し＋逃げ道メニュー到達。prod PUT は A2 必須にしない |
| 4 | 三点 | **スキップ** |

---

## 3. 変更ファイル

| 種別 | パス | 内容 |
|------|------|------|
| 新規 | 本ファイル | 正本 |
| 新規 | `D_MENU_LANE_A2_HUMAN_RUN.md` | A2-a/b/c 手順 |
| 更新 | `D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md` | §2 を A2 へ誘導 |
| 更新 | PHASE／HANDOVER／LEDGER／ROADMAP | 到達点 |

**やらない**: `Yahoo.js`、楽天聖域、B統合、MASTER 解禁ロジック。

既存コード: 確認キャンセルは `runBatchExportAmazonFacade` で `cancelled_by_user` を返す実装済。

---

## 4. 想定リスクと緩和

| リスク | 緩和 |
|--------|------|
| キャンセル試験で誤 OK→prod | `ALLOW_PROD=false` のまま、または dry_run 選択でキャンセル |
| 逃げ道で本番書込 | A2-c はメニュー到達のみ。21-⑩は dry。prod は必須にしない |
| トグル戻し忘れ | HUMAN_RUN＋CURRENT_PHASE Property 表 |
| A3 へ膨張 | MASTER／在庫>0 は除外 |

---

## 5. 検収

- [x] 方針・docs起草… **2026-08-01**  
- [x] A2-a… **2026-08-01** `cancelled_by_user`（12:07:21。PUTなし）  
- [x] A2-b… **2026-08-01**（社長「A2実機OK」・作業後トグル手順どおり）  
- [x] A2-c… **2026-08-01** `menuAmazonCourseE0Precheck` DONE checked=1 warn=0（D補助／旧E）  
- [x] docs 記録  

### 5.1 実機ログ要約

| 段階 | 記録 |
|------|------|
| A2-a | `[runBatchExportAmazonFacade] state=FAILED cancelled_by_user course=amazon`（12:07:21）。dry_run／自己発／七味 `…19as13`。`SPAPI_PUT_OFFER_CK_*` なし |
| A2-c | `menuAmazonCourseE0Precheck` RUNNING→DONE checked=1 warn=0（12:08:42）。UI「実行をキャンセルしました」＝書込なしで可 |

---

## 6. 社長確認

- [x] §2… **2026-08-01**（三点スキップ・原則コードなし）  
- [x] 実機検収（§5）… **2026-08-01**  

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 起草・承認反映（実装＝docs。実機待ち）。 |
| 2026-08-01 | **検収OK**（A2-a/b/c）。次＝A3。 |
