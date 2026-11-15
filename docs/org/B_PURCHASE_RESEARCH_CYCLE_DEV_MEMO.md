# 仕入れ検討① CYCLE — 開発・仮想テスト方針

**日付**: 2026-08-22  
**正本**: [B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md](B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)  
**仮想**: `tools/competitor_store/dry_cycle_alight_hit.py`

---

## 1. 基本方針

**C0–C5 は原則「1フェーズずつ：仮想 PASS → GAS 実装 → clasp → 実機」**（FULL P1–P3 と同型）。

| 段 | 仮想 | 実装 | 実機 |
|----|------|------|------|
| 前提 | 列定義・COMPETITOR_FIELDS CSV | — | — |
| C0 | 単独 | §5.4 | ①ログ | **fixture 7/7 PASS**（GAS **`セラー調査済みSP-API`** 実装済・要 clasp） |
| **C1+C2** | **まとめて** | A-light + ヒット再計算 | common_ss_view |
| C3 | 単独 | 自動選定 | サマリ行 | **fixture 7/7 PASS**（`dry_cycle_c3_select.py`）・**実機 PASS** |
| **C3+T** | — | C3+完走トリガー | ①ログ | **実機 PASS** アサヒ `pr_20260824_013324` |
| C4 | 単独（C3 fixture 渡し）+ 統合1本 | 完走連鎖 | P3 ログ | **fixture 8/8 PASS**・**実機 PASS** アサヒ→永谷園 arm |
| C5 | 不要 | docs | HUMAN_RUN §10・FLOW §3 | **済** 2026-08-24 |

---

## 2. まとめて仮想する組

### C1 + C2（店レーン）

- C1 が `セラー.抽出メーカー` を書き、C2 が **全店再走査** で `セラーヒット` を出す **1本のデータフロー**
- C2 単独は **バックフィル** 検証のみ。A-light 後の再計算は **C1+C2 セット** で見る
- ツール: **`dry_cycle_alight_hit.py`**（fixture 駆動・API なし）

### C3 → C4（メーカーレーン）

| 順 | 内容 |
|----|------|
| 1 | **C3 単独** … メーカーサマリ fixture の選定順 assert |
| 2 | **C4 単独** … 完走→D成功/失敗→arm |
| 3 | **統合1本** … C3 出力 → C4 状態機械（GAS 前の最終確認） |

### まとめない

- **C0 + C1**（SP-API/BB vs Keepa query/product）
- **C0–C5 一括**
- **C4 + 新規トリガー大量**（P3 拡張は別承認）

---

## 3. テスト期間 vs 完成形

| 期間 | 実装順の目安 |
|------|----------------|
| **テスト期間**（C0 未実装） | **C1→C2→C3→C4→C0** も可（手動セラーで回す） |
| **完成形** | doc 順 **C0→C1→…**（全自動） |

仮想の切り方は同じ（**C1+C2 まとめ、C0 単独**）。

---

## 4. 各フェーズの仮想 PASS 定義

| ID | PASS の意味 |
|----|-------------|
| **C1+C2** | 門・①候補・Keepaフル行数 **不変**。ヒット **再計算** が fixture 期待と一致。再巡回 **水増しなし** |
| C0 | fixture ASIN → 新規 sellerId。429 で停止。http は飛ばして済印。**`dry_cycle_seller_accum.py` 7件**（`セラー調査済みSP-API` タグ） |
| C3 | ティアA→B、要確認除外、洗い途中は続きのみ |
| C4 | D 失敗で arm しない。D2 leftover=0 必須 |
| 実機 | clasp push → ①ログ + `common_ss_view` |

---

## 5. C0 仮想 — fixture 設計

**置き場**: `tools/competitor_store/fixtures/cycle_seller_accum/`

**実行**:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/competitor_store/dry_cycle_seller_accum.py
python tools/competitor_store/dry_cycle_seller_accum.py --case T-CSA-03
python tools/competitor_store/dry_cycle_seller_accum.py --list
```

### 5.1 fixture JSON スキーマ

| キー | 意味 |
|------|------|
| `sellers_before` | バッチ前の `セラー` 行 |
| `cand` | ①候補（`門結果`・`ASIN`） |
| `keepa_full` | BB 列 JOIN 用（**GET なし**） |
| `side_effect` | ①候補・Keepaフル行数（**不変**） |
| `run.offers_by_asin` | SP-API `getItemOffers` **モック** |
| `run.stop_at_asin` | 429 停止 ASIN |
| `expect.new_seller_ids` | 追加 sellerId 順序 |
| `expect.after_asin` | 429 時の再開カーソル |

### 5.2 ケース一覧（必須）

| ID | 検証目的 |
|----|----------|
| **T-CSA-01** | **基本** … BB+オファー和集合・調査CK OFF |
| **T-CSA-02** | **cap=2** … 新規100上限の打ち切り |
| **T-CSA-03** | **429** … 2件目で停止・after_asin・BB も未追加 |
| **T-CSA-04** | **除外** … Amazon直販・既存・短ID |
| **T-CSA-05** | **`セラー調査済みSP-API`** … 済み ASIN は SP-API スキップ |
| **T-CSA-06** | **429 再実行** … 済みタグで Morita 塊を飛ばす |
| **T-CSA-07** | **http 飛ばし** … 2件目200以外でも済印。BBは残す。3件目まで進む |

### 5.3 ロジック正本（dry = GAS）

| 関数（Python） | GAS 相当 |
|----------------|----------|
| `run_seller_accum` | `prRunSellerAccum_` |
| `ok_sid` / `pass_asin_list` | `prSellerAccumOkSid_` / `prCandPassAsinListForSellerAccum_` |
| `mark_cand_spapi_done` | `prMarkCandSellerSpapiDone_` |
| `bb_map_for_pass` | `prKeepaFullBbMapForAsins_` |

**①候補書込**: **`セラー調査済みSP-API`／`セラー調査日時` のみ**（門・Keepa GET なし）。

**実スプシ smoke**: `dry_a1_seller_accum.py`（`--smoke` は API 任意）。

---

## 6. C1+C2 仮想 — fixture 設計

**置き場**: `tools/competitor_store/fixtures/cycle_alight_hit/`

**実行**:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/competitor_store/dry_cycle_alight_hit.py
python tools/competitor_store/dry_cycle_alight_hit.py --case T-CAH-03
python tools/competitor_store/dry_cycle_alight_hit.py --list
```

### 6.1 fixture JSON スキーマ

```json
{
  "id": "T-CAH-01",
  "desc": "説明（日本語）",
  "sellers_before": [
    {
      "sellerId": "AYC4Z8PML8T30",
      "店名": "モリタストア",
      "調査CK": true,
      "抽出メーカー": "永谷園|石原水産",
      "食品比率": "46.9"
    }
  ],
  "maker_sum_before": [
    { "メーカー": "永谷園", "実態": "製造メーカー", "セラーヒット": 0 }
  ],
  "side_effect": { "cand_rows": 181, "keepa_full_rows": 220 },
  "batch": {
    "shop_seller_ids": ["AYC4Z8PML8T30"],
    "mocks": {
      "AYC4Z8PML8T30": {
        "total": 30,
        "products": [
          { "asin": "B0XXX", "manufacturer": "永谷園", "brand": "永谷園" }
        ]
      }
    },
    "product_cap": 20,
    "shop_cap": 3
  },
  "expect": {
    "seller_hits": { "永谷園": 2, "石原水産": 1 },
    "sellers_after": {
      "AYC4Z8PML8T30": {
        "調査CK": false,
        "status": "done",
        "抽出メーカー_contains": ["永谷園"]
      }
    },
    "maker_sum_jitta": { "永谷園": "製造メーカー", "未知メーカー": "要確認" },
    "side_effect_unchanged": true
  }
}
```

| キー | 意味 |
|------|------|
| `sellers_before` | バッチ前の `セラー` 行（最小列） |
| `maker_sum_before` | 既存サマリ（`セラーヒット` バックフィル前後） |
| `side_effect` | ①候補行数・Keepaフル行数（**変化してはいけない**カウンタ） |
| `batch.mocks` | Keepa `/query`+`/product` の **モック**（API 呼ばない） |
| `expect.seller_hits` | 全 `セラー` 再走査後の **セラーヒット** 正解 |
| `expect.side_effect_unchanged` | true なら cand/keepa 行数不変 |

### 6.2 ケース一覧（必須）

| ID | 検証目的 |
|----|----------|
| **T-CAH-01** | **バックフィル** … batch なし。3店 history からヒット再計算 |
| **T-CAH-02** | **単店 A-light** … 1店 mock → 抽出メーカー更新・CK OFF |
| **T-CAH-03** | **再巡回水増しなし** … 同 sellerId を2回 batch してもヒットは sellerId 単位1 |
| **T-CAH-04** | **累計ヒット** … 2店が同メーカー → hit=2 |
| **T-CAH-05** | **Generic 除外** … 抽出に Generic あってもヒット・サマリ選定対象外 |
| **T-CAH-06** | **hold 巨大店** … total>1000 でも cap 分 product からメーカー抽出 |
| **T-CAH-07** | **副作用なし** … ①候補・Keepaフル カウンタ不変（A-light 契約） |

### 6.3 ロジック正本（dry = GAS の単体テスト）

| 関数（Python） | GAS 相当（実装時） |
|----------------|-------------------|
| `extract_makers_from_products` | A-light product ループ |
| `merge_extract_cell` | `抽出メーカー` `\|` 連結 |
| `recompute_seller_hits` | C2 全店再走査 |
| `apply_alight_batch` | `prRunAlightHit_` / `prMenuAlightHit` |
| `sync_maker_sum_from_hits` | サマリ insert + ヒット列 |

**除外**: `Generic`・空・非仕入先ラベルはヒットに数えない（`maker_alias.JITTAI_EXCLUDE` + 固定集合）。

**禁止 assert**: `prUpsertCandidate_` 相当呼び出し、`Keepaフル` append。

### 6.4 PASS / FAIL 出力

- 全ケース PASS → exit 0、`PASS n/n`
- 1件でも不一致 → exit 1、差分を stderr
- GAS 実装後: 同 fixture 期待値を **実機 smoke** の目安に流用（値は runId 付きで NOW に1行）

---

## 7. 実装後の実機チェック

### 7.1 C0

1. `cd purchase_research` → `clasp push`
2. メニュー「C0. セラー貯め」
3. `python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent judge`
4. 確認: 競合DB `セラー` 新規行（調査CK OFF）・①候補・Keepaフル行数不変

### 7.2 C1+C2

1. `cd purchase_research` → `clasp push`
2. A-light メニュー（実装名は C1）
3. `python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent maker`
4. 確認: `セラー` 更新・`メーカーサマリ.セラーヒット`・**①候補行数が増えていない**

### 7.3 C4

1. `cd purchase_research` → `clasp push`
2. 洗い完走（C3 または 99 洗い完走ループ）で `洗い状態=完了` → 自動 `C4 完走連鎖` ログ
3. 手動: **C4. 完走連鎖**（直近完走 or 完了行1件）
4. `python tools/common_ss_view/dump.py dump --menu 仕入れ検討① --intent judge`
5. 確認: `C4 完走連鎖` に `d_ok arm=` または `stopped d2_leftover` / `pick=0`

**仮想**:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/competitor_store/dry_cycle_c4_chain.py
```

| ID | 検証目的 |
|----|----------|
| **T-CCC-01** | D成功→次社 arm |
| **T-CCC-02** | D1 FAILED→停止 |
| **T-CCC-03** | D2 FAILED→停止 |
| **T-CCC-04** | D2 leftover→停止 |
| **T-CCC-05** | 洗い未完→スキップ |
| **T-CCC-06** | D成功・次社なし |
| **T-CCC-07** | C4 OFF |
| **T-CCC-08** | 完走サマリ更新 |

### 7.4 C3+T（本線）

1. **C3+T. 自動選定→洗い→D1→完走トリガー**（1回）
2. ①ログ: `C3+T 自動選定→洗い→D1→完走トリガー` + `trigger_armed`
3. 以降 `full_trigger 洗い完走` が `LOOP COMPLETE` まで自動
4. 完走後: §7.3 の C4 確認

**実機合格（アサヒ）**: `pr_20260824_013324`（2026-08-24 1:36–1:43）

---

## 8. C5 — docs 整備（済）

| ファイル | 更新内容 |
|----------|----------|
| [B_PURCHASE_RESEARCH_FLOW.md](B_PURCHASE_RESEARCH_FLOW.md) §3 | C0–C5 済・C3+T/C4 実機 |
| [B_PURCHASE_RESEARCH_HUMAN_RUN.md](B_PURCHASE_RESEARCH_HUMAN_RUN.md) §10 | 日常本線・合格目安・Property |
| 本ファイル §1・§7.4・§8 | 実機 PASS 行・C3+T 手順 |
| [CHANGE_LEDGER.md](../CHANGE_LEDGER.md) | C5 1行 |
| [AGENT_HANDOVER.md](../AGENT_HANDOVER.md) §6 | C5 1行 |

---

## 9. 改訂

| 日付 | 内容 |
|------|------|
| 2026-08-24 | **C5** 済。C3+T 本線・アサヒ実機合格・HUMAN_RUN §10 |
| 2026-08-23 | C4 fixture 8件 + `dry_cycle_c4_chain.py` + GAS 完走連鎖 |
| 2026-08-22 | C0 fixture 4件 + `dry_cycle_seller_accum.py` |
| 2026-08-22 | 初版。C1+C2 fixture 設計 + dry_cycle_alight_hit.py |
