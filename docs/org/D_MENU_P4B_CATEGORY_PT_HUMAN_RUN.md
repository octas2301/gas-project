# P4b — カテゴリ／PT 本線接続（人間手順）

**状態**: **P4b-a／P4b-b 合格**（2026-08-01）  
**承認**: [LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md) §2  
**C1手順**: [D_MENU_C1_HUMAN_RUN.md](D_MENU_C1_HUMAN_RUN.md) §3／§8  
**前提**: [LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md](LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md)（P4a 読取合格）

---

## 0. 実行分担（重要）

| 作業 | 誰が実行 |
|------|----------|
| GASメニュー（21-⑱・D一括出品等） | **人間**（シート上） |
| **ローカル Python**（`c1_packaged.py`／`c1_fetch_inputs.py` 等） | **Agent モード**（ターミナル代行）。人間にメニュー案内だけで終わらせない |
| Script Properties トグル | 人間（または Agent が手順を明示して確認） |
| SC アップロード | 人間 |

---

## 1. 方針（ロック済・要約）

| 項目 | 内容 |
|------|------|
| メニュー | **Z → 18〜21 → 21 → 21 SC・P4b → 21-⑱** |
| 決め方 | **P4b-d**: 複数競合ASIN（最大5）→一致チェック→browse多数決。全滅／無しは JAN→SHELF等の重み付き投票。Browse無しPT禁止 |
| 多数決 | 作業台貼付◎は廃止。**マスタ耐久ASINの Node 投票は可** |
| 正しさチェック | 不合格ASINは不採用（404・タイトル・JAN・肉↔魚） |
| C1本線 | **FOOD／SEASONING／HPC**（FOODマップ）。**PT+browse 両方必須**・既定埋込禁止（2026-08-02）。`HERB` は提案のみ・C1除外 |
| 件数 | 親最大3・ASIN試行最大5・トグル既定 false |

> **注意（2026-08-02）**: §3 の「HERB→defaults SEASONING」挙動は **旧仕様の合格記録**。現行 FOODマップではマスタ空／非許可PTは **親除外**（唐辛子既定で埋めない）。

---

## 2. P4b-a — **合格**

| 項目 | 値 |
|------|-----|
| runId | `P4B_PT_20260801_161210_546de6` |
| parent | `sanky-4538872180149-oya` |
| refAsin | `B01N5A6ESU` |
| 書込 | PT=`HERB`／browse=`唐辛子 (2430212051)` |
| sources | `rival_asin+catalog+catalog_browse` |

---

## 3. P4b-b — **合格（HERBのまま・フォールバック）**

| 項目 | 値 |
|------|-----|
| 実行 | Agent: `c1_fetch_inputs` → `c1_packaged.py --mode dry_run --sub-batch CK_daba393f8055_B2` |
| マスタ | 親 PT=**HERB**（fetch直後確認済） |
| C1 runId | `C1_CK_daba393f8055_B2_20260801_073624` |
| 出力 | `…_20260801_073624_DRYRUN.xlsm` |
| xlsm PT（行8親・行9子） | **`SEASONING`**（HERB なし＝不採用OK） |
| xlsm browse | defaults 全文 `食品・飲料・お酒 > … > 唐辛子 (2430212051)` |
| fingerprint | match／acceptedParents=1 |

---

## 4. 合格記録

| 段階 | 結果 | メモ |
|------|------|------|
| 方針〜マスタ競合実装 | **OK** | 2026-08-01 |
| P4b-a | **OK** | `…161210_546de6` |
| P4b-b（HERBフォールバック） | **OK** | C1 `…073624`／PT=SEASONING |
| トグル戻し | □ | `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED=false` |

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-08 | **P4b-d** 手順反映（複数ASIN投票・JAN／SHELF重み・Browse必須）。 |
| 2026-08-01 | 起草〜マスタ競合実装。 |
| 2026-08-01 | P4b-a合格。P4b-b手順（HERB）。 |
| 2026-08-01 | **P4b-b合格**。§0に「ローカルPythonはAgentモードで実行」を明記。 |
| 2026-08-02 | C1 FOOD: 許可PTにFOOD・PT/browse必須・既定禁止。§1注意追記。 |
