# 競合専用ストア — 要件（段階導入・マスタ温存）

**文書種別**: requirements  
**最終更新**: 2026-08-15  
**状態**: 段階3までコード済。12-⑬ **人間検収済**（2026-08-15、Listings HTTP200・qty=0・マスタ非書）。全JAN自動・トリガーなし。  
**フィールド辞書**: [COMPETITOR_FIELDS.md](COMPETITOR_FIELDS.md)  
**仮想テスト**: `tools/competitor_store/`

---

## 0. 確定事項

| 項目 | 内容 |
|------|------|
| 置き場 | **専用スプシ1本が正**。出品 GAS も領域1 GAS も同じ ID（`COMPETITOR_SS_ID`） |
| マスタ | **12-⑭** JAN＋`A.セット商品数`。**最安**（送料無料優先）。**1件目禁止**。計N袋優先。単価2倍外れは捨て。精度課題 [COMPETITOR_SETCOUNT_ACCURACY.md](COMPETITOR_SETCOUNT_ACCURACY.md) |
| ロールバック | 運用マスタの `Keepa取得_キャッシュ`・`▼ログ` は **削除しない・書込も止めない** |
| ①リサーチ／出品 | 基本1回。`purpose=research` |
| ②定時 | 12-⑫ JAN1件。在庫の正は **Amazon**。12-⑬は sellerSku **1件 GET**。全JAN自動・トリガーなし。マスタ非書 |
| 寿命 | 90日超は手動削除。退避はローカル最新1本＋Driveは12-⑪（最新1本・確認あり）。自動月次は未実装 |
| マップ | 日本語論理名×Amazon／楽天／Yahoo。左端＝マスタ候補列名。1シート＋`valid_from`。過去行は触らない |
| Amazon② | SP-API 今値。Keepa `csv` 非保存。**後回し**（④は Keepa BuyBox／Amazon現在） |
| Amazon貼付 | [B_AMAZON_COMPETITOR_PASTE](B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md)。**Z併存**: スナップショットは薄い運用のまま。倉庫は新規 `Keepaフル`（◎＋K＋生JSON） |
| ログ | `▼ログ` は専用に統合しない |
| 変化なし非書込 | **実装済（要 clasp push）。** キー＝目的＋モール＋検索JAN＋店商品コード。指紋＝表示価格・送料フラグ・楽天％／Yahoo点数・商品名。生JSON・順位・レビューは見ない |

---

## 1. 移行フェーズ

| フェーズ | 内容 | 今 |
|----------|------|----|
| 1 | 専用作成。マスタ Keepa は従来どおり。ENABLED 時のみ専用へ **追加書き**（dual-write）。読取は **マスタ優先** | 済 |
| 2 | B Step2 の楽天・Yahoo ヒットを `モールヒット` へ追記 | 済（ENABLED 時に動く） |
| 3 | ②隔日・在庫API・90日削除・退避 | **済**（Amazon在庫1SKU検収）。全JAN自動はしない |
| 最終 | マスタ Keepa キャッシュ書込廃止。読取を専用へ | やらない（戻せるように残す） |

**Property**

| Key | 未設定 | 意味 |
|-----|--------|------|
| `COMPETITOR_STORE_ENABLED` | **false**（現行のみ） | `true` のとき Keepa 書込後に専用へコピー |
| `KEEPA_PASTE_NEW_ASIN_CAP` | **20／ブロック** | 貼付Aの新規 Keepa。キャッシュヒット除外。token実測後に見直し。トリガーは未設置 |

切戻し: `COMPETITOR_STORE_ENABLED` を false または削除。マスタキャッシュはそのまま。

---

## 2. 専用スプシのシート

| シート | 粒度 | 備考 |
|--------|------|------|
| `項目マップ` | 論理名1行 | 左端 `マスタ候補列名`、論理名、Amazon、楽天、Yahoo、`適用開始日` |
| モールヒット | 検索ヒット1行 | ◎＋○を列化（中古・キャッチは生JSON）。ポイントは％と数を分離。△IDはマージしない |
| `Keepaスナップショット` | ASIN1行 | **運用ビュー**。マスタキャッシュと同系統。フル化しない |
| `Keepaフル` | ASIN1行＋生JSON | 倉庫。クロスウォーク ◎＋K。`csv[]` 非保存。**目的=リサーチ／定時**。P0スキーマ済 |
| `運用メタ` | 1行運用 | マップ版、保存先ID |
| `メーカーマスタ` | メーカー1行 | 第2クエリ語の正。語があれば公式／LLMを再走しない。カテゴリIDは載せない。90日削除の対象外 |

---

## 3. 連携契約（領域1）

- 調査複製スプシ `1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE` に **競合シートを作らない**。メーカー第2クエリ語の正は `メーカーマスタ`。
- **Keepa 属性の正は `Keepaフル`。** 領域1①が product したらここに書く（`目的=リサーチ`。90日。`csv[]` 落とす）。門の作業面だけ調査複製 `①候補` へ転記（当面 dual-write 可）。
- **Catalog 生は `モールヒット` に入れない。**
- 出品 clasp の scriptId を領域1に差し替えない。
- 領域1のモールヒット追記は `目的=リサーチ` のみ。`定時` は出品側。
- 同一 JAN のモールヒット行を上書きしない（追記）。Keepaフルは ASIN 最新1行更新、指紋変化時のみ追記。
- **2026-08-15 再ロック**: リサーチ Keepa と出品 Keepa は同じ倉庫。①候補は倉庫ではない。セラー台帳は後。進捗は [NOW](B_PURCHASE_RESEARCH_NOW.md)。範囲・開発順は [TASK_STRUCTURE](B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)（ストアが受けるのは W1 追記。Catalog 生はヒット禁止）。

---

## 4. 関連メニュー影響（段階1）

| メニュー | 影響 | 段階1の期待 |
|----------|------|----------------|
| A Keepa 20/50/全件 | `writeKeepaCache` | **必ずマスタへ書く**。ENABLED 時のみ専用スナップショット追加。**`Keepaフル` は①リサーチと A が同じ倉庫。** 90日内は再GETしない |
| Amazon貼付準備 | Catalog→貼付空欄 | 12-⑮ dry_run／12-⑯書込。`AmazonCompetitorPaste.js` |
| 12-⑦ Keepaキャッシュ古い行削除 | `menuPurgeKeepaCacheOlderThanDays` | **マスタの `Keepa取得_キャッシュ` のみ**。専用は触らない |
| 12-⑩ 定時対象JANログ | `menuLogCompetitorScheduledCandidates` | 読取のみ。再検索しない。トリガー非設置 |
| 12-⑬ Amazon在庫1SKU | `menuReadAmazonInventoryOneSku` | Listings優先、なければFBA。マスタ非書。1 SKUのみ |
| B Step2 モール横断 | 楽天Yahoo API | ENABLED 時、生ヒットを `モールヒット` 追記。**B統合**かつストアONならマスタ直書きスキップ→ **2.1＝12-⑭** |
| 12-⑭ セット紐付け | 専用ヒット読→マスタ | レ点／当B行。Z単独は従来どおり手動。`COMPETITOR_MASTER_APPLY_ENABLED` 未設定=ON。B切戻し `B_COMPETITOR_STORE_APPLY_ENABLED=false` |
| B 6.55 FBA | 寸法 | マスタキャッシュ経由のまま |
| 99-⑪ ベイクオフ | 診断シート | 専用へ自動連携しない |

聖域: `generateRakutenCSV`／`Yahoo.js`。`B_INTEGRATED_STEP_FUNCTIONS` は 2.1（12-⑭）を直後追加。15〜23はBに入れない。

---

## 5. 仮想テスト（GASメニューの代わり）

`python tools/competitor_store/run_tests.py`

| ID | 代替 | 合格 |
|----|------|------|
| T0 | スキーマ | 6シート（`Keepaフル`含む）とヘッダーが要件どおり |
| T1 | マップ＋hits | ベイクオフ1 JAN を `目的=リサーチ` で追記。競合確定列に1件目を書かない |
| T2 | Aキャッシュ読 | マスタ Keepa をコピー。マスタ行数不変 |
| T3 | キャッシュヒット | 専用 miss → マスタでヒット |
| T4 | 12-⑦ | パージ対象はマスタ日付のみ。専用行数不変 |
| T5 | 領域1 | 同じ店に `リサーチ` のみ。`定時` を作らない |
| T6 | チャンク2マップ | ベイクオフ行を論理列へ。競合確定は空。目的はリサーチ |
| T7 | 90日削除 | 古いヒットだけ消える。新しい行は残る。出品マスタIDには書かない |
| T8 | 定時候補 | 子在庫>0のみ。2日以内の定時は除外。マスタ非書 |
| T9 | ローカル退避 | 最新フォルダ1本 |
| T10 | Drive退避計画 | 正本・出品マスタIDはゴミ箱にしない |
| T11 | 定時ヒット | 目的=定時。競合確定は空 |
| T15 | 変化なし指紋 | 同指紋スキップ |
| T16 | セット紐付け | 1件目・ふるさと・4Pを落とす。12袋の最安 |

---

## 6. やらない

定時の全JAN自動、自動トリガー、Keepa 全件再取得、マスタ Keepa キャッシュ廃止。12-⑭は全件ループではなく**出品CKレ点行のみ**。

---

## 7. コード

- [AmazonSpapiPut.js](../../AmazonSpapiPut.js): `menuReadAmazonInventoryOneSku`（1 SKU GET）
- [コード.js](../../コード.js): 12-⑨〜⑭
- `tools/competitor_store/`（T0–T16）

---

## 8. 専用スプシ ID

人間作成（2026-08-15）: `1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs`  
https://docs.google.com/spreadsheets/d/1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs/edit  

Python 設定: `tools/purchase_research_path3/config.local.json` の `COMPETITOR_SS_ID`（gitignore。他ツールの config.local.json には書かない）。  
GAS: Script Property `COMPETITOR_SS_ID` に同じ ID（dual-write 用。`COMPETITOR_STORE_ENABLED` は未設定のまま＝OFF）。

シート初期化: `python tools/competitor_store/init_store.py`。書込は `tools/c1_hpc_packaged/secrets/token_sheets_rw.json`（C1 読取 token とは別）。  
仮想検証は従来どおり `local_store/`。

---

## 9. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-08-16 | B 2.1＝12-⑭。B統合かつストアONで横断シート直書きスキップ。15〜23はB外 |
| 2026-08-15 | Keepaの正はKeepaフル。①が書く。①候補は転記。90日再利用 |
| 2026-08-15 | タブ「セラー」。第1版モリタ1行。Keepaフルからセラーは貯めない |
| 2026-08-15 | P0: `Keepaフル` スキーマ＋`keepa_official.csv`。A書込は未配線 |
| 2026-08-15 | `メーカーマスタ` 追加。領域1の巨大メーカー第2クエリ語。調査複製には置かない |
| 2026-08-15 | A+C: 計N袋優先・各N袋除外。12-⑭単価2倍外れ除外。課題 ACCURACY.md |
| 2026-08-15 | 12-⑭ 専用ヒット→マスタをセット数紐付け（1件目禁止・fromP除外） |
| 2026-08-15 | 変化なし非書込を実装（指紋5項目）。領域1細分化は DOMAIN1 §1 |
| 2026-08-15 | シート名・ヘッダーを日本語化。`init_store.py` で専用ブックへタブ＋ヘッダー投入。 |
