# B / Z — ④⑤⑥ SP-APIリサーチ 人間手順

**親**: [B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md](B_SPAPI_RESEARCH_P0_P2_REQUIREMENTS.md)  
**日付**: 2026-08-13

---

## 0. 設定マスタ内寸（済）

1. スプシ `00_設定マスタ` → ダンボールサイズ表（58行付近）
2. E=`内寸A_mm`／F=`内寸B_mm`／G=`内寸C_mm`／H=`梱包種別` を確認
3. **`exclude`**: `商品入荷箱`・`Nekopos封筒（他）`（非定型・選定対象外）

---

## U1＝Keepa取得_ログ拡張（⑤・実装済）

1. `clasp push` → Keepa取得（APIが走ると新列が埋まりやすい）
2. `Keepa取得_ログ` に partNumber／ブランド／製造者／梱包_*  
3. OFF: `B_KEEPA_LOG_EXT_ENABLED=false`

---

## U2＝Keepa取得ログ整理（⑤・実装済・実機済）

1. `clasp push` → シート再読込
2. Z→12「**12-⑧ Keepa取得ログを整理（90日超or JAN最新3実行以外→archive）**」
3. 確認ダイアログの保持／移動件数を確認して実行
4. 本体 `Keepa取得_ログ` が減り、`Keepa取得_ログ_archive` に追記されること
5. マスタは変わらないこと
6. OFF: `B_KEEPA_LOG_ARCHIVE_ENABLED=false`

**保持ルール**: 直近90日 **かつ** 同一JANの最新3実行（日時ユニーク）のみ本体に残す。

**実機（2026-08-13）**: `archived=3262 kept=249 days=90 keepRuns=3`

---

## U3＝競合ASIN投票診断（⑤・実装済／U3.4a）

**前提**: 対象親の **`A.セット商品数`**（空なら**同一親の子行から継承**）を解決してから投票。

1. `clasp push` → シート再読込
2. 親にレ点 → Z→15「**15-⑯**」
3. メモの `titleScore=…/th=…/lcs=…/ok|ng` を確認
4. マスタ非書込

**Catalogタイトル（U3.4a）**:
- 正規化: NFKC・小文字・空白/記号除去
- 短い方の bi-gram カバー率 ≥ 閾値（既定 **0.80**）
- かつ最長共通部分 ≥ min(8, floor(|短|*0.5))
- どちらか空 or 長さ&lt;6 → **不一致**（Catalog票・ヒント抑制）
- 閾値: Script Property `B_KEEPA_ASIN_VOTE_U3_TITLE_SCORE`（例 0.75）

**U3.4c maker一般照合**:
- Catalog `brand` またはタイトルに、メーカーの正規化表記／カナ→ローマ字が含まれるか
- 個別ブランド表は使わない（メモ: `/makerVia:brand|title` or `/makerOut`）

**U3.4d**: メーカーの正は **`メーカー名ベース` のみ**（`メーカー名` 等の商品名用列は見ない）

**U3.4e**:
- 照合は **brand／製造者のみ**（タイトル除外）。OR 一致
- 両方空＝unknown → **推奨しない**
- brand yes でも setGuess≠親 → `set_mismatch`

**U3.4f**:
- 照合は **brand のみ**（製造者除外）
- セット親（set≠1）× Catalog単品 + brand yes → `brand_catalog_hint`（…0s1）

**次**: U3.4f 実機済（`U3_20260813_124927_d939a7`）→ **公式単品ASINをマスタへ手入れ** → FBA P1a（15-⑰）

---

## U4＝FBA P1a（Catalog寸法診断）

**入口**: Z→15「**15-⑰ FBAティア診断P1a**」  
**出力**: シート `FBAティア診断_P1a` のみ（マスタ `サイズ＆自己発/FBA` は書かない）  
**ASIN優先**: `ASINコード` → なければ `競合店ASINコード`

### 手順
1. U3推奨の**単品ASIN**をマスタへ反映（下表）。競合が違う行は **上書き必須**
2. P1a対象親だけレ点（スキップは外す）
3. 15-⑰ 実行 → `FBAティア診断_P1a` で HTTP／梱包寸法／仮ティアを確認
4. OFF: `B_FBA_P1A_DIAG_ENABLED=false`

### 採用（run U3_20260813_124927）

| 親 | ASIN | 手入れ |
|----|------|--------|
| …0460 | B0040Q3ZU2 | 必須（競合≠推奨） |
| …0019 | B00HRS69XA | 必須 |
| …8175 | B01M7YSWAF | 競合一致ならそのまま |
| …5300 | B0FJFK5NG9 | そのまま |
| …20149 | B01N5A6ESU | そのまま |
| …1013-oya | B084RJSH7W | そのまま |
| …0s1 | B084RJSH7W | 必須（競合404） |
| …66119 | B0D9VLMRGD | 任意（単品ヒント） |

### スキップ（レ点外す）
…0446（brand_fail）／…81514（死ASIN疑い）／…5127・6018・92019（set_mismatch）

缶飯で標準2b／415 確認済。母数≥20合格は公式単品整備の積み上げ後。

### 実機（runId `595a71f2`・2026-08-13）

| 指標 | 結果 | ライン |
|------|------|--------|
| 親 | 20 | ≥20 |
| Catalog HTTP200 | **20/20 (100%)** | ≥90% |
| 梱包寸法3辺 | **11/20 (55%)** | ≥50% |
| 要確認_寸法乖離 | **0** | — |
| 型番非空／JAN一致 | P1aシートに列なし→**未計測** | ≥50%／≥90% |

**梱包なし9行**はすべて次の2 ASINの重複（Catalog HTTP200だが attributes に梱包寸法なし）:

- `B0FL2CH7NL`（蒸し生姜湯）
- `B0FL2FX1WZ`（六漢生姜湯）

**梱包あり（P1b候補・ユニークASIN 6）**: `B0CJ8MT2CT`／`B0FQV5847X`／`B0FQV1NR77`／`B0DQ7S9TZN`／`B0CVN4XP1D`／`B0DG8JXJH2`  
仮ティア例: 小型／標準1／標準2e。マスタサイズ列は未変更。

**ゲート解釈**: HTTP＋梱包%クリア＋P1b完了。**2026-08-13 暫定合格（社長OK）** — 型番/JANは FBA必須ゲートから外し U7 側へ。クリティカル無し。

---

## U4b＝FBA P1b（本線書込）

**入口**: Z→15「**15-⑱ FBAティア本線書込P1b**」  
**書込列**: `サイズ＆自己発/FBA` の右隣 **`FBAティア`**／**`FBA手数料_円`**（スプシへ列新設済）  
**触らない**: `サイズ＆自己発/FBA`（自己発専用）

### 手順
1. 公式単品ASINを N列等へ入れた親にレ点（**梱包寸法ありのみ**。`595a71f2`後は6ASIN・親11件に絞り済）
2. （任意）先に 15-⑰ で診断確認
3. **15-⑱** → 確認ダイアログ YES
4. マスタの `FBAティア`／`FBA手数料_円` が空だったセルだけ埋まる
5. OFF: `B_FBA_P1B_WRITE_ENABLED=false`

### 今回のP1b対象（`595a71f2`梱包あり）
`B0CJ8MT2CT`／`B0FQV5847X`／`B0FQV1NR77`／`B0DQ7S9TZN`／`B0CVN4XP1D`／`B0DG8JXJH2`  
（生姜湯 `B0FL2CH7NL`／`B0FL2FX1WZ` はスキップ＝レ点OFF済）

### 実機（runId `P1b_20260813_162652_6111aa`・2026-08-13 16:26〜）

| 項目 | 結果 |
|------|------|
| 対象親 | **11** |
| 書込 `FBAティア` | **11** |
| 書込 `FBA手数料_円` | **11** |
| スキップ（既存／ASINなし／梱包なし／乖離／不適合） | **すべて 0** |
| `サイズ＆自己発/FBA` | **未変更** |

| ASIN | ティア | 手数料 |
|------|--------|--------|
| B0CJ8MT2CT | 標準2e | 430 |
| B0FQV5847X | 標準1 | 318 |
| B0FQV1NR77 | 標準1 | 318 |
| B0DQ7S9TZN | 標準2e | 430 |
| B0CVN4XP1D | 小型 | 288 |
| B0DG8JXJH2 | 標準1 | 318 |

**合格**: この母数の P1b は完了。**FBA本線いったんOK**（型番/JANはU7へ移管）。

### スキップ条件
Catalog梱包寸法なし／3辺和差>15%／ティアが settings 不適合／セル既存値あり／ASINなし

---

## U1〜U7

| ID | 手順 | 合格 |
|----|------|------|
| U1 | Keepaログ拡張列 | **実装済** |
| U2 | Z 12-⑧ ログ整理 | **実装済・実機済** |
| U3 | Z 15-⑯ 投票診断 | **U3.4f実機済** |
| U3b | Z 15-⑳ 競合空のみ自動 | **実装済**（要 clasp push） |
| U4 | Z 15-⑰ FBA P1a | **実機済・ゲート暫定合格**（`595a71f2`） |
| U4b | Z 15-⑱ FBA P1b | **実機済**（`P1b_20260813_162652_6111aa`: 11/11書込） |
| U5 | Z 15-⑤／15-⑥ 品番＋自社INT- | **実装済** |
| U6 | B Step6.6 | **実装済**（要 clasp push） |
| U7 | Catalog品番 | **診断済・本線当面不要** |

---

## U3b＝競合店ASIN 空のみ自動（15-⑳）

**入口**: Z→15「**15-⑳ 競合店ASIN空のみ自動**」  
**書込**: `競合店ASINコード` **空のみ**＋黄セル  
**採用**: `brand_yes`／`brand_set_match`／`unanimous`／`majority`／`brand_prefer`、および同一親の子セット一致競合（`own_child_set_match`）  
**◎矛盾**: Keepaログの◎があり推奨と異なる → 書かない  
**OFF**: `B_COMP_ASIN_AUTOFILL_ENABLED=false`  
**承認**: [LV4_B_COMPETITOR_ASIN_AUTOFILL_APPROVAL.md](LV4_B_COMPETITOR_ASIN_AUTOFILL_APPROVAL.md)

### 手順
1. セット数確定後、競合が空の親にレ点（**既存値がある親は対象外＝黄にならない**）
2. **15-⑳** 実行 → 完了ダイアログの書込数／スキップ内訳を確認
3. 黄セル＋セルメモ（runId）を目視
4. 問題あれば該当セルを消し Property OFF

### 2026-08-13 牛丼511 調査
- clasp ログ（2回）: `COMPASIN_20260813_174159_3d8234`／`COMPASIN_20260813_174300_ff2b63`  
  → いずれも **`targets=1 filled=0 skipped=1`**、スキップは **`row=90 skip=no_rec result=brand_unknown`**（**511は一度も対象に入っていない**）
- スプシ511（親・レ点）: **O空**だが **P=`…/dp/B084RJSH7W`**。子512は O=`B084RJSH7W`・セット1
- **主因**: 空判定が `pickCompetitorAsinCandidateFromParentRow_`（**O＋P列URL**）→ Pだけで「既存あり」扱いされ **511は対象外**。黄セルは「今回書込分」なので当然付かない
- 副因（旧高信頼）: `brand_unknown`／`unanimous` 等は書かない／同一親の子競合は兄弟票に使わない → 空親だと書込0になりやすい
- 対策: **O列のみで空判定**／高信頼拡大／`own_child_set_match` フォールバック／ダイアログにスキップ内訳。O511は再試験用に空へ戻した（P URLは残置可）

---

## U7＝メーカー品番 Catalog 診断（マスタ非書込）

**入口**: Z→15「**15-⑲ メーカー品番U7診断**」  
**出力**: シート **`メーカー品番診断_U7`**（毎回クリア再書）  
**比較**: Catalog型番（attributes）／Catalog JAN（identifiers）／Keepa `partNumber`  
**ASIN優先**: `ASINコード` → なければ `競合店ASINコード`  
**OFF**: `B_U7_PART_DIAG_ENABLED=false`  
**件数**: `B_U7_PART_DIAG_MAX_PARENTS`（未設定時は P1a と同じ既定）  
**本線**: **当面不要**（食品寄りで JAN 汚染。本番は U5/U6）

### 手順
1. 公式単品寄りの親にレ点（P1b対象など）
2. **15-⑲** 実行
3. シートで Cat型番非空率・JAN一致・Keepa一致・`推奨ソース` を確認
4. 本線（Step6.7）は EANガード付き再診断後の別承認

### 見る列
`Cat型番`／`Cat型番_U5ガード`／`JAN一致`／`Keepa_partNumber`／`一致_CatvsKeepa`／`推奨ソース`

---

## U6＝B Step6.6

**入口**: B.統合実行の **6.6 メーカー品番（下書き・空のみ）**（手動は従来どおり 15-⑥）  
**関数**: `menuFetchMakerModelForBIntegratedStep_` → U5 と同じ core（quiet）  
**OFF**: `B_MAKER_MODEL_FETCH_ENABLED=false`（Step 自体はスキップして続行）

---

## U5＝メーカー品番（空のみ）

**入口**: Z→15「**15-⑥ メーカー品番取得（レ点・親SKU行）**」（試験は 15-⑤ 選択行でも可）  
**メーカー**: マスタ **`メーカー名ベース`のみ**（`メーカー名`は見ない）  
**書込**: **`メーカー品番下書き` のみ**（空のときだけ）。`メーカー型番`／`型番`／旧`メーカー品番` へは書かない。  
出典: `メーカー品番出典元`（あれば）  
**優先**: Keepa `partNumber` → SerpAPI抽出 → 自社 `INT-`＋8桁hex（最大12・黄セル・出典 `INT-self`）  
**結果シート**: `メーカー品番取得結果`（sourceKind 列あり）

### 手順
1. 品番空の親にレ点（少数推奨）
2. **15-⑥** 実行
3. マスタ品番・黄セル・結果シートの sourceKind を確認
4. OFF: `B_MAKER_MODEL_FETCH_ENABLED=false`

---

## 緊急停止

| Key | 値 |
|-----|-----|
| `B_KEEPA_LOG_EXT_ENABLED` | false |
| `B_KEEPA_LOG_ARCHIVE_ENABLED` | false |
| `B_KEEPA_ASIN_VOTE_U3_ENABLED` | false |
| `B_KEEPA_ASIN_VOTE_U3_REQUIRE_SET` | false |
| `B_FBA_P1A_DIAG_ENABLED` | false |
| `B_FBA_P1B_WRITE_ENABLED` | false |
| `B_MAKER_MODEL_FETCH_ENABLED` | false |
| `B_LOGISTICS_USE_3D_FIT` | false |

---

## 検収メモ

| 日付 | 内容 | 結果 |
|------|------|------|
| 2026-08-13 | U1〜U3.3 | 調査単品3件改善。0446はCatalog誤推奨残 |
| 2026-08-13 | U3.4a | タイトル score ゲート。要 clasp push＋15-⑯ |
| 2026-08-13 | U3.4b | 0s1: brand yes Catalog優先。0446: makerInTitle。実測0446=0.923/ok |
| 2026-08-13 | U3.4c | maker一般照合（英字↔カナ＋Catalog brand）。別名表なし。要 clasp push |
| 2026-08-13 | U3.4d | メーカー正＝メーカー名ベースのみ。実測サバトンはメーカー名空 |
| 2026-08-13 | U3.4e | brand/製造者のみ・タイトル除外・unknown非推奨。要 clasp push |
| 2026-08-13 | U3.4f | brandのみ＋セット親×Catalog単品ヒント。要 clasp push |
