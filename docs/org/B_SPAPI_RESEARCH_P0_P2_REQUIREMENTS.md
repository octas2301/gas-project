# B / Z — ④サイズ・FBA／⑤競合ASIN／⑥メーカー品番（SP-API最小追加）

**文書種別**: 要件定義（実装前）  
**日付**: 2026-08-13  
**状態**: **④自己発3D／FBA P1a・P1b 実装済・実機OK**／**⑤ U1〜U3 実装済・競合列本線は承認待ち**／**⑥ U5・U6 実装済／U7診断実装済**（U7本線は後続）
**親**: [CURRENT_PHASE.md](../CURRENT_PHASE.md)／[AGENT_HANDOVER.md](../AGENT_HANDOVER.md) §8  
**手順**: [B_SPAPI_RESEARCH_P0_P2_HUMAN_RUN.md](B_SPAPI_RESEARCH_P0_P2_HUMAN_RUN.md)  
**関連**: [B_ASIN_N_AUTO_FILL_REQUIREMENTS.md](B_ASIN_N_AUTO_FILL_REQUIREMENTS.md)（公式カタログASIN）／[LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)（Catalog JAN）／[PRICING_V1_REQUIREMENTS.md](../PRICING_V1_REQUIREMENTS.md)（Step3.1）

**実装優先（2026-08-13 確認）**: **④自己発3D →（後）FBA／⑤／⑥**

**ゴール一文**: 今ある道具で資産化（P0）→ SP-API Catalog 読取で精度実験（P1a）→ 合格後に④⑥本線と3D箱フィット（P1b）。⑤は人間◎を上限とする。

---

## 0. スコープ

| 番号 | 内容 | 本包 |
|------|------|------|
| **④** | `サイズ＆自己発/FBA`（自己発箱＋FBAティア）＋**3D箱フィット** | P0マスタ整備／P1a寸法実験／P1b本線 |
| **⑤** | 競合ASIN・URL精度（Amazon中心） | P0ログ・投票／P1 Catalog票 |
| **⑥** | メーカー品番（B載せ＋SP-API＋自社生成） | P0 B載せ／P1b SP-API |

**含めない**

- Open Food Facts（Amazon寸法不足時の将来）
- 問屋PDF／パッケージOCR（見積・発注プロジェクト）
- GS1／CLIP／Rainforest 等の新ベンダー
- 楽天CSV聖域・Yahoo.js
- 人間◎の自動上書き
- 複数箱分割（**1出荷＝1箱**）

**番号の混同防止**: 既存「⑥ N列ASIN自動」（Step6.5）とは別。本書の⑥＝メーカー品番。

**3者レビュー**: **不要**（方針合意済。LV4承認包＋P1a実測で足りる）。

---

## 1. 共通定義

### 1.1 公式カタログASIN（SP-APIを信じる対象）

[N列ASIN](B_ASIN_N_AUTO_FILL_REQUIREMENTS.md) と同一条件 **かつ単品相当**:

| 条件 | 内容 |
|------|------|
| ◎ | Keepa取得_ログ最新回 **または** ASIN貼付の評価◎ |
| ブランド＝メーカー | 正規化一致 |
| ASIN | 親 `競合店ASINコード` 先頭／URL抽出 |
| **単品** | セット出品ASINは参考のみ（本線の寸法・型番に使わない） |
| JAN | 8桁以上 |

### 1.2 Catalog寸法の注意

| 種類 | 扱い |
|------|------|
| **商品サイズ** | 自己発3Dフィットの商品寸法として優先候補 |
| **梱包サイズ** | Amazon外箱の可能性 → 自己発箱選定では**要確認**/過大になりやすい |
| セット品ASIN寸法 | 使わない |

### 1.3 人間◎（⑤）

Amazon競合ASINの正解上限は**人間◎継続**（名前＋画像を見た人の判断。機械は◎を付けない）。自動は **候補列**・ログ・要確認のみ。評価列の上書き禁止。貼付自動化の要件メモ（実装しない）: [B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md](B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md)。

### 1.4 フェーズ

| フェーズ | 内容 | マスタ本線書込 |
|----------|------|----------------|
| **P0** | 今ある道具＋設定マスタ内寸＋Keepaログ＋⑥B載せ＋Z診断 | ⑥空のみ／ログ／設定マスタ |
| **P1a** | SP-API Catalog `attributes` **診断のみ** | 診断シートのみ |
| **P1b** | 合格後：④3D＋FBA／⑥SP-API本線 | 空のみ＋要確認 |
| **P2** | Open Food Facts 等 | 本包外 |

---

## 2. ④ サイズ＆自己発/FBA ＋ 3D箱フィット

### 2.1 現状（ベース）

- Step3.1: 利益ベースで送料サイズ選定。`3辺和=幅+奥行+高さ×セット数` は `B_LOGISTICS_USE_PHYS_RANK` **既定OFF**
- FBAティア自動: ほぼ未実装
- ダンボール表: 金額はB/D、サイズはC列テキストのみだった

### 2.2 設定マスタ（済・2026-08-13）

**表**: `00_設定マスタ` ダンボールサイズ（A58付近）

| 列 | 内容 |
|----|------|
| A〜D | 既存（表示サイズ文言・金額）。**維持** |
| **E** | `内寸A_mm`（数値・向きなし） |
| **F** | `内寸B_mm` |
| **G** | `内寸C_mm` |
| **H** | `梱包種別`（`rigid`／`soft`／**`exclude`**） |

- 単位 **mm 統一**
- C列テキストは人間用。計算は E/F/G
- スクリプト: `write_box_inner_dims.py`／`mark_box_exclude.py`
- **`exclude`**: 非定型で内寸を入れられない行。**3D・資材自動選定の対象外**（2026-08-13: `商品入荷箱`・`Nekopos封筒（他）`）
- 封筒2辺のみは第3辺既定 **10mm**（手修正可）

### 2.3 3Dフィット（④自己発・実装）

| 項目 | 内容 |
|------|------|
| Property | `B_LOGISTICS_USE_3D_FIT` **未設定=ON** |
| 商品寸法 | `AI情報取得data`。照合優先: **商品名正規化＋卸値(税込)** → 商品名 → **JAN＋卸値** → JAN先頭（警告） |
| 箱 | E/F/G。`exclude`・内寸空はスキップ |
| 向き | 商品・箱とも自由（6向き） |
| N | 1〜40・**1箱**。格子＋ずらし1段（**各軸に1個以上入るときだけ**ずらし可。半辺すり抜け禁止） |
| soft | 箱を 8% 大きく見なして判定 |
| 選定 | **サイズ昇順 first-fit**。**rigid を先に全サイズ試行**し、入らなければ soft。Compact の stagger は当面維持 |
| 送料 | `00_設定マスタ` **自己発送表** |
| 箱代 | ダンボール資材表 **D列** |
| 合計（送料+箱代） | **利益判定・警告のみ**（箱選定には使わない） |
| FBA | **P1a診断**＋**P1b本線**（`FBAティア`／`FBA手数料_円`。自己発列は別） |
| 入口 | 既存 Step3.1／Z「3.1 想定物流費AI試算（単体）」 |
| **利益との関係** | 箱・サイズはサイズ選定のみ。利益不足でも箱を変えない。`送料利益確保不可` は警告のみ |
| 寸法あり・不適合 | 最安ネコポス埋め**しない**（空＋理由 `サイズ選定不可…`） |
| 寸法なし | 従来フォールバック容認 |

**廃止（2026-08-13）**: 利益確保不可時の「最小セットのみネコポス強制」。箱代最安フィット。

### 2.4 FBAティア

#### P1a（診断・実装済）

| 項目 | 内容 |
|------|------|
| 入口 | Z→15-⑰ `menuFbaTierCatalogDiagnoseP1aForCheckedParents` |
| 対象 | レ点親（子SKU空）。最大 `B_FBA_P1A_MAX_PARENTS`（既定20） |
| ASIN | `ASINコード` → なければ `競合店ASINコード` |
| Catalog | GET `/catalog/2022-04-01/items/{asin}` `includedData=summaries,attributes,identifiers` |
| 抜出 | `item_package_dimensions`／`item_dimensions`／`item_package_weight` 等 |
| 出力先 | シート **`FBAティア診断_P1a`** のみ（毎回クリア再書） |
| 比較 | AI梱包（`getAiPackDimsForLogistics_`）と3辺和差。**>15%** → `要確認_寸法乖離=TRUE` |
| 仮ティア | **`00_設定マスタ` A=`FBA手数料`** を first-fit（B=区分名・D=手数料・F=条件）。空／不適合時のみ概算フォールバック |
| F条件 | `25x18x2.0cm/250g`＝3辺box／`50cm/2kg`＝**3辺和**＋重量。標準外枠(45×35×20・9kg)外は大型系のみ |
| 出力列 | `仮FBAティア`／`FBA手数料_円`／`ティア判定ソース` |
| **禁止** | マスタ列 `サイズ＆自己発/FBA` への書込 |
| Property | `B_FBA_P1A_DIAG_ENABLED` 未設定=ON／`false`でOFF |

#### P1b（本線書込・実装済）

| 項目 | 内容 |
|------|------|
| 入口 | Z→15-⑱ `menuFbaTierCatalogWriteP1bForCheckedParents` |
| 対象 | レ点親。最大 `B_FBA_P1B_MAX_PARENTS`（未設定時は P1a と同じ既定） |
| ASIN | N列`ASINコード`（＝競合ASINの相乗り先）→ なければ `競合店ASIN` |
| 入力 | Catalog **梱包**寸法＋重量（AI寸法のみでは書かない） |
| 判定 | 設定マスタ `FBA手数料` first-fit（`source=settings` のみ書込） |
| 乖離 | Catalog↔AI 3辺和差 **>15%** → 書込スキップ |
| 書込列 | **`FBAティア`**／**`FBA手数料_円`**（`サイズ＆自己発/FBA` の右隣。自己発列は触らない） |
| 空のみ | 既存値があるセルは上書きしない |
| Property | `B_FBA_P1B_WRITE_ENABLED` 未設定=ON／`false`でOFF |

#### P1b（旧メモ・上書き）

- 料金シミュレーター近似: Catalog入力 ↔（任意）Product Fees APIは乖離検証のみ
- 自己発サイズとFBAは別系統（不一致＝即エラーにしない）

### 2.5 P0/P1a での④

- P0: 内寸列のみ（済）。本線ロジック変更なし
- P1a: 診断シートに Catalog寸法・AI梱包寸法・仮FBAティアを並べる（Keepa列は任意・後続）

---

## 3. ⑤ 競合ASIN・URL精度

### 3.1 方針

| 指標 | 扱い |
|------|------|
| 人間精査後◎ | **上限（約92%）**。置き換えない |
| 自動◎ | P0ログ改善＋投票で底上げ |
| 楽天/Yahoo URL | 既存モールAPI継続 |

### 3.2 P0

1. **Keepa取得_ログ拡張（U1・実装済）**（同一API応答・追加呼出なし）  
   列: `partNumber`／`ブランド`／`製造者`／`梱包_L_cm`／`梱包_W_cm`／`梱包_H_cm`／`梱包_重量_g`  
   - 入口: 既存 Keepa取得（ASIN貼付）。`ensureKeepaFetchLogSheet` が欠列を末尾追加  
   - 寸法: Keepa `packageLength` 等（mm）→ cm  
   - キャッシュヒット時はブランド／製造者／梱包_* をログへコピー追記（既存ログ非改変）。**読取はヘッダー名**（setCount空でも旧形式と誤判定しない）。梱包が空かつ `梱包_checked` 未記入なら API再取得。API更新時は既存ブランド／製造者／梱包を空で上書きしない。梱包が空なら itemLength 等で補完。  
   - OFF: `B_KEEPA_LOG_EXT_ENABLED=false`  
2. **行数対策（U2・実装済）**: 直近90日 **かつ** JANあたり最新3実行のみ本体に残し、他は `Keepa取得_ログ_archive` へ移動。Z→**12-⑧**  
   - OFF: `B_KEEPA_LOG_ARCHIVE_ENABLED=false`  
   - 日数: `B_KEEPA_LOG_ARCHIVE_DAYS`（既定90）／実行数: `B_KEEPA_LOG_KEEP_RUNS_PER_JAN`（既定3）  
3. **多源投票診断（U3・実装済／U3.1）**（マスタ自動上書きなし）: 人間◎／マスタN列・競合店／Catalog JAN  
   → シート `競合ASIN投票診断_U3`（`要確認`列）。Z→**15-⑯**（**セット数確定後**）  
   - OFF: `B_KEEPA_ASIN_VOTE_U3_ENABLED=false`  
   - 件数: `B_KEEPA_ASIN_VOTE_U3_MAX_PARENTS`（既定30）  
   - Catalog JAN票: `B_KEEPA_ASIN_VOTE_U3_CATALOG_JAN`（未設定=ON）※**親 `A.セット商品数`=1 のときのみ票に入れる**  
   - セット必須: `B_KEEPA_ASIN_VOTE_U3_REQUIRE_SET`（未設定=ON。欠ける行は `set_qty_missing`）  
   - **U3.1b**: 親のセット数が空なら同一親SKUの**子行から継承**（子に1があれば1、なければ子の最小。メモに `セット継承:`）  
   - U3.1診断列: `親セット数`／`setGuess_*`／`ブランド一致_*`  
   - **U3.2**: ブランド一致を推奨に反映（`brand_yes`／`brand_prefer`／`brand_fail`）。単品親の conflict は Catalog を `conflict_catalog_hint`。set≠1でも Catalog を `CatalogHint`。競合 Catalog 404 は dead。setGuess≠親セットは `set_mismatch`  
   - **U3.3**: 同一JAN・同セットの他行競合を `兄弟セット` 票。`sibling_set_match`／`sibling_alt_set`。Catalog と商品名のタイトル不一致時は Catalog 票・ヒントを抑制  
   - **U3.4a**: Catalogタイトル判定を U3専用に厳格化（bi-gramカバー率≥`B_KEEPA_ASIN_VOTE_U3_TITLE_SCORE` 既定0.80＋LCS下限。短文/空は不一致）。P4b `amazonP4bTitleLooksRelated_` は変更しない  
   - **U3.4b**: `brandCat=yes` の Catalog はタイトルngでもヒント維持（…0s1）。conflict で brand unknown ならメーカー照合必須（…0446）。兄弟は**他親SKUのみ**
   - **U3.4c**: maker照合を一般化（個別ブランド表なし）。カタカナ↔ローマ字（ヘボン簡易）＋ Catalog `brand`／タイトル。メモに `/makerVia:brand|title` または `/makerOut`。`サバトン`↔`Sabaton` 等は規則で一致
   - **U3.4d**: メーカー照合の正は **`メーカー名ベース` のみ**（`メーカー名` や商品名用の他メーカー列は見ない）
   - **U3.4e**: メーカー照合は **brand／製造者のみ**（タイトル除外）。どちらか一致で yes。両方空＝unknown。**unknown は推奨しない**（あやふや禁止）。brand yes でも setGuess≠親なら `set_mismatch`
   - **U3.4f**: 照合は **brand のみ**（製造者除外）。セット親（set≠1）で Catalog が単品＋brand yes のときだけ `brand_catalog_hint`（…0s1・サイズ用）

### 3.3 P1

- Catalog JAN検索は **U3 で既に1票**（`B_KEEPA_ASIN_VOTE_U3_CATALOG_JAN`）
- 公式単品ASINは重み大。人間◎と矛盾 → 書込停止＋要確認（本線書込は別承認）
- FBA P1a 再実行は U3 で推奨された単品ASIN整備後

### 3.4 競合列・空のみ本線（実装済 2026-08-13）

- 承認包: [LV4_B_COMPETITOR_ASIN_AUTOFILL_APPROVAL.md](LV4_B_COMPETITOR_ASIN_AUTOFILL_APPROVAL.md)（A〜G 承認済）
- 入口: Z **15-⑳** `menuFillCompetitorAsinFromU3ForCheckedParents`
- 書込: `競合店ASINコード` **空のみ**／黄セル／`brand_yes`・`brand_set_match` のみ
- 人間◎があり推奨と異なる → 書かない
- OFF: `B_COMP_ASIN_AUTOFILL_ENABLED=false`

---

## 4. ⑥ メーカー品番

### 4.1 優先順

| 順 | ソース |
|----|--------|
| 1 | SP-API `manufacturerPartNumber` 等（公式単品ASIN・JAN一致）※P1b |
| 2 | Keepa `partNumber` |
| 3 | SerpAPI（現行。期待低・維持） |
| 4 | **自社品番生成**（商品名ベース等 → `INT-`＋短いハッシュ、最大20、要確認黄セル） |
| 5 | 子SKU（**export時**フォールバック。Q10b） |

### 4.2 P0 / U5・U6（実装済 2026-08-13）

- 既存 `menuFetchMakerModelFromApisForCheckedParentRows`（15-⑥）＋選択行（15-⑤）を拡張
- 優先: **Keepa partNumber → SerpAPI → 自社 `INT-`＋8桁hex（最大12・黄セル）**（型番不明は書かない）
- 書込先は **`メーカー品番下書き` のみ**（`メーカー型番` 等へは書かない）。空のみ。出典列 `メーカー品番出典元`
- メーカー参照は **`メーカー名ベース` のみ**（`メーカー名` は見ない。U3と同方針）
- 候補ガード: **JAN同一（数字正規化）は却下**／純数字は **4桁未満のみ却下**（5桁 `16100` 等は可）→ Keepa/Serp とも次候補 or INT-
- Property `B_MAKER_MODEL_FETCH_ENABLED` 未設定=ON
- SerpAPI精度は上げない（実行率のみ。Keepaヒット時は Serp スキップ）
- **U6**: B統合 **Step6.6**（`menuFetchMakerModelForBIntegratedStep_`）。6.5の直後。quiet（アラートなし）

### 4.3 P1b / U7

| 状態 | 内容 |
|------|------|
| **U7診断（実装済）** | Z **15-⑲** → `メーカー品番診断_U7`。マスタ非書込。OFF: `B_U7_PART_DIAG_ENABLED=false` |
| **U7本線** | **当面不要**（2026-08-13 実機: 食品寄りで `part_number`＝JAN/EAN汚染・Keepaと同汚染。Catalog最優先本線化しない）。本番品番は **U5/U6（Keepa→Serp→INT-）** |
| **再検討条件** | EAN疑いガード（8〜14桁純数字却下等）付きで再診断し、真の型番率が十分なカテゴリが出たとき |

見つからなければ自社生成（④⑤と独立して先に15-⑥単体で試験可）。

---

## 5. 実装ユニット（Z起点・最小単位）

| 順 | ID | 内容 | 入口 | 本線書込 |
|----|-----|------|------|----------|
| 1 | U0 | Keepa取得（既存） | A | 既存 |
| 2 | U1 | Keepaログ拡張列 | A内包 | ログのみ |
| 3 | U2 | ログ整理アーカイブ | Z 12-⑧ | archive |
| 4 | U3 | 競合ASIN投票診断 | Z 15-⑯ | 診断 |
| 5 | U4 | FBAティア診断P1a（Catalog寸法） | Z 15-⑰ | 診断のみ |
| 6 | U5 | 品番＋自社生成 | Z 15-⑤⑥ | **実装済**（空のみ） |
| 7 | U6 | 品番B Step6.6 | B | **実装済** |
| 8 | U7 | Catalog品番診断→（後）SP-API品番本線 | Z 15-⑲／B 6.7 | **診断実装済**／本線は別承認 |

**最終載せ先**: A=Keepa／B=Step／C・D=マスタ経由間接／Z=診断正。

### P1a / FBA本線 合格基準（公式単品ASIN ≥20件）

| 指標 | ライン | FBA本線（U4/U4b） | U7（SP-API品番） |
|------|--------|------------------|------------------|
| Catalog HTTP 200 | ≥90% | **必須** | 参考 |
| 寸法3辺取得 | ≥50% | **必須** | 参考 |
| 型番非空 | ≥50% | **必須から外す**（2026-08-13 暫定） | 計測時に再評価 |
| JAN一致（JANあり時） | ≥90% | **必須から外す**（同上） | 計測時に再評価 |

**暫定合格の理由（2026-08-13・社長OK）**  
- 実機 `595a71f2` / `P1b_…6111aa` で HTTP・梱包・本線書込にクリティカル無し  
- 食品系 Catalog の `attrKeys` に品番キーがほぼ無く、型番ゲートは FBA 精度と非連動  
- 品番は U5/U6（Keepa→Serp→INT-）で別途カバー  

**FBAトラック**: HTTP＋梱包クリア＋P1b実機済 → **完了（いったんOK）**。  
**U7**: コード着手は別承認。型番/JANは U7 切る前に測るか、要件で再定義する。

未達（HTTP/梱包） → P1bしない（P0の⑥B載せのみ継続可）。

---

## 6. B統合 Step 案（実装承認時に確定）

```
P0:  … 6.5 N列 → 6.6 メーカー品番 → 7 → 7.6
P1b: … 3.05 SP-API/3D寸法（3.1の前）… 6.7 SP-API品番（6.6の後）
```

`B_INTEGRATED_STEP_FUNCTIONS` 変更は **EC重要変更・実装前承認必須**。

---

## 7. 変更予定ファイル（実装時）

| フェーズ | ファイル | 概要 |
|----------|----------|------|
| P0済 | `00_設定マスタ` E〜H | 内寸・種別 |
| P0済 | `tools/c1_hpc_packaged/write_box_inner_dims.py` | 内寸書込 |
| P0 | `コード.js` | Keepaログ・12-⑧・Step6.6・投票診断 |
| P1a済 | `コード.js` | 15-⑰ Catalog寸法→`FBAティア診断_P1a`・仮ティア |
| P1b | `コード.js` | FBA表本線・6.7 |
| docs | 本ファイル／HUMAN_RUN／承認包 | 正本 |

**触らない**: Yahoo.js、楽天CSV、P4b本線の勝手な改変。

---

## 8. リスクと復元

| リスク | 深刻度 | 対策 |
|--------|--------|------|
| 誤FBA/誤箱で価格判断 | 中 | P1aまで本線④書込禁止 |
| 誤自社品番で出品 | 中 | 黄セル要確認 |
| B Step挿入ずれ | 低〜中 | 実行中にpushしない |
| 内寸列 | なし | 現行読取はA〜Dのみ。E〜H追加済でもStep3.1は不変 |

復元: 各 Property OFF／git revert／内寸列は無視で従来運用可。

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-13 | **⑤競合空のみ本線**: 15-⑳ 実装。U7本線は当面不要（JAN汚染実機） |
| 2026-08-13 | **U7診断**: Z 15-⑲ Catalog型番 vs Keepa →`メーカー品番診断_U7`。マスタ非書込。`B_U7_PART_DIAG_ENABLED` |
| 2026-08-13 | **⑤競合列空のみ本線**: 承認包起草（実装前）。[LV4_B_COMPETITOR_ASIN_AUTOFILL_APPROVAL.md](LV4_B_COMPETITOR_ASIN_AUTOFILL_APPROVAL.md) |
| 2026-08-13 | **ゲート暫定**: FBA本線は HTTP+梱包のみ必須。型番/JANは U7 側へ移す（社長OK・クリティカル無し） |
| 2026-08-13 | **P1b実機 `P1b_20260813_162652_6111aa`**: FBAティア/手数料 11/11書込・スキップ0。U7は型番/JAN方針待ち |
| 2026-08-13 | **P1a実機 `595a71f2`**: HTTP20/20・梱包11/20。生姜湯2ASINは梱包属性なし。型番/JANは診断シート未計測。P1bは梱包ありに限定 |
| 2026-08-13 | **⑤ U2**: Keepa取得_ログ整理。90日超 or JAN最新3実行以外→`Keepa取得_ログ_archive`。Z 12-⑧ |
| 2026-08-13 | **⑤ U1**: Keepa取得_ログに partNumber／ブランド／製造者／梱包寸法・重量。追加APIなし。`B_KEEPA_LOG_EXT_ENABLED` |
| 2026-08-15 | §1.3: ◎は人間のみ・候補列。貼付自動化メモ（未実装）[B_AMAZON_COMPETITOR_PASTE](B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md) |
| 2026-08-13 | **FBA P1a×設定マスタ**: 仮ティア＝`FBA手数料` first-fit＋手数料円。備考Fをbox/三辺和パース |
| 2026-08-13 | **FBA P1a**: Z 15-⑰ Catalog寸法診断シート・仮ティア。マスタサイズ列は書かない。`B_FBA_P1A_DIAG_ENABLED` |
| 2026-08-13 | **rigid優先**: 段ボ等を先にサイズ昇順。soft（紙袋等）は rigid 全滅後。Compact stagger 当面維持 |
| 2026-08-13 | **寸法突合**: 商品名＋卸値(税込)優先（同JAN複数行対策）。JAN先頭一致はフォールバック |
| 2026-08-13 | **サイズ昇順 first-fit**: 箱代最安廃止。送料=自己発送／箱代=資材D／合計=利益判定のみ。寸法あり不適合は空欄 |
| 2026-08-13 | **サイズのみ**: 利益不可でも箱をネコポスへ上書きしない。`size_only` ログ |
| 2026-08-13 | v1.1 ④自己発3D実装・exclude・優先④→⑤⑥。FBAは後続 |
| 2026-08-13 | v1 要件ロック。④3D・向き自由・N≤40・1箱。⑤人間◎。⑥自社生成。設定マスタE〜H書込。P0→P1a→P1b |
