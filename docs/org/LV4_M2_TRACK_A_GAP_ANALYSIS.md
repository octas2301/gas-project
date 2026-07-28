# M2（TRACK=A・既存ASINオファー）— 要件ギャップ洗い

**日付**: 2026-07-27（更新 2026-07-28）  
**状態**: **v1実装済＋発汗実機合格**。SCは公式 ListingLoader xlsm 埋めが正（簡易CSV単独は不可の場合あり）。残GAP＝Loader自動埋め。  
**正本**: [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §0・§3.2・§8・§12.4  
**Facade**: [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) **U6**  
**HUMAN_RUN**: [D_MENU_M2_HUMAN_RUN.md](D_MENU_M2_HUMAN_RUN.md)  
**承認パッケージ**: [LV4_M2_IMPLEMENTATION_APPROVAL.md](LV4_M2_IMPLEMENTATION_APPROVAL.md)  
**依存**: M1検収済。C1 は **B用** — **M2 では流用しない**。  

### 決定ロック（2026-07-27）
- PACKAGED=**案L**（`tools/m2_offer_packaged`）  
- ゲート=**manual_ok**（人間目視）  
- 試験=**発汗** `lifec-4560151300832-48s11`／ASIN `B07YND44VN`（競合店ASIN）

---

## 1. ゴール（再掲）

既存カタログ ASIN に、自社 SKU・価格・在庫 0/1・配送テンプレで **オファーを載せる**。  
新規バリエーション作成（M1／C1）とは別パイプライン。

---

## 2. 層ごとの現状

| 層 | 要件（§8／§12.4） | 現状 | 判定 |
|----|-------------------|------|------|
| ルート判定 TRACK=A | Aのみ。B条件でも新規にしない | `AmazonApprovalExport.js` 実装済 | **OK** |
| TRACK 未設定 | 実行しない | 実装済 | **OK** |
| ASIN必須 | 欠けたら生成停止／スキップ | Resolve: 親 `ASINコード` 空→`SKIPPED_NEED_HUMAN`。Build: 行に asin | **OK（親列前提）** |
| GENERATED 行 | SKU＋ASIN＋価格＋在庫＋配送 | `variationRole=offer`・track=`A` で出力 | **OK（骨格）** |
| 在庫 | バルク内 0/1（マスタ非書込） | `inventoryMode`→stockOut | **OK** |
| 販売中スキップ | 在庫>0 は出さない | `SKIPPED_IN_STOCK` | **OK** |
| 配送テンプレ | 既定 `送料無料パターン` | Property 可 | **OK** |
| ブランドゲート | `SKIPPED_BRAND_GATE` | **実装済**（`manual_ok` 必須） | **OK(v1)** |
| 出品制限 | 同上 | 人間目視（v1） | **運用** |
| 子ASIN優先 | （要件はマスタASIN） | 子→親。`ASINコード`／**競合店ASIN**／URL | **OK** |
| 単品（子なし） | Aは単品多い | TRACK=A は親子完全性チェックを bypass | **OK想定** |
| PACKAGED | Offer／Inventory Loader 系 | **案L CSV＋`m2_listing_loader_fill.py`** | **OK**（自動DLは対象外） |
| SC→UPLOADED_OK | §12.4 | 発汗1SKU成功→E-5 | **OK（実機）** |
| HUMAN_RUN／承認書 | U6 | **更新済** | **OK** |
| メニューE／D | M1コース想定 | M2は Property=`A`＋手順差。E文言は後で薄い案内可 | **後続** |

---

## 3. GENERATED（TRACK=A）列契約（現行コード）

ヘッダー（Bと共通）:

`track,parentSku,childSku,sellerSku,manufacturerPart,productName,brand,priceAmazon,inventory,gtin,asin,mainImageUrl,subImageUrls,amazonCategory,setCount,shippingTemplate,variationRole`

| 列 | A行の典型 | メモ |
|----|-----------|------|
| track | `A` | |
| variationRole | `offer` | Bの parent/child ではない |
| asin | 親または子の `ASINコード` | 空なら Resolve で親ごとスキップ |
| brand | 空文字 | Bは `ノーブランド品` 固定 |
| gtin | JAN 等 | あれば載せる |
| mainImageUrl | 任意 | Aの Resolve では画像必須にしていない |
| inventory | 0 または 1 | マスタ非書込 |

**C1 `c1_packaged.py` は offer 行を前提にしていない**（親子 HPC 埋め）。M2 PACKAGED は **別ツール／別テンプレ**。

---

## 4. PACKAGED 方針（未決・社長確認）

### 4.1 候補（いずれか1つに固定する）

| 案 | 内容 | 向き |
|----|------|------|
| **案L** | Seller Central **在庫ファイル／Listing Loader** 系（オファー中心・列少ない） | §8「Offer Only / Inventory Loader」に近い |
| **案P** | Product Type 純正 xlsm の **部分更新（オファー列のみ）** | C1に近いが列マップが別・過大になりやすい |
| **案H** | 当面 **GENERATED→人間がSC手貼り**（自動化は記録＋チェックリストのみ） | 最短で §12.4 の一部を疑似完了。本線自動化ではない |

**推奨（起草）**: まず **案L** を本線候補とし、SCで公式テンプレ名を1つ採取→列マップ→ローカル変換（C1同様 Python可）。案Hはスパイク／並行逃げ道。

### 4.2 開いている問い（承認前に閉じる）

1. 公式テンプレの正式名（日本語 UI 表記）は何か  
2. 必須列の最小セット（SKU／ASIN／価格／数量／コンディション／配送グループ 等）  
3. 在庫0オファーを SC が受理するか（M1 U5と同型の見え方確認）  
4. すでに自社が出している ASIN への二重オファー防止は人間選定か／コードか  

---

## 5. ブランドゲート（GAP詳細）

要件: 認証・出品制限 → `SKIPPED_BRAND_GATE`（押し切らない）。

現行: 当該 `reason` を返す処理なし。BOTH 時も ASIN 有無だけで A/B 分岐。

**M2 v1 提案**:

| 方針 | 内容 |
|------|------|
| v1最小 | マスタ列（例: 要確認・ブランドゲートメモ）または Script Property の **禁止ASIN／禁止ブランドリスト** でスキップ。無ければログ警告のみで続行は **禁止**（誤出品リスク） |
| v1推奨 | 試験SKUは **人間がゲート無しと確認済み** のみ。コードは `SKIPPED_BRAND_GATE` スタブ＋Property `APPROVAL_AMAZON_LV4_BRAND_GATE_MODE=manual_ok` 等 |

実装承認時に採用案を1つに固定。

---

## 6. 試験SKU選定基準（定義）

人間が1親（または1子）を選ぶ。**自動選定はしない**。

| # | 条件 | 必須 |
|---|------|------|
| 1 | マスタに **ASINコード** あり（親または出品対象子） | 必須 |
| 2 | マスタ **在庫数（出品用）=0**（販売中スキップ回避） | 必須 |
| 3 | Amazon 承認①で当該親／子が **APPROVED** | 必須 |
| 4 | **自社がまだオファーしていない**（または試験用に上書き可と社長判断） | 必須 |
| 5 | ブランド認証・出品制限に当たらない（目視） | 必須 |
| 6 | `販売価格amazon` が数値として有効 | 必須 |
| 7 | 可能なら **単品または子1件**（バリエーション複雑さを避ける） | 推奨 |
| 8 | HPC／FOOD 等 PT は問わない（オファーはカタログ既存前提） | — |

選定結果は HUMAN_RUN §0 の表に追記（SKU／ASIN／選定日）。

---

## 7. 推奨実装スライス（承認後）

| 順 | スライス | 内容 | コード |
|----|----------|------|--------|
| M2-0 | 本ギャップ＋承認 | docs | なし |
| M2-1 | GENERATED 実機 | `TRACK=A`・SKIP_EXPORT→本出力・ログ確認 | 改修はギャップ埋めのみ |
| M2-2 | PACKAGED v1 | 案Lテンプレ＋列マップ＋dry_run/prod | **新規**（C1流用禁止） |
| M2-3 | SC＋記録 | 手動UP→サマリ→21-③ or E-5 | 手順 |
| M2-4 | ブランドゲート | 採用案のコード化 | 小〜中 |

---

## 8. 聖域・禁止

- 楽天 CSV／Yahoo 出品 API 本体  
- マスタ在庫・JAN の一括書込  
- C1 HPC テンプレへの offer 無理埋め  
- 販売中SKU無人上書き（U1）  
- TRACK 未設定のまま実行  

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-28 | **実機合格**: 公式 ListingLoader xlsm→SC成功／E-5。簡易CSV単独UP不可。残GAP＝自動埋め。 |
| 2026-07-27 | **v1実装**: 案Lツール／競合ASIN解決／manual_ok。試験=発汗。 |
| 2026-07-27 | 初版。GENERATED(A)は骨格OK・PACKAGEDとブランドゲートが主GAP。 |
