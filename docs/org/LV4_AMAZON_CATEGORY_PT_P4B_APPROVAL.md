**状態**: **P4b-a／P4b-b 合格**（2026-08-01）。三点スキップ継続。**P4b-c**＝D新規先頭ゲートとして [LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md](LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md) にスコープイン（**実装済 2026-08-02**）。**P4b-d**＝参照ASIN検証＋複数ASIN多数決＋競合無しハイブリッド（**2026-08-08 方針ロック／実装**）  
**親・前提**: [LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md)（**P4a 実機合格**）  
**ロードマップ**: [AMAZON_DEV_ROADMAP_P0_P4.md](AMAZON_DEV_ROADMAP_P0_P4.md)（P4b＝カテゴリ／PTのマスタ書込・C1・D本線接続）  
**手順**: [D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md](D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)  
**列マップ**: [D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md](D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)／[D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)  
**三者レビュー**: **方針・実装ともスキップ**（社長明示 2026-08-01／P4b-dもスキップ）  

---

## 1. 目的

P4a で確認した SP-API 読取（Definitions／Catalog）を、**新規カタログ（レーンB＝C1）と将来の選定補助**に接続する。  
現場の「カテゴリ／PT決めが作業先頭」を、**楽天ジャンル必須依存なし**で進める。  
参考ASINは **マスタの耐久列（競合店ASIN／競合URL）** を集め、**一致チェック合格分だけ** Catalog し、browseNode で多数決する。競合が無い／全滅したら **JAN即決 → SHELF／楽天／Yahoo／名称の重み付き投票**。  
**Browse 未確定のとき PT だけ書かない**（FOOD＋Browse空の半端状態を禁止）。

| 段階 | 内容 | 成功定義 |
|------|------|----------|
| **P4b-0** | 本包の方針ロック | §2… **済**／マスタ競合改定 **済（2026-08-01）** |
| **P4b-a** | 提案書込（空セルのみ・トグルOFF既定） | マスタ競合（またはフォールバック）から PT／browse が入る |
| **P4b-b** | C1 がマスタ PT／browse を読む | **マップがある PT のみ**本線採用（当面 SEASONING／HPC）。空なら従来固定 |
| **P4b-c** | D新規先頭: PT空ならP4b→棚判定→無ければ停止 | **実装済**（[D新規ゲート承認](LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／Drive棚） |
| **P4b-d** | ASIN検証・複数多数決・競合無しハイブリッド | **2026-08-08**（誤ASIN 404／魚肉取り違え対策） |

---

## 2. 社長確定方針（2026-08-01 ロック／**2026-08-08 P4b-d 改定**）

| # | 論点 | 決定 |
|---|------|------|
| 1 | 対象の分離 | **P4b＝SP-API Product Type＋推奨ブラウズ**。手数料キー `amazon カテゴリー`（Gemini 15-⑫）は **触らない** |
| 2 | 書込列 | **`Amazon Product Type`**／**`Amazon Browse Node`**。無ければ人間が列追加 |
| 3 | 上書き | **空セルのみ書込**。既存値・手入力は尊重 |
| 4 | 候補の決め方（競合あり） | **マスタ耐久列の複数ASIN（最大5）**: 親＋レ点子の `競合店ASINコード`・競合URL抽出ASIN・自 `ASINコード` を重複除去して収集。（a）各ASINを Catalog（上限内）。（b）**一致チェック合格**したものだけ browseNodeId を投票。（c）最多得票の Node＋SHELF preferred PT を採用。（d）**作業台ASIN貼付の◎多数決はしない**（揮発のため）。（e）特定PT名の優先はしない（SHELF Node→preferred は可） |
| 4b | 正しさチェック | **不合格ASINは採用しない**（書込に使わない）: Catalog≠200、タイトル非類似、JAN不一致（マスタJANあり時）、Browse語と商品名のカテゴリ衝突（例: 肉×魚介）。ログに `REJECT`／`WARN` |
| 4c | 競合無し／競合全滅 | **Stage A**: JAN→Catalog identifiers 検索→①相当チェック→即決。**Stage B**: SHELF browsePath×商品名（重み3）／楽天ジャンル・要確認（2）／Yahooカテゴリ・要確認（2）／メーカー・商品名ベース（1）の**重み付き投票**。閾値未満は Browse 未確定のまま停止誘導 |
| 4d | 書込契約 | **Browse 文字列が確定しない限り PT を書かない**（Definitions単独の PT 埋め禁止） |
| 5 | PTとC1 | Catalog／投票で Node 確定後、**SHELF**で preferredProductType を優先。C1はマップ実装済PT |
| 6 | C1接続 | マスタ PT／browse **両方必須**（grocery／seasoning の require）。空なら親除外。**json既定browseで埋めない** |
| 7 | D接続 | **P4b-c** 維持。D失敗はモーダル内表示に加え **alert ポップアップ** |
| 8 | 件数 | 1回あたり親 **最大3**。ASIN試行 **最大5／親**（Property可） |
| 9 | Keepa | **新規取得しない** |
| 10 | xlsm自動DL | **しない** |
| 11 | 実装順 | §2 P4b-d → コード → HUMAN_RUN／LEDGER |
| 12 | 三点 | スキップ |

---
## 3. 変更予定ファイル

### 3.1 方針まで — docs（済）

| 種別 | パス | 内容 |
|------|------|------|
| **新規** | `docs/org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md` | 本承認包 |
| **新規** | `docs/org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md` | 手順・検収欄 |
| 更新 | P4a承認／ROADMAP／PHASE／HANDOVER／LEDGER | 誘導 |

### 3.2 実装（マスタ競合版・2026-08-01 実装承認）

| 状態 | パス | 内容 |
|------|------|------|
| **済** | `AmazonCategoryPt.js` | 競合店ASIN／URL／自ASIN → Catalog1回＋WARN。多数決削除 |
| 済 | `コード.js`／`AmazonSpapiPut.js`／C1 | メニュー・GET共有・マスタPT優先 |

**戻し**: Property OFF。git revert。

### 3.3 やらない（本包）

- 手数料列 `amazon カテゴリー` の Gemini 置換・削除  
- Keepa の新規取得・メニューAの D 埋め込み  
- **作業台**ASIN貼付◎多数決の復活（マスタ耐久列の複数ASIN投票は P4b-d で可）  
- FISH 等の C1本線（マップ未）以外の拡大  
- レーンC／P2-③／P3・楽天聖域・Yahoo・B統合  

---

## 4. 仕様（契約）

### 4.1 用語の分離（混同防止）

| 名前 | 用途 | P4b |
|------|------|-----|
| `amazon カテゴリー` | 手数料 VLOOKUP キー | **非対象** |
| `Amazon Product Type` | Listings／C1 の商品タイプ | **対象（提案）** |
| `Amazon Browse Node` | 推奨ブラウズ | **対象（提案）** |
| `競合店ASINコード` | マスタの競合代表ASIN | **参考ASIN正本（優先）** |
| `競合AmazonページURL` | 人間調査の競合URL | **ASIN抽出・信頼メモ** |
| `ASINコード` | 自社／相乗り等 | **競合が空のときのフォールバック** |
| `Amazon PT URL` | 画像 | **非対象** |

### 4.2 トグル（実装時）

| Property | 既定 | 意味 |
|----------|------|------|
| `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED` | **false** | マスタへの書込を許可 |
| `APPROVAL_AMAZON_P4B_PT_MAX_PARENTS` | 3 | 1実行の親上限 |

### 4.3 fail-closed／警告

- LWA／403／候補0 → **書込しない**  
- 既存セル非空 → スキップ  
- 競合正しさチェック不合格 → **そのASIN不採用**（他候補へ）  
- Browse 未確定 → **PTも書かない**  
- Keepa 新規取得はしない  
- C1 FOODマップ: マスタ空または本線非許可 PT → **除外（fail-closed）**。HPCマップ: 空／非許可 → 現行 defaults  

### 4.4 P4a からの引き継ぎ

| # | P4a結論 | P4bでの扱い |
|---|---------|-------------|
| 1〜3 | Definitions／Catalog 可 | マスタ競合ASIN（または自ASIN）の Catalog／Definitions |
| 4 | Keepa | P4b本線では使わない（新規も既存貼付も依存しない） |
| 5 | xlsm自動DL不可 | 追わない |

---

## 5. リスクと緩和

| リスク | 緩和 |
|--------|------|
| 誤った競合ASIN | WARN（JAN不一致等）。人間URL精査を信頼。止めない |
| 子行ごとに別競合ASIN | 親→レ点子の代表1つに解決ルール固定 |
| C1に未対応PT | C1はマップPTのみ採用 |
| Keepa token | 新規取得禁止 |
| 上書き事故 | トグル既定false・空のみ |

**戻し**: Property OFF。git revert。

---

## 6. 検収

- [x] §2 方針承認… **2026-08-01**  
- [x] §2 多数決改定… **2026-08-01**（履歴）  
- [x] §2 **マスタ競合改定**… **2026-08-01**（多数決廃止・WARNのみ）  
- [x] 三点スキップ… **2026-08-01**  
- [x] 実装承認・初回／多数決コード… **2026-08-01**（旧）  
- [x] **実装承認**（マスタ競合版）… **2026-08-01**  
- [x] マスタ競合版コード… **2026-08-01**  
- [x] P4b-a 再実機（マスタ競合）… **OK** `P4B_PT_20260801_161210_546de6`／HERB／`B01N5A6ESU`  
- [x] P4b-b C1（**HERBのまま**・defaultsフォールバック）… **OK** `C1_CK_daba393f8055_B2_20260801_073624`／xlsm PT=`SEASONING`  
- [x] docs  
- [ ] トグル false 戻し  

---

## 7. 社長確認

- [x] §2 初回ロック… **2026-08-01**  
- [x] §2 多数決改定… **2026-08-01**  
- [x] §2 マスタ競合改定… **2026-08-01**（チェック＝警告のみ）  
- [x] 実装承認＋三点スキップ（旧コード）… **2026-08-01**  
- [x] マスタ競合版の実装承認… **2026-08-01**  
- [x] P4b-a 実機… **2026-08-01**  
- [x] P4b-b 実機（HERBフォールバック）… **2026-08-01**  
- [ ] トグル戻し確認  

---

## 8. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-08 | **P4b-d**: 複数ASIN＋一致チェック＋browse多数決。競合無しはJAN→重み付きSHELF投票。Browse無しPT禁止。Dエラーalert。 |
| 2026-08-01 | **P4b-c** を D新規ゲート承認へスコープイン（方針のみ）。 |
| 2026-08-01 | 起草〜マスタ競合コード。 |
| 2026-08-01 | P4b-a合格（rival_asin／HERB）。 |
| 2026-08-01 | **P4b-b合格**（マスタHERB→C1 dry_run PT=SEASONING）。ローカルPythonはAgent実行とHUMAN_RUN明記。 |
| 2026-08-02 | **C1 FOOD改定**: 許可PTにFOOD追加。PT/browse必須・既定browse禁止。ハイライト=楽天→Yahoo→箇条書き①。 |
| 2026-08-02 | **SHELF Browse 網羅**: Node ID ルーティング・preferredPT。MEAT固定エイリアスはフォールバック。 |
