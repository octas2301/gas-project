# プロジェクト全体の位置づけと現在の開発フォーカス

**最終更新**: 2026-08-14（**B統合ハード死対策・実装済み（ローカル）**／clasp・Property・HUMAN_RUN待ち）
**読み方**: 次の Agent は `docs/AGENT_HANDOVER.md` の **§1.5・§2** に従い、**本ファイルを最初に読み**、続けて §2 の必読一覧でプロジェクト全体をインプットする。

---

## 0. セッション引き継ぎ（2026-08-14）

**場所**: **B統合ハード死対策** — `B済`キュー自動パック（最大3・シリーズ温存）・Step1後切断・当B集合・番犬非設置。  
**コミット**: 未（指示時）。  
**正本**: [org/B_HARD_DEATH_SCOPE_REQUIREMENTS.md](org/B_HARD_DEATH_SCOPE_REQUIREMENTS.md)／[HUMAN_RUN](org/B_HARD_DEATH_SCOPE_HUMAN_RUN.md)

### 次にやること（人間）
1. Script Properties: **`B_WATCHDOG_ENABLED=false`**
2. 残っている `runBWatchdogFromTrigger` トリガーを削除
3. `clasp push`
4. AIにストック可。Bを1回起動。1列目`B済`が付く。完了後に次パックが連鎖  

### Property
| Key | 推奨 |
|-----|------|
| `B_WATCHDOG_ENABLED` | **`false`（本対策中必須）** |
| `B_SCOPE_INSERTED_PRODUCTS_ONLY` | 未設定=ON |
| `B_STEP1_FORCE_SLICE_AFTER` | 未設定=ON |
| `B_STEP1_TOP_INSERT_ENABLED` | 上挿入するなら **`true`**／戻すなら `false` または削除 |

---

## 0t. セッション引き継ぎ（2026-08-14・上挿入）※履歴

**場所**: **B Step1 上挿入** — スプシ AK/IB 値固定＋テンプレ1–2＋GAS（Property 未設定=旧末尾）。  
**復元**: `_local_backup/pre_B_STEP1_TOP_INSERT_20260814_001210/`

### 済（当時）
- [多数決](org/B_STEP1_TOP_INSERT_THREE_REVIEW_MAJORITY.md)／[要件](org/B_STEP1_TOP_INSERT_REQUIREMENTS.md)／[HUMAN_RUN](org/B_STEP1_TOP_INSERT_HUMAN_RUN.md)

### Property（当時）
| Key | 推奨 |
|-----|------|
| `B_STEP1_TOP_INSERT_ENABLED` | 上挿入するなら **`true`**／戻すなら `false` または削除 |
| `B_PARENT_ROW_HEIGHT_ENABLED` | 未設定=ON／緊急 `false` |
| `B_DF_STRATEGY_FORMULA_COPY_ENABLED` | 未設定=ON |
| `B_WATCHDOG_ENABLED` | 本対策中は **false**（下記0節） |

---

## 0f. セッション引き継ぎ（2026-08-13・⑧f）※履歴

**場所**: **⑧f 列ロック確定実装**（基本ロジック＋例外列は要件 §11〜§13）。  

### 済
- **⑧f**: [要件](org/B_MASTER_CELL_COLOR_RULES_REQUIREMENTS.md) §11〜§13／[凡例](org/B_MASTER_CELL_COLOR_RULES_HUMAN_RUN.md)
  - `masterColorApplyLockedRow_`／`MASTER_COLOR_*_LETTERS_`／Step1適用
  - Z **15-㉑** 選択行／**15-㉒** 11〜775
  - スプシ: レ点行＋**11〜775** をロック色で塗色（API）
  - 親赤白 `#f4cccc`（白字なし）／エラー濃赤は別
- **⑧a＋⑧c** ほか（オレンジ未導入／条件付き書式非変更）

### Property
| Key | 推奨 |
|-----|------|
| `B_PARENT_ROW_HEIGHT_ENABLED` | 未設定=ON／緊急 `false` |
| `B_DF_STRATEGY_FORMULA_COPY_ENABLED` | 未設定=ON |
| `B_WATCHDOG_ENABLED` | 本対策中は **false**（§0） |

---

## 0s. セッション引き継ぎ（2026-08-12・シリーズ推定）※履歴

**場所**: **マスタシリーズ 不明禁止＋商品名推定**（7.6合格後）。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: `clasp push` → シート再読込
2. 「不明」の親: **Step6 同期** → `マスタシリーズ` が推定名 or 空
3. 緊急停止: `SERIES_INFER_FROM_NAME_ENABLED=false`

### 正本
- [org/SERIES_INFER_FROM_NAME_REQUIREMENTS.md](org/SERIES_INFER_FROM_NAME_REQUIREMENTS.md)

---

## 0y. セッション引き継ぎ（2026-08-12・7.5／7.6）※履歴

**場所**: **メニュー8 Yahoo 7.5／7.6**。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. clasp push → Z 7.6／B Step7.6 試験 → **7.6合格**

---

## 0z. セッション引き継ぎ（2026-08-12・E本線化）※履歴

**場所**: **E本線化**（楽天ジャンル／Yahoo cat＝メニュー8都度API）＋⑥ N列ASIN。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: `clasp push` → シート再読込
2. 旧 false キー削除
3. レ点 3〜10親で B または手動7.5

### 正本
- [org/D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md](org/D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md)

---

## 0a. セッション引き継ぎ（2026-08-12・⑥N列）※履歴

**場所**: **⑥ N列ASIN自動**（◎×ブランド＝メーカー → N列＋黄セル）。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: `clasp push` → シート再読込
2. 試験: Z→**15-⑮**
3. E: 要件確認 → **本線化済**（§0へ）

---

## 0b. セッション引き継ぎ（2026-08-12・メニュー8）※履歴

**場所**: メニュー8 **v1.13**（ログで判明: 列J=商品名ベースまで削っていた＋FO二重計上）。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: 版履歴で **7.5前**に戻す（商品名ベース回復）
2. 人間: `clasp push`（v1.13）
3. 7.5を **1回** → GD≒70–75・商品名ベース残存を確認 → Property false

### 正本
- [org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)（v1.13）
- [org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)

### Property
| Key | 推奨 |
|-----|------|
| `AMAZON_AI_AUTO_ADOPT_ENABLED` | **false**（手動7.5のときだけ一時 true） |
| `B_INTEGRATED_MENU8_ENABLED` | 未設定／true |

---

## 0b. セッション引き継ぎ（2026-08-10）※履歴

**場所**: Amazon出品は **本番運用可**（缶飯・相乗り等 SC出品中確認済。A3サブ画像も合格）。開発フォーカスはサブ画像楽天出口の品質改善など。  
**コミット**: 未（指示時）。  

### 次にやること
1. 人間: `clasp push`（常時ON既定のコード反映）
2. Script Properties: 下表の常時ONキーが **false のまま残っていれば削除 or true**（1回だけ）
3. （任意）サブ画像楽天の写真ルール改善は `photo_realism_rules.py`

### Property チェックリスト（人間・本番常時ONセット 2026-08-10）

**考え方**: 毎日の D 出品で Script Properties を切り替えない。  
**未設定＝ON**（コード既定）。既に `false` が入っているキーは **削除**するか **true に1回**書き換える。  
緊急停止だけ明示 `false`。

**本番で常時ON（未設定で可／明示 false のみ緊急停止）**
| Property | 推奨 |
|----------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | **true／未設定** |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | **true／未設定**（開始前確認ダイアログは残る） |
| `APPROVAL_AMAZON_LV4_ENABLED` | **true／未設定** |
| `APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED` | **true／未設定** |
| `APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_ID` | 監視フォルダID（必須） |
| `APPROVAL_AMAZON_LV4_TRACK` | **B**（新規GENERATED） |
| `APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS` | **true／未設定** |
| `AMAZON_SHELF_REGISTRY_FILE_ID` | Drive の `shelf_registry.json` ファイルID |
| `AMAZON_IMAGE_CANDIDATE_FOLDER_ID` | Drive 07（C本線） |

**常時OFFのまま（毎回ONにしない）**
| Property | 推奨 | 理由 |
|----------|------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY` | **false／未設定** | マスタ在庫で数量送信（承認②） |
| `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED` | **false／未設定** | PT空の新規だけ一時 true |
| `AMAZON_IMAGE_U2_ENABLED` | **false／未設定** | Cコースが実行時だけ一時ON |
| `AMAZON_U4_URL_EMBED_ENABLED` | **false／未設定** | 同上（C-2／自動U4経路） |

**任意**
| Property | メモ |
|----------|------|
| `APPROVAL_AMAZON_LV4_SC_SUMMARY_INTERVAL_MIN` | 既定15。疎にするなら60 |
| `YAHOO_OAUTH_WARN_DAYS` | 既定7 |
| `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED` | 未設定＝**true**。旧運用に戻すときだけ `false` |
| `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED` | 未設定＝**true** |
| `YAHOO_OAUTH_REDIRECT_URI` | アプリ登録と一致する戻り先URL |

### SEASONINGで確定した必須・落とし穴（再発防止）
| 項目 | 対応 |
|------|------|
| 商品の形式 | 既定 `粉末`（列134） |
| 商品の重量の単位 | 重量ありなら単位必須。既定 `グラム`（列166） |
| 検索キーワード | **1枠のみ**（空白結合）。5枠分割は SC **99016** |
| 画像URL | C1は `Amazon MAIN/PT URL` のみ。サブはU4で楽天→R2 |
| マッチングsheet | 楽天サブとAmazon PTの**二重ドラッグ不要**（MAIN白抜きのみ手動） |
| 冪等 | 同じ親の再GENERATEDは D内レ点「失敗後の再GENERATED」または 21-④ `UPLOAD_FAILED` 後 |

### M2 正本リンク
- [org/D_MENU_M2_HUMAN_RUN.md](org/D_MENU_M2_HUMAN_RUN.md)  
- [org/LV4_M2_IMPLEMENTATION_APPROVAL.md](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)  
- [org/LV4_M2_TRACK_A_GAP_ANALYSIS.md](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)  
- `tools/m2_offer_packaged/`  
- [org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md](org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)／`tools/spapi_smoke/`  
- [org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)／`tools/spapi_listings_write/`  
- [org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md](org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md)（v1.1）  
- [org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md](org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)  
- [org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md](org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)（v1.2c）  
- [org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md](org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md)／[org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md](org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)  
- [org/LV4_SPAPI_GAS_PUT_APPROVAL.md](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)／[org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)（v1.4 第1段・API実機合格／SC最終更新は反映待ち）  
- [org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)（v1.4 **第2段・実機合格（API）**）  
- [org/LV4_SPAPI_D_ENTRY_APPROVAL.md](org/LV4_SPAPI_D_ENTRY_APPROVAL.md)／[org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)（**Dレ点本線・相乗り自己発 dry_run／prod 実機合格**）
- [org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)（**レ点本線・拡大実機合格**）
- [org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md](org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)（**FBA属性・実機合格**）
- [org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md](org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)（**デュアル Phase1・検収OK**）
- [org/LV4_DUAL_OFFER_PHASE2_APPROVAL.md](org/LV4_DUAL_OFFER_PHASE2_APPROVAL.md)／[org/D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md](org/D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md)（**Phase2・検収OK**）
- [org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md](org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md)／[org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)（**レーンA1・検収OK**）
- [org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md](org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md)／[org/D_MENU_P2_DC_HUMAN_RUN.md](org/D_MENU_P2_DC_HUMAN_RUN.md)（**P2調査完了＋§7.4方針ロック**・③は低優先）
- [org/LV4_P1_FILE_MIN_APPROVAL.md](org/LV4_P1_FILE_MIN_APPROVAL.md)／[org/D_MENU_P1_HUMAN_RUN.md](org/D_MENU_P1_HUMAN_RUN.md)（**P1 File最少・検収OK**）
- [org/LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md](org/LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md)／[org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md](org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md)（**P4a読取PoC・実機合格**）
- [org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)（**P4b-a／b合格。P4b-c＝D新規ゲート方針**）
- [org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md](org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md)（**D新規ゲート＋Cursor手渡し・実装済**）
- [org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md](org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md)／[org/D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md](org/D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md)（**SHELF Browse 網羅・Node ルーティング**）
- [org/LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md](org/LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md)／[org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md](org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)／[org/MAP_SC_ERROR_LEDGER.md](org/MAP_SC_ERROR_LEDGER.md)（**正本=sheet／派生=MD／実行=JSON**）
- [org/AMAZON_DEV_ROADMAP_P0_P4.md](org/AMAZON_DEV_ROADMAP_P0_P4.md)
- [org/LV4_C_COURSE_CONSOLIDATION_APPROVAL.md](org/LV4_C_COURSE_CONSOLIDATION_APPROVAL.md)／[org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)
- [org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md](org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)

### 外出先チェックリスト
[org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md) — 安眠完了済み（履歴）  

### 正本・手順リンク
- [org/LV4_M2_IMPLEMENTATION_APPROVAL.md](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)（**M2 v1・実機合格**）  
- [org/LV4_M2_TRACK_A_GAP_ANALYSIS.md](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)／[org/D_MENU_M2_HUMAN_RUN.md](org/D_MENU_M2_HUMAN_RUN.md)（**公式Loader正**）  
- [org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)（安眠完了・履歴）  
- [org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)  
- [org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)／[org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)  
- [YAHOO_CATEGORY_BRAND_STAGE.md](YAHOO_CATEGORY_BRAND_STAGE.md)／[org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md](org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)／[org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md](org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)  
- [org/D_MENU_C1_HUMAN_RUN.md](org/D_MENU_C1_HUMAN_RUN.md)  
- [org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)  
- [org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md](org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)  
- [org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  
- [org/D_MENU_U4_HUMAN_RUN.md](org/D_MENU_U4_HUMAN_RUN.md)  
- [org/YAHOO_OAUTH_REAUTH_HUMAN_RUN.md](org/YAHOO_OAUTH_REAUTH_HUMAN_RUN.md)（Yahoo再認証半自動・C14／B17メール）  

---

## 1. プロジェクト全体（30秒）

| 項目 | 内容 |
|------|------|
| **目的** | 楽天・Yahoo・Amazon 向けに、リサーチ〜商品情報・価格・セット構成〜出品に必要なデータを **スプレッドシート＋GAS** で整備・自動化する。一人社長＋AI部門モデルは [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)。 |
| **クリティカルパス** | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) — リサーチ・見積もり → 出品情報 → 各モール出品。 |
| **実装の中心** | コード.js、主シートは ▼商品マスタ(人間作業用)、AI情報取得data、00_設定マスタ。 |
| **6領域・成果物** | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md)。 |

---

## 2. 現在のフェーズ（いま優先している開発）

- **フォーカス領域**: Amazon出品は **本番常時ONセット適用済**。販促 **2層減衰実装**（カレンダー`減衰中%`＋B中1%。restore=減衰中%・[要件§10.13/10.14](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)）。並行でサブ画像楽天品質。  
- **Lv4（Amazonバルク）**:  
  - HPC／FOOD／T2／U3／**U2**／**U4（楽天サブ流用）**: 実装済・運用合格。  
  - C1 SEASONING: 七味 **出品中**。  
  - C1 GROCERY: 缶飯 **出品中**（100521待ち解消・人間確認済）。  
  - 本線UX: [org/D_MENU_U2_HUMAN_RUN.md](org/D_MENU_U2_HUMAN_RUN.md)／[org/D_MENU_U4_HUMAN_RUN.md](org/D_MENU_U4_HUMAN_RUN.md)／[org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。  
- **楽天Nav Stage1**: 実機PASS。Propertyはfalse。  
- **Lv3（Yahoo）**: 2026-07-20 人間検収完了。  
- **Lv2（楽天）**: 2026-07-20 人間検収完了。  
- **Lv1（承認キュー）**: 2026-07-17 人間検収完了。  

- **このフェーズの完了条件（目安）**  
  - ~~缶飯新規カタログが公開~~ **済（人間SC確認）**。  
  - ~~サブ画像 Amazon（A3）~~ **合格（人間確認）**。  

- **並行・継続（後回し可）**: T3 ZIP自動／Dc API強化／レーンC・P3／ε。  

- **スコープ外（次モール着手前）**  
  - 承認②、販売中SKU無人上書き、clasp push 自動化。  
  - generateRakutenCSV 本体・Yahoo.js 出品API本体の改変。

---

## 3. 直近で確定した仕様・前提（docs 反映済み）

- **AI組織**: **ハイブリッド**（承認①掲載は軽い／本命は承認②在庫反映）→ 日中無人出品（**ユニーク画像50枚＋実働25分**で自動分割・原則12:00前）→ 夜確認。残リストは明示取消まで残し、マスタ無しは画面非表示（履歴残す）。在庫0原則。販売中(在庫>0)は原則スキップ。CS／問屋は送信禁止・下書き可。詳細は org 憲章・マトリクス・多数決メモ §4.2・§4.3。  
- **キーワード案・商品名案の本線は OpenAI**。Gemini 版の商品名案は **コードに存在するが本流メニュー未配線**。  
- **バリエーション単位**は **Gemini → OpenAI 補完**（◎ASIN 経路も同様）。**内容量パース**は Gemini 中心（ChatGPT 自動フォールバックなし）。  
- **11-③（runProductNameProposalsForRows）** は、**OpenAI 失敗時もマスタ商品名があればバリエーションに届く**（既定）。旧挙動は PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL=false。**B Step7** は **variation まで進み得る**（商品名が空でも単位・内容量のみ更新されうる）。  
- 詳細は [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)、[商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md)。

---

## 4. 次にやること（優先順）

1. （済）8/14 taper prod: Cinderellas b/s 22→18。**店頭ともに18%確認済**。  
2. 日次 `taper_send.py --poll --mail`（dry_run）。Windows タスク **`OctasAmazonTaperDryRun`** 09:00。安定後に `--prod`。  
3. **開発完了後**: タイムセール運用マニュアル化。  
4. 8/27 Smile apply リマインド。9/4 restore＝減衰中%（仮想: 14）。  
5. T3／ε／U7（B SP-API品番・Dc④⑤）は各ゲート承認後。  
6. remote `git push` は指示時。  
7. スコープ外（後で再提案）: B開始自動 apply／B終了自動 restore／数量メール `--send`／P1c-2／P1c-C／レーンA。  

---

## 5. 深掘りリンク（領域別）

| テーマ | ドキュメント |
|--------|----------------|
| **AI組織・承認** | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)、[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)、[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)、[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)、[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)、[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)、[org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md) |
| 楽天ジャンル Nav（並行） | [RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)（Stage0）、[RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)（Stage1）、[RAKUTEN_NAV_GENRE_STAGE3.md](RAKUTEN_NAV_GENRE_STAGE3.md)（**Stage3都度API・実装済・要push**） |
| Gemini / OpenAI・11-③ vs B Step7・429 | [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) |
| 商品マスタ人間作業エリア | [商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md) |
| 価格・送料・再③・B 統合 | [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md)、AGENT_HANDOVER **§8** |
| リサーチ・セット構成 | [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) |
| エージェント共通・必読一覧 | [AGENT_HANDOVER.md](AGENT_HANDOVER.md) **§2** |

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-12 | **B P1実装**: シリーズID／メーカー辞書A181／カタログスラッグ13。要 clasp push。[要件](org/B_P1_SERIES_MAKER_CATALOG_REQUIREMENTS.md)。 |
| 2026-08-12 | **B番犬P0実装**: ハード死後自動再開・Step5行CP・プルダウンTIME_SLICE・サマリ・進捗メニュー。要 clasp push。[要件](org/B_WATCHDOG_RESUME_REQUIREMENTS.md)。 |
| 2026-08-12 | **販促2層減衰**: カレンダー減衰中%＋B中1%オーバーレイ。restore=減衰中%。`減衰開始日`。fetchは現在%のみ。要 clasp push／`--schema-only`。 |
| 2026-08-10 | **本番常時ONセット**: PUT／ALLOW_PROD／LV4／SCサマリの未設定既定=true。MASTER_QTY・P4B・U2・U4は常時OFF。A1〜B2・A3合格反映。要 clasp push。 |
| 2026-08-02 | **セットMAIN Phase A 実装**（`tools/set_main_image`・スモークOK）。実素材はHUMAN_RUN。[承認](org/LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)。 |
| 2026-08-02 | **セットMAIN Phase A 方針ロック**（Pillow・07/C接続・一括・Amazon先）。実装は別承認。[承認](org/LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)。 |
| 2026-08-02 | **七味出品中確定**（48h待ち削除）。D内レ点＋ダイアログ **clasp push済**。缶飯は100521待ち。 |
| 2026-08-02 | **D内レ点実装**（失敗後再GENERATED・案A）。[承認](org/LV4_D_REMAKE_MENU_APPROVAL.md)。 |
| 2026-08-02 | **D内レ点方針ロック**＋MAP sync HUMAN_RUN固め。[承認](org/LV4_D_REMAKE_MENU_APPROVAL.md)／[MAP HUMAN_RUN](org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)。 |
| 2026-08-02 | **SHELF Browse 網羅**（抽出・MAP・P4b/D Node ルーティング・缶飯 GROCERY）。[承認](org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md)。 |
| 2026-08-02 | **D新規ゲート＋Cursor手渡し実装**（Drive `AMAZON_SHELF_REGISTRY_FILE_ID`・3者スキップ）。要 clasp push＋File ID設定。[承認](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-08-01 | **Amazon候補を1子SKU＝1枚**（C画像コース）。要 clasp push。[C HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。 |
| 2026-08-01 | **相乗り ASIN空 soft skip**（相乗りのみ含む）。`コード.js`。要 clasp push。[承認§2.1](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-08-01 | **D新規ゲート＋Cursor手渡し方針ロック**（docs）。実装は別承認。[LV4_D_NEW_PT_SHELF_GATE_APPROVAL](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-07-29 | **A（D入口版）実装**: D ラジオに既存相乗り dry/prod。要 clasp push→実機。 |
| 2026-07-29 | **A（D入口版）承認**（社長「承認」）。実装可・範囲は承認 §2・§3。 |
| 2026-07-29 | **A（D入口版）承認起草**: D ラジオで Amazon 新規／既存相乗り（dry_run／prod）。コードなし。 |
| 2026-07-29 | **SP-API v1.4 第2段 実機合格（API）**: 21-⑫ VALID／21-⑬ ACCEPTED（batch `…b7a053`）。SC目視は反映待ち。 |
| 2026-07-29 | **SP-API v1.4 第2段 実装**: 21-⑫⑬／承認①済抽出。 |
| 2026-07-29 | **SP-API v1.4 第2段 承認**（社長「承認する」）。実装可・範囲は承認 §2・§3。 |
| 2026-07-29 | **SP-API v1.4 第2段 承認起草**（承認①済→GAS直PUT・コードなし）。三者レビュー不要。 |
| 2026-07-29 | **SP-API v1.4 実機合格（API）**＋ENDPOINT正規化。SC最終更新は反映待ち。 |
| 2026-07-29 | **SP-API v1.4 実装**: 21-⑩⑪／`AmazonSpapiPut.js`。要 clasp push→実機。 |
| 2026-07-29 | **SP-API v1.4 承認起草**（GAS直呼び・コードなし）。三者レビュー不要。 |
| 2026-07-29 | **SP-API v1.2b 実機合格**＋親レ点出さないを完了扱い。 |
| 2026-07-29 | **SP-API v1.2c／v1.3 実機合格**: 子レ点→Drive→fetch-drive prod（`…48s11`）。 |
| 2026-07-28 | **SP-API v1.2c**: 21-⑧＝子SKUレ点のみ（選択行廃止）。 |
| 2026-07-26 | **メニュー8緊急修正**: sync除外。要 clasp push。 |
| 2026-07-26 | **メニュー8 v1実装**: Z→7.5・空欄のみ採用。次＝clasp push→HUMAN_RUN。 |
| 2026-07-26 | **メニュー8要件＋承認パッケージ**: Amazon AI生成＆一括採用（M-A・空欄のみ・要確認列ごと）。実装承認待ち。 |
| 2026-07-26 | **C1-1b実装**（master_csv・必須列・タックスはマスタ）。次＝未送信SKUでSC。 |
| 2026-07-26 | **C1列マップ下書き**（成功PACKAGED差分・SC必須列）。次＝C1-1b。 |
| 2026-07-26 | **C1実装承認**（`tools/c1_hpc_packaged`・HUMAN_RUN手順）。次＝実機。 |
| 2026-07-26 | **C1三点反映**（MAJORITY・URL空スキップ・親一式除外・指紋v1本番停止）。次＝実装承認。 |
| 2026-07-26 | **C1要件起草**（ローカル本線・HPC・DRY_RUN・03新規）。次＝3者。 |
| 2026-07-26 | **U4 実機合格**: 21-⑦ `U4_20260726_090920_1366af`・マスタ URL 確認。D冪等は想定内。 |
| 2026-07-26 | **U4 v1 実装**: 21-⑦・GENERATED Amazon URL優先。次=clasp push＋HUMAN_RUN。 |
| 2026-07-26 | **U4要件＋承認パッケージ起草**。次＝社長実装承認。 |
| 2026-07-26 | **T2再検証合格**: `80s10` URL単独・18320なし・店頭OK。T3急がない。候補退避 clasp push済。 |
| 2026-07-25 | **U2 実機合格**: ②③④・`02` MAIN1件。Dは冪等ブロック想定内（`LV4_20260725_094425_914290`）。 |
| 2026-07-25 | **C子レ点選定**＋候補`07` Property済。次=clasp push→HUMAN_RUN。 |
| 2026-07-25 | **U2三点＋社長回答反映**（MAJORITY新規）。コミット待ち。 |
| 2026-07-25 | **U2方針確定**: 案α本線・MAIN=sheet／02=出口・εバックログ。 |
| 2026-07-25 | **U3実機合格**＋**U2要件起草**。runId `LV4_20260725_072635_948892`。 |
| 2026-07-25 | **D×Amazon U3 v1**: D `amazon`/`full_amazon`・即時ファサード。clasp push待ち。 |
| 2026-07-25 | **D×Amazon U0クローズ**: 3者反映＋社長回答。手ZIP正／T3実装待ち／将来API必須。 |
| 2026-07-24 | **D×Amazon要件U0**: [org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md](org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md)。T3保留。次=社長承認→3者。 |
| 2026-07-24 | **T2 PoC成功**: runId `R2T2_20260724_221107_7f9cf7`・URL画像表示。トグルfalseへ。次=T3要承認。 |
| 2026-07-24 | **T2 clasp push済**: 8 files（`AmazonDriveImageExport.js`含む）。次=Property＋21-⑥→URL200→トグルoff。 |
| 2026-07-24 | **帰宅引き継ぎ**: §0全面更新。FOOD成功・Nav PASS・T2実装済・次=自宅 clasp push＋21-⑥。 |
| 2026-07-23 | **帰宅引き継ぎ**: HPC suburl試験UP中・FOOD v4再解析（再UP禁止）。§0全面更新。 |
| 2026-07-23 | **§11.0 HPCクローズ**: 1〜8・10＋U5。画像ZIP優先。正本 titlefix。FOOD／M2／21-⑤は別ゲート。 |
| 2026-07-22 | **帰宅**: R2＋8枚MAINは200確認。§0をURL埋め→SC再UP向けに更新。 |
| 2026-07-22 | **自宅IDE引き継ぎ**: DRY_RUN＋本GENERATED成功（..._B2）。当時はPACKAGED→SC。 |
| 2026-07-21 | **セッション終了（深夜）**: clasp push成功。当時は次＝ドライラン（完了済み）。 |
| 2026-07-21 | **Lv4**: docsを新subBatchIdに統一。clasp pushは invalid_rapt で未完了 → clasp login 後に再実行。 |
| 2026-07-20 | **Lv4 実装**: AmazonApprovalExport／ApprovalQueue amazon加算／メニュー21。次は人間検収。 |
| 2026-07-20 | **Lv4 第3回三点＋Q15/Q16**: TRACK未設定＝非実行／amazon抽出同一チケット／列メモ注記。次は実装承認。 |
| 2026-07-20 | **Lv4 Q11–Q14反映**（inventoryMode・親単位再試行・GTIN証跡シート・送料無料パターン）。方針Q&Aは一通り閉じた。次は三点レビュー再実施→実装承認。 |
| 2026-07-20 | **Lv4 Q7–Q10b反映**（3モール同一承認①・TRACK=B強制・出品者SKU=子SKU／メーカー品番）。次は残未決→実装承認。 |
| 2026-07-20 | **Lv4社長Q&A反映**（純正xlsm・D-1 PACKAGED・BはマスタJAN残し・再生成上書き＋ログ追記）。次は残Q（Lv1抽出等）→実装承認。 |
| 2026-07-20 | **Lv4三点レビュー反映**（[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md)）。在庫書込禁止・DONE分離・親SKU抽出・GTIN着手ゲート。次は実装承認。 |
| 2026-07-20 | **Lv4要件ドラフト**（[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)）。M1=Bノーブランド新規／M2=A既存。 |
| 2026-07-20 | **Lv3人間検収完了**（本番 runId LV3_20260720_112238_777998）。フォーカスを Lv4 Amazon へ。在庫列「在庫数／在庫数計算」分離を§8.2相当で記録。 |
| 2026-07-20 | **Lv2案A修正**: 一時レ点を親＋子に変更（バリエーション商品のシングルSKU誤認を解消）。CSV本体非改変。次は再検収（clasp push → 19-①）。 |
| 2026-07-19 | **Lv2実装**（RakutenApprovalExport.js・メニュー19・案A）。次は人間検収。手動楽天CSVは非改変。 |
| 2026-07-17 | **Lv1人間検収完了**（batch A1_20260717_224813_ad7e65 → APPROVED）。フォーカスを Lv2要件確認／実装承認へ。 |
| 2026-07-17 | **Lv2要件ドラフト**（[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)）。実装は Lv1検収後。 |
| 2026-07-17 | ハイブリッド・残リストA・実行分割仮既定（25分／ユニーク50）を§3に反映。 |
| 2026-07-17 | **Lv1実装**（ApprovalQueue.js・メニュー18・Web approval_queue）。次は人間検収。 |
| 2026-07-17 | **Lv1要件**追加（[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)）。 |
| 2026-07-17 | **Lv0最終承認**反映。モール順楽天先。フォーカスを Lv1 へ。上書きは当面手動(U1)。 |
| 2026-07-17 | **Lv別実装プラン叩き台**追加（[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、コードなし）。 |
| 2026-07-17 | 並行: 楽天ジャンル Nav Stage1 実装（専用診断シート追記・Property 既定オフ）。 |
| 2026-07-15 | **AI組織 Phase0**: 3者多数決反映・RUNBOOK（親1＋並列3）・多数決メモ追加。実装は次フェーズ。 |
| 2026-07-15 | **AI組織 Phase0**: org 憲章・承認マトリクスをフォーカスに。実装は次フェーズ。 |
| 2026-03-19 | 初版。全体＋現在フェーズの引き継ぎを本ファイルに集約。 |
| 2026-03-22 | OpenAI 失敗時の 11-③ バリエーション継続（既定）を §3 に反映。 |
