# LV4 — Amazon MAP：正本=sheet／派生=MD／実行=JSON（sheetから生成）

**状態**: **方針ロック＋実装**（2026-08-02）  
**HUMAN_RUN**: [D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md](D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)

---

## 1. 正本と派生（必須）

| 役割 | 置き場 | 誰が触る | 備考 |
|------|--------|----------|------|
| **編集正本** | スプレッドシート `▼設定(Amazonマッピング)` の **MAP**（＋ERRORS） | 人間＋Cursor（ルール蒸留） | 列対応・変換・禁止・既定の実行ルール |
| **派生（記録）** | `docs/org/MAP_SC_ERROR_LEDGER.md` 等 | Cursor | SCエラー全文・調査・なぜ直したか。実行正本にしない |
| **実行用** | `tools/c1_hpc_packaged/*_column_map.json` | Agent（生成） | **sheet から生成**。手 pen は原則禁止 |

**一文**: **正本=sheet／派生=MD／実行=JSON（sheetから生成）**

---

## 2. 運用フロー

```text
人間: sheet（MAP）を直す
SC結果 → Cursor: MDに記録 + 結論を sheet MAP/ERRORS へ反映
PACKAGED直前 → Agent: sync_map_sheet_to_column_json.py（sheet→JSON）→ B-T1/C1
```

- MD → JSON の直接反映は **しない**（必ず sheet 経由）。
- SHELF（Browse網羅）は別系統（`sync_shelf_browse_to_map_sheet.py`）。本承認は **属性MAP**。

---

## 3. 同期範囲（v1）

sheet MAP 行（`attrKey`）から JSON へ反映する項目:

| JSON 領域 | 元（MAP列） |
|-----------|-------------|
| `defaults.*` | `defaultValue`（有効行・固定／既定があるとき） |
| `master_columns.*` | `masterColPrimary`／`masterColFallback`（`doNotUse` は除外） |
| `c1_quantity_policy` | `transform`（セット数を数値化／ユニット数＝セット数／サイズ名から重量を取る）＋ omit 系ノート |
| `xlsm_header_aliases` の日本語候補 | `scHeaderJa`／`scHeaderAlias`（既存英キーは維持） |
| `map_sheet_sync` メタ | 同期時刻・行数・差分サマリ |

触らない: 指紋・`cols` 数値（B-T1 項目名解決が正）・SHELF／browseIndex。

---

## 4. ツール

| コマンド | 用途 |
|----------|------|
| `python sync_map_sheet_to_column_json.py --profile grocery [--dry-run]` | sheet→`food_fish_grocery_column_map.json` |
| `python sync_map_sheet_to_column_json.py --profile seasoning [--dry-run]` | sheet→`food_seasoning_column_map.json` |
| `python append_map_sheet_error.py ...` | ERRORS 行追記（任意） |

---

## 5. 検収

- [x] 本承認に正本定義を明記
- [x] HUMAN_RUN に PACKAGED 前必須手順を記載
- [x] `sync_map_sheet_to_column_json.py --dry-run` で差分表示
- [x] grocery で apply 後、JSON に `map_sheet_sync` と数量ポリシーが残る（2026-08-02）

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | 起草・実装。正本=sheet／派生=MD／実行=JSON。 |
| 2026-08-02 | HUMAN_RUN §0b チェックリスト固め（PACKAGED前／SC失敗／SHELF分離）。 |
