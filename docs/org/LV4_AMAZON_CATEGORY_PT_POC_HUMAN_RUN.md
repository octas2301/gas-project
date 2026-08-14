# Amazon カテゴリ／PT 読取 PoC — 人間手順（P4a）

**状態**: **実機合格**（ローカルPython読取。マスタ非書込）  
**承認**: [LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md](LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md)  
**基盤**: [D_MENU_SPAPI_SMOKE_HUMAN_RUN.md](D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)（同一 LWA／`config.local.json`）

---

## 0. できること／できないこと

| できる | できない |
|--------|----------|
| LWA＋`searchDefinitionsProductTypes` | マスタ／承認ログ／SC への書込 |
| `getDefinitionsProductType`（候補1件） | 全件ループ |
| Catalog Items（分類系 `includedData` 拡張） | 純正 `.xlsm` の SP-API 自動DL（結論はレポート） |
| KeepaはメニューAで少件数（任意・token消費） | P4bのマスタ自動選定 |

---

## 1. 準備

1. `tools/spapi_smoke/config.example.json` → `config.local.json`（既存スモークと同じで可）
2. `pip install -r tools/spapi_smoke/requirements.txt`
3. （任意）`keywords` / `item_name` / `smoke_asin` を config に設定

```json
{
  "smoke_asin": "B07YND44VN",
  "keywords": "七味",
  "item_name": "",
  "product_type": ""
}
```

`product_type` 空なら search の先頭候補で getDefinitions する。

---

## 2. 実行

```text
cd tools/spapi_smoke
python spapi_smoke.py --poc-category -v
```

従来スモークのみ:

```text
python spapi_smoke.py -v
```

---

## 3. 見るログ／成果物

- コンソール: LWA／search PT／getDefinitions／Catalog の HTTP と要約
- `tools/spapi_smoke/out/SPAPI_CATEGORY_PT_*_REPORT.json`（トークン本文なし）
- レポートの `conclusions` に各#の可否メモ

Keepa（#4）: メニューAで同ASINを少件取得し、カテゴリ系フィールドの有無を本承認包またはレポートに手追記。

xlsm自動DL（#5）: SP-API Definitions は **JSONスキーマ**であり純正 `.xlsm` バイナリは取れない、をレポート結論に記載（実行不要・調査結論）。

---

## 4. 検収

- [x] searchDefinitionsProductTypes が JP で候補を返す（または権限エラーを記録）
- [x] getDefinitionsProductType が1件取れる（またはスキップ理由）
- [x] Catalog で classifications／productTypes 等が取れる／取れないを記録
- [x] マスタ未変更
- [ ] Keepaは任意・少件のみ（未実施可）

### 実機結果（2026-08-01）

| # | 結果 |
|---|------|
| 1 | **OK** `keywords=七味` → 候補 `HERB`, `SEASONING`（getDefinitionsは **SEASONING優先**。強制は config `product_type`） |
| 2 | **OK** `getDefinitionsProductType(HERB)` HTTP200・トップキー10 |
| 3 | **OK** Catalog `B07YND44VN` で productTypes／classifications／attributes 取得可（当該ASINは BODY_DEODORANT。競合参照の経路確認が目的） |
| 4 | Keepa: 未実施（任意） |
| 5 | **結論** xlsm自動DL APIは本経路に無し（Definitions＝JSON） |

レポート例: `tools/spapi_smoke/out/SPAPI_CATEGORY_PT_*_REPORT.json`（gitignore・署名URL含むためコミット禁止）

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | 初版。smoke 拡張 `--poc-category`。 |
| 2026-08-01 | 実機合格記録（#1〜3・#5）。Keepaは任意残。 |
