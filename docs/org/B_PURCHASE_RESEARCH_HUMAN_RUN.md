# 仕入れ検討① — 人間作業（複製書き・専用GAS）

**日付**: 2026-08-15  
**正**: [B_PURCHASE_RESEARCH_NOW.md](B_PURCHASE_RESEARCH_NOW.md)／[V1](B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md)／[TASK_STRUCTURE](B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)

出品用 `.clasp.json` は **触らない**。`clasp push` は親フォルダで走らせない。

---

## 0. いまの成果（ローカル・貼付用）

`tools/purchase_research_path3/out/ishihara_gate_paste.csv`  
石原143件。門: 通過65／落ち78（価格30・順位40・冷凍8）。閾値は仮（2000円・15万）。

---

## 1. 書き分け

| 置き場 | タブ | 中身 |
|--------|------|------|
| 競合DB `1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs` | `Keepaフル` | Keepa属性の正（W1から①が書く。90日。csv[]なし） |
| 競合DB 同上 | `メーカーマスタ` | 第2クエリ語。語があれば再採取しない |
| 調査複製 `1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE` | `①候補` ほか | 門の転記。Catalog生・モールヒットは置かない |

調査複製は `contact@octas2301.com` に編集者共有が必要。

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\purchase_research_path3
python apply_research_ss.py
```

---

## 2. 専用 GAS（clone 済）

scriptId: `1oRfQzsECsGbuft8S7ksuxfKwAgq-AgFdSse41lE2lJVP-rPb4qx4yNRr`  
ローカル: `purchase_research/`（親 `.clasp.json` は出品のまま）

**push は人間が PowerShell で行う。Agent は clasp push しない。指示するときは必ず:**

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp pull
clasp logs
clasp push
```

**中身**: 外注マクロは残す。①は `PurchaseResearch.js`（モリタ `/query` GET → `①ログ`）。

**初回 push 前に** GAS エディタ → プロジェクトの設定 → Script Properties:

Keepa 契約は **1本**。出品と同じキー文字列を、①プロジェクトの Property にも入れる（プロジェクトが違うので自動では入らない）。トークン残は出品 A と①で共用。

| Key | 値 |
|-----|-----|
| `KEEPA_API_KEY` | 既存の Keepa キー（出品と同じ） |

コンテナバインドのため SS ID Property は不要（`getActiveSpreadsheet`）。

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

**M3 人間（W3 GAS）:** `clasp push` のあと、メニュー「モリタ 取扱→門」。Keepaフル90日内は GET しない。stats空は落ち。目安: ①候補の門と大きく食い違わない。

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

---

## 3. やらない

- 出品 `コード.js` への①追加
- 親フォルダでの clasp push
- 原本調査 SS への書き
- 公式サイトの商品照合（ASIN突合）。カテゴリ語の採取だけは可（NOW）
