# 楽天マトリクス — 一時合算＋レ点スコープ除外

**文書種別**: requirements  
**状態**: 2026-09-05 チャット確定・実装済  
**最終更新**: 2026-09-05  

**対象**: `generateAiImageMatrix`（C「楽天のみ」）  
**関連**: [D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md](D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md)／[D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)

---

## 1. 目的

1. **追加再マップ**: 一度 C したあと 02 に画像を足して再実行しても、**一時フォルダの既存＋02新規**の両方でマッピングできる。
2. **別商品残留**: 一時に吉野家などレ点外 JAN が残っていても、**青枠に混入しない**（手退避を必須にしない）。

---

## 2. 確定仕様

### 2.1 ① 合算取得（`RAKUTEN_MATRIX_MERGE_WORK_FOLDER`）

| 項目 | 内容 |
|------|------|
| 未設定 | **ON** |
| 02直下 | 画像を一時フォルダへ移動（02クリア）。処理上限内なら fileList 先頭に載せる |
| 一時既存 | ON時は **同じ run で合算**（同一 fileId は1回）。02が空でも一時だけ、02があっても一時を足す |
| OFF時 | 旧挙動: 02に1枚でもあれば一時は読まない。02空なら一時のみ |
| 上限 | `PROCESS_LIMIT`（100）は合算後に適用。02新規を優先してから一時を埋める |

### 2.2 ② レ点スコープ除外（`RAKUTEN_MATRIX_SCOPE_FILTER_ENABLED`）

| 項目 | 内容 |
|------|------|
| 未設定 | **ON** |
| スコープ | 今回マトリクス行の親SKU・子SKU・それらに含まれる JAN |
| `out` | ファイル名の JAN／SKU がスコープ外 → **AIしない・青枠に入れない**（一時に残したままスキップ可） |
| `in` | スコープ内 → 従来どおりファイル名 autobind → 余り AI |
| `unknown` | JAN/SKU 手がかりなし → AI可。未マッチは **グレー候補のみ**（青枠へ強制投入しない） |
| 先頭行青枠投入 | **廃止**（別商品混入の主因だった） |

### 2.3 戻し

- `RAKUTEN_MATRIX_MERGE_WORK_FOLDER=false`
- `RAKUTEN_MATRIX_SCOPE_FILTER_ENABLED=false`

---

## 3. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-09-05 | 初版。合算＋スコープ除外を実装ロック。 |
