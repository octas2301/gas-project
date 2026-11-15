# 楽天画像アップ — マスタ既存URLスキップ（枠単位）

**状態**: 2026-09-05 実装  
**関連**: [D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md](D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md)

## 正

マトリクスから R-Cabinet へ送るとき、マスタの当該「楽天メイン画像N／楽天サブ画像N」に **すでに `http://` または `https://` URL** があれば:

- その枠は **アップロードしない**
- マスタも **上書きしない**
- ログに `マスタ既存URLのためアップ省略`

## トグル

| Key | 未設定 | false |
|-----|--------|-------|
| `RAKUTEN_UPLOAD_SKIP_EXISTING_MASTER_URL` | **ON** | 旧どおり常に送信 |

## 対象外

- D「画像アップロードをスキップ」（工程全体）とは別
- 同一実行内の Drive fileId キャッシュ（従来どおり）
- 楽天API照会による既登録判定（しない・マスタURLのみ）

## 併記

C マトリクス1回の画像取得上限: **100枚**（`PROCESS_LIMIT`）
