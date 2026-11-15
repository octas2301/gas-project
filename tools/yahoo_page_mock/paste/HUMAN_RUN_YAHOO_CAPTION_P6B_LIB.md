# Yahoo説明 P6b — 魚種専用画像を追加画像libへUP

正本: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

## 何をするか

1. まぐろ／ぶり本文3枚（T03/T04/T11）を圧縮準備  
2. **人間が** Pro「追加画像」へUP・反映  
3. GASが lib 存在チェック → caption に `maguro_body_*` / `buri_body_*` を自動差し込み  

楽天CDN直載せはしない（P6a欠画の教訓）。lib自動UP APIは未着手。

## ファイル名（固定）

| 魚種 | ファイル |
|------|----------|
| まぐろ | `maguro_body_t03_sizzle.jpg` / `t04_aji` / `t11_arrange` |
| ぶり | `buri_body_t03_sizzle.jpg` / `t04_aji` / `t11_arrange` |
| かつお | 既存 `katsuo_body_*`（変更なし） |

URL例: `https://shopping.c.yimg.jp/lib/octas/maguro_body_t03_sizzle.jpg`

## 手順

### A. 圧縮ファイル準備（ローカル・済ならスキップ可）

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/yahoo_page_mock/prep_odashi_lib_bodies.py
```

出力: `tools/yahoo_page_mock/out/lib_ready/maguro_body_*.jpg` と `buri_body_*.jpg`

### B. 追加画像UP（API推奨・Pro手UP不要）

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/yahoo_page_mock/upload_lib_images_api.py
```

公式 `uploadLibImage`。認証は出品スプシの Yahooマッピング B12–B15。  
成功後に公開URLが 200 になること:

- https://shopping.c.yimg.jp/lib/octas/maguro_body_t03_sizzle.jpg
- https://shopping.c.yimg.jp/lib/octas/buri_body_t03_sizzle.jpg

（従来の Pro 手UPも可だが本線にしない）

### C. GAS反映

```powershell
cd C:\Users\takuy\Desktop\gas-project
clasp push
```

まぐろ1袋・ぶり1袋に出品CK → **Yahoo説明をAPI反映**

ログ期待:
- `lib_slug_check slug=maguro ready=true`
- `imgSrc=fish_lib_maguro`（ぶりは `fish_lib_buri`）

Proで本文画像が魚種専用になっていること（かつお写真のままならUP未反映）。

## 切戻し

- 追加画像を消す／未UPのまま → 自動でかつおlibにフォールバック
- Property `YAHOO_CAPTION_RICH_P6_IMG_ENABLED=false`
