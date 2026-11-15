# Yahoo追加画像の元ファイル置き場

おだし系本文（かつお／まぐろ／ぶり）。採用落ち候補は置かない。

| ファイル | テーマ |
|----------|--------|
| `katsuo_body_t03_sizzle.jpg` | T03 シズル |
| `katsuo_body_t04_aji.jpg` | T04 味 |
| `katsuo_body_t11_arrange.jpg` | T11 アレンジ |
| `maguro_body_t03_sizzle.jpg` 等 | 同上（JAN 4906283045317 / theme_new_full） |
| `buri_body_t03_sizzle.jpg` 等 | 同上（JAN 4906283047410 / theme_new_full） |

```powershell
cd C:\Users\takuy\Desktop\gas-project
# まぐろ・ぶりを set_main_image 出力から取り込み＋圧縮
python tools/yahoo_page_mock/prep_odashi_lib_bodies.py
# または既存 assets だけ圧縮
python tools/yahoo_page_mock/prep_lib_images.py
```

出力: `tools/yahoo_page_mock/out/lib_ready/` → Pro「追加画像」へUP  
URL: `https://shopping.c.yimg.jp/lib/octas/ファイル名.jpg`  
手順: [HUMAN_RUN_YAHOO_CAPTION_P6B_LIB.md](../../tools/yahoo_page_mock/paste/HUMAN_RUN_YAHOO_CAPTION_P6B_LIB.md)
