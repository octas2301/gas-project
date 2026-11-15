# Yahoo説明 P6c — 採用テーマから売れそう順の上位3

正本: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md) §0.1

人手で型・文案を増やさない。本文はマスタ **`楽天サブ画像1〜10`** に入っている当選サブ。見出し＝catch、補足＝salesTalk／subs、紹介＝それら＋魚種／主原料から機械が組む。内容量の数字は紹介に使わない（大容量は残す）。写真用overviewと「商品説明文AIから取得」は種にしない。画と文は同一 themeId。空欄・落選は pool に入れない。

## 何をするか

1. ローカルでテーマJSONを組む（AIが上位3）
2. 上位3の画像を Yahoo lib へUP
3. clasp push のあと Property を ON → 説明API反映

未設定／false のあいだは **P5の固定文のまま**（本番を壊さない）。

## 手順

```powershell
cd C:\Users\takuy\Desktop\gas-project
# 同一JANで親が複数（oya2/oya3）→ 親ごとにJSON・パックキーが分かれる
python tools/yahoo_page_mock/build_yahoo_caption_themes.py --jan 4906283045119 --prep-lib
python tools/yahoo_page_mock/upload_lib_images_api.py --p6c
```

パック参照は **親SKU → slug → JAN**。同一JANに親が複数あるとき JAN キーは載せない。

既存の上位3だけ文字数を合わせる（pool再取得しない）:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/yahoo_page_mock/build_yahoo_caption_themes.py --fit-only
```

PC（caption 5000）とスマホ用説明（10000）は同じ合わせ方。画と見出しは変えず、紹介文だけ枠に収める。

```powershell
cd C:\Users\takuy\Desktop\gas-project
clasp push
```

Script Properties: `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED=true`

出品CK → **Yahoo説明をAPI反映**

ログ期待: `[YAHOO_CAPTION_RICH] p6c_themes n=3 jan=…` と `p6c_fit field=caption|sp_additional ... len=…`（len が caption≤5000 / スマホ≤10000）。`p6c_skip_over` が出たらその欄は P5 のまま。

Proで見出し＝catch、その下に overview、画像が同じテーマであること。

## 切戻し

- Property `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED=false`（未設定＝OFF）
- 既存 P6a/P6b の lib 差し替えはそのまま
