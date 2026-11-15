# かつお caption 仮貼付（戻し可）

本線化の正本（列名・LOCK・シリーズ同一HTML・他商品汎用）: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

**シリーズ方針**: 内容量は `1袋あたり50g`。袋数SKUでも同じHTMLでよい。

対象商品: **食べるおだし（かつお）1袋**  
商品コード: `ishi-4906283045119-32s14-2`  
商品URL: https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-32s14-2.html  
ストア管理: https://pro.store.yahoo.co.jp/pro.octas

## 貼るファイル（テスト合意）

| 欄 | ファイル | 内容 |
|----|----------|------|
| **商品情報**（PCスマホ共通・検索） | （プレーン短文のみ） | HTML不可。いまの短文で十分 |
| **商品説明（PCのみ）** | `katsuo_1bag_caption_PASTE.html` | **日本語のみ**（CTA・画像・セット）。表示幅600px |
| **フリースペース1（PC）** | `katsuo_1bag_fs1_PASTE.html` | **EN / 中 / 韓 / TH / VI** 長文（NCR） |
| **フリースペース2（PC）** | `katsuo_1bag_fs2_PASTE.html` | **PT / Filipino** 長文 |
| **スマホ用フリースペース** | `katsuo_1bag_caption_PASTE_sp.html` | 日本語ブロック + **EN / 中 / 韓 / TH / VI**（PCのFS1相当をやや圧縮。PT/TLはPC FS2） |

再生成:

```powershell
cd C:\Users\takuy\Desktop\gas-project
python tools/yahoo_page_mock/_gen_katsuo_paste.py
```

### なぜこう分けるか

- **商品情報** … 検索用短文。HTML・画像・セットは置けない  
- **商品説明** … PCの見た目（CTA・説明画像・セット）。多言語は入れない（長さ・役割分担）  
- **FS1 / FS2** … PCの海外向け長文。5+2でSEA（TH/VI）もカバー  
- **スマホFS** … 日本語＋EN/中/韓/TH/VI。PT/TLは欄が足りないのでPCのFS2に任せる  

### 外国語が `?` / `〓` になる原因（確定）

Yahooは **日本語以外の直書きUTF-8** を壊すことがある。  
対策: 中・韓・タイは **数値文字参照**（例: `&#31616;`＝简）。VI/PTはアクセントなしASCII。  
**全文差し替え**しないと画面の `?` は消えない。

### スマホでタグが文字になる原因（確定）

`<font>` `<center>` は PC専用。スマホFSに入れるとタグが文字で出る。  
現行PASTEは `div` / `b` / `br` / `a` / `img` / `table` のみ。

## 手順（失敗しても戻せる）

### A. 貼る前（必須・バックアップ）

1. ストアクリエイターPro → 上記商品を開く  
2. **商品説明**・**FS1**・**FS2**・**スマホFS** の現在内容をそれぞれコピーし、バックアップ保存  
3. 最低限、商品説明は同フォルダ `BACKUP_BEFORE_PASTE.html` に上書き

### B. 仮貼付

1. **商品情報** … HTMLが入っていたら削除し、プレーン短文のみ（例は下記）  
2. `katsuo_1bag_caption_PASTE.html` → **商品説明（PCのみ）**  
3. `katsuo_1bag_fs1_PASTE.html` → **フリースペース1**  
4. `katsuo_1bag_fs2_PASTE.html` → **フリースペース2**  
5. `katsuo_1bag_caption_PASTE_sp.html` → **スマホ用フリースペース**  
6. 保存 → **反映**  
7. 本番URLで確認  
   https://store.shopping.yahoo.co.jp/octas/ishi-4906283045119-32s14-2.html  

確認ポイント:
- PC: CTA・画像・セットが崩れない／多言語はページ下（FS）側  
- FS1: 中・韓・タイが `?` でない  
- FS2: PT/Filipino が表示される（不要なら空に戻して長さ比較可）  
- スマホ: EN/中/韓/TH/VIが `?` でない／スクロールが許容か（PT/TLはPCのみ）  

### C. 戻す（失敗時）

1. バックアップを各欄に貼り直し  
2. 保存→反映  

### 商品情報の例（プレーンのみ）

```
だし香る醤油仕立ての、ふわりとほどける食感。国産かつおを使った食べるおだし。ご飯・おつまみ・サラダに。
```
（HTMLタグ・`<` `>` は禁止）

## 注意

- 出品GASは caption を上書きしうる。**仮貼付は検証用**  
- 旧 `katsuo_additional2_i18n_PASTE.html` は本テストでは使わない（`katsuo_1bag_fs2_PASTE.html` を使う）  
- CTA画像URL:  
  - https://shopping.c.yimg.jp/lib/octas/cta_gourmet_sets.jpg  
  - https://shopping.c.yimg.jp/lib/octas/cta_gourmet_bestsellers.jpg  
- 説明画像（本採用サブ T03/T04/T11）:  
  - https://shopping.c.yimg.jp/lib/octas/katsuo_body_t03_sizzle.jpg  
  - https://shopping.c.yimg.jp/lib/octas/katsuo_body_t04_aji.jpg  
  - https://shopping.c.yimg.jp/lib/octas/katsuo_body_t11_arrange.jpg  
  UP手順: `UP_ADDITIONAL_IMAGES.txt` / `HUMAN_IMAGE_TEST.md`  
