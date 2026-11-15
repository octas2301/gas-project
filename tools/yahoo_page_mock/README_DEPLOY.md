# Yahoo商品ページ装飾｜導入メモ（獲得URL待ち）

最終更新: 2026-09-13  
対象ストア: `octas`  
採用導線:  
1. **セットボタン**（`cta_gourmet_sets.jpg`）→ ページ内 **セットコーナー（#sets）**  
2. コーナーで **各セットのメイン画像（全件）** → 各セット商品ページ  
3. **売れ筋ボタン** → ストアTOP  

Drive正本: `G:\マイドライブ\03.楽天・Yahoo!商品登録（CSV一括UL）\04.クーポン画像`

## 販促施策（確定・2026-09-13）

方針: **集客 → 転換 → 客単価 → リピート** で役割を被らせない。  
原資前提: 実効利益はおおむね **1点あたり200円台**。まとめ割は送料浮き込み。

### 常設クーポン（少数）

| 役割 | 施策 | 条件 | 備考 |
|---|---|---|---|
| 転換 | **50円OFF** | **1,500円以上** | 公開・通常購入のみ |
| 客単価 | **5%OFF** | **2点以上** | 上限150円。画像 `coupon_2items_5off.jpg` |
| 新規 | **3%OFF** | 初来店・未購入 | 上限150円（画面に欄がある場合） |
| リピート | **サンキュー3%** | 購入者・非表示 | 円上限欄なし。主力500〜2000円想定 |

見送り: 999円以下50円／3点系／5点系／サンキュー定額50円（転換と重複）。

### 広告・その他

| 施策 | 設定 |
|---|---|
| PR | パッケージの2%相当は加入済。有料一律は原則なし（主力だけ短期テスト可） |
| アイテムリーチ | 主力のみ |
| 定期 | 5%（済）。都度クーポンは通常購入のみ |
| レビューPt | 定額のまま |

### 公開期間・リマインド（GAS）

- 四半期ごと。例: **〜2026/10/01 0:00**、次期 **〜2027/01/01 0:00**
- 発行の自動延長は不可。`PromoRemind.js` がシート「▼販促施策リマインド」で **2四半期分**を持ち、期限前にメール
- メニュー: `🛒 Yahoo!出品` → `販促リマインド（クーポン期限）`
- 宛先: `▼設定(Yahooマッピング)` **B17**
- 停止: Property `PROMO_REMIND_ENABLED=false` またはトリガー削除

### 重ね方（禁止・上限）

- 1注文のクーポンは **実質1枚** を想定して設計
- **初来店％ × PR3% × まとめ** の常時三重は禁止
- 常時の合計原資目安: 利益200円台に対し **おおよそ30〜50円／件**（まとめは送料浮き込みで別枠可）

### ページ装飾との対応

- **セット**: 上部は専用ボタンのみ。押下先はセットコーナー。コーナー内でメイン画像を全件（セット数2以上・昇順）リンク
- **売れ筋**: `cta_gourmet_bestsellers.jpg` → ストアTOP
- 単位はメイン画像側。セットボタンは「お得なセットを見る」（単位なし）
- クーポン条件画像は商品HTMLに埋め込まない
- `SETS_BLOCK` 自動生成は後段

## いまできること / 待ち

| 項目 | 状態 |
|---|---|
| セットボタン→コーナー→MAIN全件（方針・ひな形） | 済 |
| CTA画像をDriveへ保存 | 済（`04.クーポン画像`） |
| SETS_BLOCK自動生成（マスタ連携） | **後段** |
| 各セットYahoo URL／MAIN URL | **待ち** |
| Yahoo追加画像アップロード | **待ち**（人間） |

## 追加画像URL（クーポン）

| key | 画像URL |
|---|---|
| coupon_50yen_under999 | https://shopping.c.yimg.jp/lib/octas/coupon_50yen_under999.jpg |
| coupon_100yen_over1000 | https://shopping.c.yimg.jp/lib/octas/coupon_100yen_over1000.jpg |
| coupon_2items_2off | https://shopping.c.yimg.jp/lib/octas/coupon_2items_2off.jpg |
| coupon_3items_3off | https://shopping.c.yimg.jp/lib/octas/coupon_3items_3off.jpg |
| coupon_5items_5off | https://shopping.c.yimg.jp/lib/octas/coupon_5items_5off.jpg |
| coupon_first_visit_3off | https://shopping.c.yimg.jp/lib/octas/coupon_first_visit_3off.jpg |
| coupon_next_visit_3off | https://shopping.c.yimg.jp/lib/octas/coupon_next_visit_3off.jpg |

## 人間が追記する欄

| key | 説明 |
|---|---|
| CTA_SETS_IMG | 追加画像の `cta_gourmet_sets.jpg`（セットコーナーへの入口） |
| CTA_BESTSELLERS_IMG | 追加画像の `cta_gourmet_bestsellers.jpg` |
| STORE_TOP_URL | 固定 `https://store.shopping.yahoo.co.jp/octas/` |
| SETS_BLOCK | コーナー内：MAIN画像×全セットのリンクHTML（自動生成は後段） |

## ひな形ファイル

- `templates/katsuo_caption_v1.html` … 商品説明。上部ボタン＋#setsコーナー
- `templates/katsuo_caption_v1_draft_nclaim.html` … 貼付用ドラフト
- `templates/katsuo_additional1_i18n_v1.html` … フリースペース1（EN/ZH/KO）
- `templates/katsuo_additional2_i18n_v1.html` … フリースペース2（TH/VI/PT/TL）
- `templates/placeholders.json` … 置換キー一覧

## 貼付ルール

1. 上部セットボタンは `#sets`（コーナー）へ。メイン画像の直リンクはコーナー内のみ
2. セットは全件（2以上・昇順）。単位はMAIN画像側
3. CTA画像のDrive正本: `04.クーポン画像`
4. クーポン条件画像・獲得URLは商品説明に埋め込まない

## CTA画像の追加アップロード（人間）

Drive `04.クーポン画像` から Yahoo追加画像へ:

- `cta_gourmet_sets.jpg`
- `cta_gourmet_bestsellers.jpg`
