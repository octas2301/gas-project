# Yahoo説明 API反映（案A'）人間手順

正本: [docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](../../../docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

## P6b（魚種専用lib UP）

→ [HUMAN_RUN_YAHOO_CAPTION_P6B_LIB.md](HUMAN_RUN_YAHOO_CAPTION_P6B_LIB.md)

---

## P6a（本文画像自動差し込み）

→ [HUMAN_RUN_YAHOO_CAPTION_P6_IMG.md](HUMAN_RUN_YAHOO_CAPTION_P6_IMG.md)

---

## P5（変数自動埋込）

→ [HUMAN_RUN_YAHOO_CAPTION_P5_VARS.md](HUMAN_RUN_YAHOO_CAPTION_P5_VARS.md)

---

## P3（シリーズ全袋）

→ [HUMAN_RUN_YAHOO_CAPTION_P3_SERIES.md](HUMAN_RUN_YAHOO_CAPTION_P3_SERIES.md)

---

## 事前

```powershell
cd C:\Users\takuy\Desktop\gas-project
clasp push
```

Script Properties: `YAHOO_CAPTION_RICH_LOCK_ENABLED=true`  
（意味: テンプレがある商品は caption/FS に生成HTMLを**同送**。省略しない）

## 使い分け

| メニュー | 用途 |
|----------|------|
| **Yahoo説明をAPI反映** | 初回・テンプレ差し替えの投入（必須項目＋装飾） |
| **▶ 出品実行** | 日常出品。Property ON＋テンプレ解決可なら装飾HTMLも毎回同送 |

## 検収（出品実行）

1. かつお子SKU 1本だけ出品CK → **▶ 出品実行**
2. ログ: `[YAHOO_CAPTION_RICH] inject … keys=caption,additional1,additional2,sp_additional`
3. 公開URLに CTA／`katsuo_body_t03` 等（消えない）

## 切戻し

Property `YAHOO_CAPTION_RICH_LOCK_ENABLED=false`
