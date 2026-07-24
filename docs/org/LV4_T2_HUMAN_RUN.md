# 21-⑥ Drive→R2（T2 PoC）人間手順

**実装**: `AmazonDriveImageExport.js` / メニュー **21-⑥**  
**承認**: T2のみ（2026-07-24）。T3以降は別。  
**検収**: **2026-07-24 PASS**（runId `R2T2_20260724_221107_7f9cf7`・公開URLで画像表示）。  
**Agent**: `clasp push` はしない（人間がローカルで実施）。再実行時はトグルを必ず false に戻す。

---

## 1. 反映

```text
clasp push
```

スプレッドシートを再読み込みし、**Z → 21 → 21-⑥** があることを確認。

## 2. Script Properties

| キー | 値 |
|------|-----|
| `AMAZON_DRIVE_R2_UPLOAD_ENABLED` | 実行時だけ `true`（終わったら `false`） |
| `AMAZON_DRIVE_R2_POC_SKU` | 例: `lifec-4560151300139-oya`（`{SKU}.MAIN.jpg` が Drive `02` 配下にあること） |
| `AMAZON_DRIVE_IMAGE_FOLDER_ID` | 省略可（既定=02のID）。別フォルダなら設定 |
| `R2_ACCOUNT_ID` | Cloudflare アカウントID |
| `R2_ACCESS_KEY_ID` | R2 APIトークンの Access Key |
| `R2_SECRET_ACCESS_KEY` | Secret（チャットに貼らない） |
| `R2_BUCKET` | 省略可 → `octas-amazon-imag` |
| `R2_PUBLIC_BASE` | 省略可 → `https://pub-d974bd81c7d84f9bbc65f8479d3f85d4.r2.dev` |

## 3. 実行

1. Property オン＋SKU設定  
2. **21-⑥** → OK  
3. 完了ダイアログの URL をブラウザで開き **HTTP 200** を確認  
4. `AMAZON_DRIVE_R2_UPLOAD_ENABLED=false` に戻す  

## 4. 失敗時

- 「無効です」→ Property が false  
- 「SKU未設定」→ `AMAZON_DRIVE_R2_POC_SKU`  
- 「見つかりません」→ Drive `02\親フォルダ\{SKU}.MAIN.jpg` の有無  
- R2 http 4xx → 鍵・bucket・アカウントID  

## 5. やらないこと

- T3（ZIP量産）・全件ループ・楽天03操作  
- Property に鍵をチャットへ貼る  
