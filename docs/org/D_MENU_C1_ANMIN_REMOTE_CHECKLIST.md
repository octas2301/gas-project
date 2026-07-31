# 安眠チェッカー — 外出先チェックリスト（2026-07-27）

**用途**: スマホ／外出先だけで進める確認用。Python・clasp は不要。  
**状態**: **完了**（SC合格＋E-5済・2026-07-27）  
**正本の流れ**: [D_MENU_C1_HUMAN_RUN.md](D_MENU_C1_HUMAN_RUN.md)／[D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)  
**現状ハブ**: [CURRENT_PHASE.md](../CURRENT_PHASE.md) §0

---

## 固定値（コピペ用）

| 項目 | 値 |
|------|-----|
| 親SKU | `lifec-4560151300924-oya` |
| 子SKU | `lifec-4560151300924-19s124` |
| **E-5／ログ用 subBatchId** | `A1_20260726_225610_4f0558_B2` |
| SC ファイル名 | `relax_PACKAGED_HPC_lifec-4560151300924-oya.xlsm` |
| SC Batch ID | `182816020660` |
| 処理サマリ | Downloads `relax_PACKAGED_HPC_lifec-4560151300924-oya-processing-summary.xlsm` |
| SC結果 | 処理2／成功2／失敗0／警告0／エラー総数0 |

---

## A. 外出先でやること（チェック）— **完了**

### A1. SC 結果確認 — **済**
- 処理サマリを Downloads に保存済み
- フィード処理結果: 処理2／成功2／失敗0／警告0／エラー総数0
- テンプレ: 親・子とも「成功」

### A2. E-5 — **済**
```text
A1_20260726_225610_4f0558_B2
```

### A3〜A4
任意確認・チャット報告は不要（記録は HUMAN_RUN へ反映済）

---

## B. 外出先ではやらない

| 作業 | 理由 |
|------|------|
| C1 Python／fetch／prod 再実行 | 自宅PC＋ローカルパス前提 |
| SC への再アップロード本番 | 失敗解析後に自宅で |
| Script Properties 変更 | 誤操作リスク |
| `clasp push` | ローカル認証 |

---

## C. 帰宅後（任意）

1. E 区間 Property を false  
2. （任意）`config.local.json` の `generated_csv` を `{subBatchId}_GENERATED.csv` に変更し `relax` 固定をやめる  
3. コミットは指示時  

---

## 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-27 | **完了記録**: SC合格＋E-5。サマリ=Downloads。 |
| 2026-07-27 | 初版。SCキュー送信後・外出先確認用。 |
