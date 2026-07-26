# 安眠チェッカー — 外出先チェックリスト（2026-07-27）

**用途**: スマホ／外出先だけで進める確認用。Python・clasp は不要。  
**正本の流れ**: [D_MENU_C1_HUMAN_RUN.md](D_MENU_C1_HUMAN_RUN.md)／[D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)  
**現状ハブ**: [CURRENT_PHASE.md](../CURRENT_PHASE.md) §0

---

## 固定値（コピペ用）

| 項目 | 値 |
|------|-----|
| 親SKU | `lifec-4560151300924-oya` |
| 子SKU | `lifec-4560151300924-19s124` |
| **E-5／ログ用 subBatchId** | `A1_20260726_225610_4f0558_B2` |
| SC ファイル名 | `relax_PACKAGED_HPC_lifec-4560151300924-oya.xlsm`（接頭辞 `relax` は設定残り・中身は安眠） |
| SC Batch ID | `182816020660` |
| 送信時刻目安 | 2026-07-27 午前1:16（JST）・当時「キューに入っている」 |

---

## A. 外出先でやること（チェック）

### A1. SC 結果確認（必須）

1. Seller Central → **アップロードのステータスの確認**
2. 上記ファイル名 or Batch `182816020660` を開く
3. 記録する（メモ or スクショ）:
   - [ ] ステータス（完了／エラー 等）
   - [ ] SKU成功／失敗／警告／エラー総数
   - [ ] 処理サマリがあれば要点1行

**合格の目安（前回 relax 0405 と同様）**: 処理SKU≒2・成功≒2・失敗0・エラー総数0

### A2. E-5（SC成功後・Sheetsが使えるとき）

1. スプレッドシート → メニュー **E. Amazon出品コース（一時）** → **E-5**
2. 入力する ID（これだけ）:

```text
A1_20260726_225610_4f0558_B2
```

3. **入れないもの**: `relax` ／ Batch ID `182816020660` ／ファイル名全体  
4. メニュー E が無い／動かない → 帰宅後（未 `clasp push` の可能性）

### A3. 任意（見た目確認）

- [ ] マスタで親・子行の Amazon URL／ASIN が空でないか  
- [ ] カタログ上で新規SKUが見えるか（反映遅延あり）

### A4. チャットへ返す文面テンプレ

```text
SC結果: 成功=__ / 失敗=__ / 警告=__ / エラー総数=__
Batch: 182816020660
E-5: 済 or 未（理由）
```

---

## B. 外出先ではやらない

| 作業 | 理由 |
|------|------|
| C1 Python／fetch／prod 再実行 | 自宅PC＋ローカルパス前提 |
| SC への再アップロード本番 | 失敗解析後に自宅で |
| Script Properties 変更 | 誤操作リスク |
| `clasp push` | ローカル認証 |

失敗・大量警告時は **文言をメモして帰宅**。再PACKAGEDは自宅。

---

## C. 帰宅後（自宅PC）

1. E-5 未なら実行（subBatchId は上表）
2. Property トグルを使い終わったら false（E 区間）
3. （任意）`config.local.json` の `generated_csv` を `{subBatchId}_GENERATED.csv` に変更し `relax` 固定をやめる
4. HUMAN_RUN／CURRENT_PHASE に SC合格を追記・コミットは指示時

---

## 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-27 | 初版。SCキュー送信後・外出先確認用。 |
