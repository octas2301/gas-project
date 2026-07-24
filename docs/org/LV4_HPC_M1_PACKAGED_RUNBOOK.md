# Lv4 HPC（M1）PACKAGED〜検収 手順書（1枚）

**文書種別**: 運用手順（HPC／HEALTH_PERSONAL_CARE 成功ルートの再現用）  
**最終更新**: 2026-07-24  
**正本の詳細**: [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §11.0・§6.1.2・§11.5  
**画像設計**: [LV4_R2_IMAGE_PIPELINE_POC.md](LV4_R2_IMAGE_PIPELINE_POC.md)  
**状態**: HPC M1 は **クローズ済**。本手順は同型バッチの再現／引き継ぎ用。

---

## 0. 固定ID（成功実績）

| 項目 | 値 |
|------|-----|
| subBatchId | `A1_20260721_083100_06b90a_B2` |
| 親SKU | `lifec-4560151300139-oya` |
| 親ASIN | `B0H9WZV641` |
| 正本 xlsm | `…_PACKAGED_corrective_titlefix.xlsm`（8/8） |
| Drive 置き場 | `04.amazon…\03`（最終）／`06`（純正 HPC）／`02`（MAIN/PT）／`04`（ZIP）／`05`（ログ） |

Property: `APPROVAL_AMAZON_LV4_ENABLED=false`（再GENERATED禁止）／`TRACK=B`

---

## 1. 事前ゲート

1. GTIN免除の証跡が状態シートにある（未記録カテゴリは B 保留）  
2. 純正テンプレは Drive `06\HEALTH_PERSONAL_CARE.xlsm`  
3. 画像は **`{SKU}.MAIN.jpg`** を Drive `02` に用意。本運用の正は **SC ZIP**（R2 URL 単独は 18320 になり得る）  
4. FOOD／他 PT／HPC新SKUと **ファイル・SKUを混ぜない**

---

## 2. PACKAGED ルール（HPC）

| 項目 | 値 |
|------|-----|
| 行 | 5=属性／6=サンプル維持／**7〜=実データ（7=親）** |
| テーマ | `サイズ`（日本語） |
| Browse | `ドラッグストア > 衛生用品・ヘルスケア > 検査キット (4520899051)` |
| ブランド | `ノーブランド品`。メーカー／タイトルに登録ブランド名を載せない |
| ハイライト | **空**（`title_differentiation`） |
| 電池 | 「必要ですか？」「含まれていますか？」とも **いいえ** |
| 修正登録 | **同一SKUのみ**（新SKU禁止）。アクション＝作成または置換 |
| 商品名 | 子は区別できること（titlefix）。親は共通名可 |

---

## 3. SC 手順（成功ルート）

```text
① PACKAGED xlsm を SC「商品スプレッドシート」へ 1 回 UP
② 続けて Upload Images で ZIP（中身 {SKU}.MAIN.jpg …）
③ 処理サマリで件数確認（例: 8/8）
④ 全在庫で親子バリエーションを目視
⑤ スプレッドシートメニュー 21-③ UPLOADED_OK（正しい subBatchId）
⑥ APPROVAL_AMAZON_LV4_ENABLED=false を確認
⑦ U5: 在庫0 → SCで在庫切れ・数量0
```

**部分成功時**

- 子だけ単品・親未リンク → 同一SKUで corrective（§6.1.2）→ 目視後 21-③  
- 100521 のみ → **再UPせず最大48h待つ**  
- 18320（画像）→ ZIP を本線に。URLだけに頼らない  

---

## 4. 成功条件（HPC検収）

- [ ] processing-summary が全SKU成功（または方針どおりの許容）  
- [ ] 親＋全子が同一バリエーション家族  
- [ ] 21-③ `UPLOADED_OK`  
- [ ] ENABLED=false  
- [ ] U5 在庫0確認  
- [ ] 成功値を `accepted_values_db/success/` に残す（Drive `05` 推奨）

---

## 5. やってはいけないこと

- 連打UP／SKU付け替え新規／楽天CSV・Yahoo.js への波及  
- Cloud からの `clasp push`／SC代行  
- ハイライトに箇条書きを入れる（仕様／説明へ）  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-24 | 初版（§11.0 クローズ内容を1枚化） |
