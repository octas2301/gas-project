# E. 楽天ジャンル／Yahoo cat — 本線確定

**文書種別**: 要件＋実装メモ  
**日付**: 2026-08-12  
**状態**: **本線化実装済**（7.5／7.6 分岐込み。要 `clasp push`）  
**親**: [RAKUTEN_NAV_GENRE_STAGE3.md](../RAKUTEN_NAV_GENRE_STAGE3.md)／[YAHOO_CATEGORY_BRAND_STAGE.md](../YAHOO_CATEGORY_BRAND_STAGE.md)／メニュー8

---

## 確定回答

| ID | 確定 |
|----|------|
| E-1 | A: B でジャンル／Yahoo **常時ON**（未設定＝ON）。**B のメニュー8は Step7.6**（売れ筋のみ） |
| E-2 | A: トグル独立 |
| E-3 | **上書き**: 都度API成功時に AI／旧値を差し替え（既存コードどおり） |
| E-4 | 当面件数制限なし。試験はレ点 **3〜10親** |
| E-5 | A: 不合格は書かない＋要確認のみ（**7.6 ブランドは 38074 フォールバック可**） |
| E-6 | OK（食品3件／SHP存在／要確認で壊れない／時間） |

### Yahoo モード（追加確定）

| 入口 | Yahooカテゴリ | Yahooブランド |
|------|---------------|---------------|
| Z **7.5** | `price_aware`（価格参照・従来） | 市場投票／AI・Drive → SHP |
| Z **7.6**／**B Step7.6** | `popular_only`（売れ筋重みのみ・価格非参照） | メーカー名 SHP 優先 → **`38074`** |

---

## 正本の流れ

```text
Step5 AI出品取得
  → ジャンル／Yahooカテゴリ・ブランドの AI生成・Drive CSV候補は既定スキップ
  → （KW・説明・梱包等は従来どおり Gemini/OpenAI）

Step6〜7
  → ジャンル作業エリアは空のままになりやすい

Step7.6 メニュー8（B）／Z 7.5 or 7.6
  → KW・trim
  → 楽天 Ichiba投票→Nav（正本）→ 楽天ジャンルID 上書き可
  → Yahoo: モードに応じ市場＋SHP検証 → Yahooカテゴリ／ブランド 上書き可
```

---

## Property（未設定＝ON）

| Key | 未設定 | 緊急停止 |
|-----|--------|----------|
| `AMAZON_AI_ADOPT_RAKUTEN_GENRE_ENABLED` | **ON** | `false` |
| `AMAZON_AI_ADOPT_YAHOO_CATEGORY_BRAND_ENABLED` | **ON** | `false` |
| `B_STEP5_SKIP_GENRE_YAHOO_AI_ENABLED` | **ON**（Step5でジャンルAIスキップ） | `false` で旧AI生成に戻す |
| `B_INTEGRATED_MENU8_ENABLED` | **ON**（Bの7.6） | `false` |

**注意**: 以前試験で `=false` を書いたままなら、**キー削除**または `true` にしないと本線にならない。

手動だけ: `AMAZON_AI_AUTO_ADOPT_ENABLED=true`（終わったら false）。7.5＝価格参照／7.6＝売れ筋のみ。ジャンル／Yahooトグルは戻さなくてよい。

---

## 復元

- 各トグル `false`
- Step5旧挙動: `B_STEP5_SKIP_GENRE_YAHOO_AI_ENABLED=false`
- Bを旧7.5価格参照に戻す: Git revert（または非推奨 `menuAmazonAiAdoptForBIntegratedStep_` を配列に戻す）
- Git revert

---

## 人間手順（要約）

1. `clasp push` → シート再読込  
2. 旧 false キーがあれば削除  
3. レ点 3〜10親で **B** または **Z 7.6** → Yahooが売れ筋のみ＋ブランド38074寄り  
4. （比較）Z **7.5** で価格参照パスの確認  
