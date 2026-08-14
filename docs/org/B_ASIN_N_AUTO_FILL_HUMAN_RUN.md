# N列ASIN自動（⑥）— 人間手順

**正本**: [B_ASIN_N_AUTO_FILL_REQUIREMENTS.md](B_ASIN_N_AUTO_FILL_REQUIREMENTS.md)

---

## 事前

1. `clasp push`（本機能を含むローカル差分）
2. シート再読み込み → Z→15 に **15-⑮**、B に Step **6.5** があること
3. 対象親に出品CK。`メーカー名`＝`ブランド`（または▼マスタ(ブランド)）。`競合店ASIN` があり、同一JANで Keepa◎（ログまたは貼付）

---

## 単体試験（推奨）

1. 試験親1件だけレ点。N列は空
2. **Z → 15-⑮**
3. 期待: N列に競合ASIN・黄背景・`要確認_ASINN`＝要確認
4. ログ: `[N列ASIN自動][ASINN_…]` で `filled=` / `skip=`

スキップ例: `N_already_filled` / `brand_ne_maker` / `no_competitor_asin` / `not_circled`

---

## B統合

- Step6 のあと **6.5** が自動実行（`B_ASIN_N_AUTO_FILL_ENABLED` 未設定＝ON）
- 緊急停止: Property を `false`

---

## 確認後

- 相乗り出品前に黄セルを目視 → OKなら `要確認_ASINN` を確認OK（任意）／背景クリア可
- 誤記入なら N列を空に戻す
