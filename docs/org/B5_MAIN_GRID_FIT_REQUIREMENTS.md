# B-⑤ MAIN — 規定マス内収容（おしゃれ優先）要件

**文書種別**: 要件定義  
**最終更新**: 2026-09-03  
**状態**: **P1 実装済**（段階閾値＋ソフトプロンプト）  
**整合**: [B45_MAIN_SUB_ALIGN_REQUIREMENTS.md](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) §7.6ter／[D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md](D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md) §3.2

---

## 0. 方針（人間ロック 2026-09-01）

| 項目 | 正 |
|------|-----|
| **最優先** | おしゃれさ・目を引く力（IMAGE_STYLE_TYPO 雰囲気） |
| **文字載せ** | **AI が描く**（OpenAI フリーレイアウト）。**Pillow 後載せ文字は NG** |
| **20マス** | **努力目標・完全許容**（≤20 は retry しない） |
| **21–24** | **バッファ**（20 に収まらなかったときの許容帯。**retry しない**） |
| **≥25** | **再生成トリガー**（最大 `P3_FREE_MAX_TRIES`） |
| **事後判定** | compose try ごと **Gemini 3 Flash** gate（≈0.18円/枚）。失敗時 catch_v7 |
| **レビュー UI** | catch_v7 100マスガイド（無料・即時）。最終採用は人間 |

副作用（デザイン性の著しい低下）が出た場合は **環境変数で切戻し**（§6）。

---

## 1. 格子・定数

| 定数 | 値 | 意味 |
|------|-----|------|
| `P3_CATCH_TARGET_CELLS` | 16 | 理想（金丸4マス予約込みの運用目標） |
| `P3_TEXT_MAX_CELLS` | 20 | 楽天20%・**完全許容上限** |
| `P3_TEXT_BUFFER_MAX_CELLS` | 24 | バッファ上限 |
| `P3_TEXT_RETRY_MAX_CELLS` | 25 | **この値以上で retry**（`hitN >= 25`） |
| ROI | 左上 4×7 | 計測・期待帯域 |

---

## 2. tier 分類

| tier | hitN | compose retry | rank ペナルティ |
|------|------|---------------|----------------|
| `ok` | ≤20 | しない | なし |
| `buffer` | 21–24 | **しない** | なし |
| `retry` | ≥25 | する | +500 |

実装: `tools/set_main_image/b5_grid_budget.py` → `classify_grid_tier` / `grid_needs_retry`

---

## 3. 生成前 — ソフトプロンプト（Pillow なし）

文案ごとに `_p3_text_zones` から列・行を算出し、`_prompt_p3_free` に **【10×10マス・努力目標】** ブロックを注入。

- 目指す: **20マス以内**（理想16）
- 21–24 はバッファ、25以上は避ける
- **おしゃれさ最優先** を明記

### 3.1 キャッチ高さ・書字（人間合格 2026-09-03）

| 項目 | 正 |
|------|-----|
| キャッチ高さ | **上限約155px・下限なし**（`B5_CATCH_H_MAX`／`B5_CATCH_H_MIN`。空＝この既定） |
| 書字方向 | **縦書き＝おしゃれ、ではない**。縦は**和風トーン商品のみ**少数採用（`b5_catch_zone_adopted.plan_writing_modes`） |
| 検収サンプル | かつお／まぐろ／ぶり MAX155 全件（`*_20260903_155154`／`*_161505`／`*_163817`） |

オフ: `B5_GRID_SOFT_BUDGET=0`

---

## 4. compose ループ

1. OpenAI 画像生成（scene＋後載せ文字）
2. `measure_grid_dual`（local + Gemini）
3. `gateHitN` で tier 判定
4. **tier ≠ retry かつ 市松なし かつ キャッチあり** → try ループ終了
5. try は `_tries/{theme}/` に原画＋オーバレイ保存（上書きしない）
6. 完了後 `_tries/index.html` 自動生成

---

## 5. retry 文

| 条件 | 内容 |
|------|------|
| キャッチ無し | 既存強め指示 |
| hitN ≥25 | おしゃれさ維持＋装飾コンパクト化 |
| hitN 21–24 で retry する場合のみ（通常は break するため稀） | ソフトに20以内を促す |

---

## 6. 切戻し（副作用時）

| 環境変数 | 効果 |
|----------|------|
| `B5_GRID_TIER_POLICY=strict20` | **旧挙動**: 21超で retry 扱い（2026-09-01 前の厳しめ） |
| `B5_GRID_RETRY_AT=26` | retry 開始を26以上に変更 |
| `B5_GRID_SOFT_BUDGET=0` | ソフトプロンプトブロックを出さない |
| `B5_GRID_GEMINI=0` | Gemini gate オフ（catch_v7 のみ） |
| `B5_CATCH_H_MAX=165` 等 | キャッチ高さ上限の試験値（空＝**155**） |
| `B5_CATCH_H_MIN=171` 等 | 下限を戻す（空／0＝**下限なし**） |

Git revert 対象ファイル: `b5_grid_budget.py`, `b5_main_base_compose.py`, `b5_virt_abc.py`（`grid_budget_block`）

---

## 7. 実装ファイル

| ファイル | 役割 |
|----------|------|
| `b5_grid_budget.py` | tier・budget・ソフトプロンプト・retry 文 |
| `b5_grid_gemini_measure.py` | Gemini 計測・try HTML |
| `b5_main_base_compose.py` | compose 統合 |
| `b5_virt_abc.py` | `_prompt_p3_free(grid_budget_block=…)` |

---

## 8. 受入（P1）

- [ ] compose で hitN=22 の try が **retry されない**
- [ ] hitN=25 以上で retry が走る
- [ ] `_tries/index.html` に tier 表示
- [ ] 目視10枚で「旧より地味化」が無い（副作用チェック）

---

## 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-09-03 | **キャッチ高さ MAX155・下限なし合格**（かつお／まぐろ／ぶり）。縦書き≠おしゃれ・和風のみ少数。既定 `B5_CATCH_H_MAX=155` |
| 2026-09-01 | 初版。Pillow NG・tiered 閾値・ソフトプロンプト・Gemini gate 連携 |
