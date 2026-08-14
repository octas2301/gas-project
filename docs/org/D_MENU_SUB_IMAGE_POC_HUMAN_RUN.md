# Amazon サブ画像 — B-③右列レ点＋採用ログ

**状態**: PoC（マスタ／R2未書込）  
**本線（楽天出口）**: [D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)  
**作業シート**: `競合画像取得（必要時B-③実行）`（A–Hは候補、右列がサブ用）  
**採用ログ**: `サブ画像採用ログ`（**レ点した行だけ**）  
**旧シート** `サブ画像競合候補（人間確認）` は廃止（curate実行時に削除）

## 列（B-③）

| 列 | 内容 |
|----|------|
| A–H | 従来どおり（マスタ行〜プレビュー） |
| 商品名 | JANから解決 |
| 自動判定 | keep / drop_main / drop_unrelated / review |
| 参照themeId／参照phaseOrder／参照cluster | **自動**（LPページ分類。人手編集しない） |
| サブ採用CK | **レ点＝採用**（チェックボックス） |
| メモ | 人手 |

## 運用

```text
cd tools/set_main_image
# 右列を埋める＋採用ログ同期＋旧シート削除
python sub_image_b3_curate.py --jan 4538872180149 --jan 4538872281013 --jan 4538872285127 --classify

# 人が B-③ でサブ採用CKをレ点編集したあと
# → スプレッドシート「B-④ サブ画像作成（Python／Cursor指示）」でコマンドをコピーして Agent へ
# または手動:
python sub_image_b3_curate.py --sync-adopt-log-only
python sub_image_ai_compose_poc.py --providers openai --openai-quality medium --target-slots 5 --jan …

# 本線手順の正本: D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md

# 合成（レ点のみ・JAN分離・各10パターン）— Pillow 切り出し版（精度不足時は下のAI版へ）
python sub_image_part_poc.py --from-b3-adopt --jan 4538872180149 --jan 4538872281013 --jan 4538872285127 --max-images 12 --auto-accept

# AI合成 v3（本線）: 心理プロセス順スロット5×A/B / gpt-image-2 medium / パッケージ厳禁
# 分類はページ全体のみ。不足フェーズは競合phase被り→想像(最大2)
python sub_image_ai_compose_poc.py --providers openai --openai-quality medium --target-slots 5 --jan 4538872281013

# 旧: Gemini + OpenAI 両比較
# python sub_image_ai_compose_poc.py --providers gemini,openai --route all --jan ...

# 段階1: Gemini本線 + falはシズル(T03)のみ・文字禁止
# キー: secrets/fal_api_key.txt または FAL_KEY / Geminiも従来どおり
pip install fal-client
python sub_image_ai_compose_poc.py --providers gemini,fal --route auto --jan 4538872180149
# falスモークのみ
python sub_image_fal_poc.py --txt-only
# fal競合改変4本比較（背景/リライト/湯気/文字色・商品色ロック）
python sub_image_fal_edit_compare_poc.py --jan 4538872180149

# 旧: 全テーマを全プロバイダで回す
# python sub_image_ai_compose_poc.py --providers gemini,fal --route all --jan ...
```

### Pillow 切り出し版
出力（商品ごと）: `00.テスト出力/sub_image_parts_poc/<JAN>_<runId>/`
- `04_compose/P01_…P10_….jpg` … 10パターン提案
- `04_contact_10patterns.jpg` … 一覧
- `03_parts_propose/*_part*_clean.png` … 縁背景除去後のパーツ

パーツ合成時は縁の元背景を透明化してから貼る（文字は改変しない）。

### AI合成版 v3（gpt-image-2 medium・スロット駆動）
方針:
- 採用競合をLPテーマ①〜⑳に**ページ全体**で分類（`primary`＋任意`secondary`、最大2）
- **phaseOrder** はマスタ付与。スロットを心理プロセス順（フック→…→クロージング）で最大5（`--target-slots 6`可）
- 各スロット **A/Bの2案**（計10枚／12枚）。ファイル名例: `S01_T03_shizuru_sizzle_ABa.jpg`
- 不足時: 競合のphase被り追加 → それでも不足なら**想像スロット最大2**（型のみ・事実数字禁止）
- 連続類似抑制: `contentCluster` 隣接禁止
- トンマナ: **ベージュ固定**／**商品パッケージ改変厳禁**
- 文字量抑制（成分表例外）／AI改変最大約50%
- 本線モデル: **`gpt-image-2` `quality=medium`**

#### 将来拡張メモ（未実装）
- **パーツ単位テーマ分類**は見送り（コスト・工数）。必要になったら検出→切出し→パーツ分類→スロット集約を Phase2 で追加する。
- 現状はページ全体分類（方式A）のみ。

出力（商品ごと）: `00.テスト出力/sub_image_ai_compose/<JAN>_<runId>/`
- `01_refs/` … テーマ分類に使った採用画像
- `03_openai/Sxx_Tyy_<slug>_ABa.jpg` … OpenAI（本線）
- `02_gemini/` … `--providers` に gemini を含めたとき
- `04_fal/` … fal シズル（hybrid時）
- `04_contact_themes_both.jpg` … 一覧
- `THEMES.md` … スロット＋source（competitor / competitor_dup_phase / invented）
- `05_instruction_report/` … 日本語指示ボード
- `_meta/run_meta.json` … `version: 3`

OpenAIクレジット枯渇時は `--providers gemini` で代替可。

## B-③再実行時（GAS）

1. 破棄前のB-③レ点 ＋ `サブ画像採用ログ` をマージ  
2. 最新A–Hを書き直し、一致URLにレ点復元  
3. 採用ログを **新一覧上でレ点が付いている行だけ** に書き換え  

要: ローカルで `clasp push`（`コード.js`）。
