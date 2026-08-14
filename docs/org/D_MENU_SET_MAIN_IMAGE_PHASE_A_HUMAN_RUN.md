# セットMAIN画像 Phase A — 人間／Agent 手順

**承認**: [LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md](LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)  
**ツール**: `tools/set_main_image/`（Pillow・生成AIなし・出力JPEG quality≈85）  
**作業ルート**: `G:\マイドライブ\05.画像生成（セットMAIN）\`  
**量産出口（楽天）**: `G:\マイドライブ\03.楽天・Yahoo!商品登録（CSV一括UL）\02.楽天アップロード画像保存場所`  
**Cメニュー**: 「セットMAIN量産（楽天・Cursor指示）」→ 文字色（日本語）選択 → Cursor指示文をコピー

---

## 0. フォルダ構成（確定）

```text
05.画像生成（セットMAIN）/
  01.amazon白抜きベース/     … Amazon用1個セット
  02.楽天ベース/             … 楽天用1個セット（金丸込み・数字なし推奨）
  94.文字色見本/             … 01_えんじ.png … 05_茶.png（色正本の見本）
  95.CANVA数字・セット/      … digit / pair / unitset（Canva・上書き禁止）
  97.楽天数字レイヤ/         … 任意。数字「1」だけ不透明＝位置ガイド
  98.楽天金丸素材/           … 任意。02に金丸が無いとき用（現状ツールは未自動貼付）
  99.octas期限管理シール素材/ … 食品共通1枚
```

ファイル名に **親SKU** を含めると自動選択しやすい。無い場合はフォルダ内の最新画像を使う。

---

## 1. 文字色（Cメニュー／CLI）

| # | 日本語 | hex |
|---|--------|-----|
| 1 | えんじ（既定） | `#800020` |
| 2 | 青 | `#001b44` |
| 3 | 黒 | `#0d1321` |
| 4 | 緑 | `#1b4332` |
| 5 | 茶 | `#3d2b1f` |

`--text-color 1` または `--text-color 茶`。見本は正方形パッチ（数字ゼロ不要）。

---

## 2. 「数字レイヤ」とは何か（重要）

**書体見本シートではありません。**

| 方式 | 何を置くか | いつ使う |
|------|------------|----------|
| **本線（推奨）** | Canva 透過グリフ＋実行時再着色 | 95に素材があるとき |
| **補助** | 97に「数字1だけ」の透過PNG | 位置・色を画像から取りたいとき |

### 書体カタログ（フォールバック `--font-id`）

```text
cd tools\set_main_image
python compose_set_main.py --list-fonts
```

※Canva有料フォントそのものは配布できないため、**Windows標準フォントで近似**します（自動生成フォールバック用）。

---

## 3. 98 金丸素材は必要か？

| ケース | 必要？ |
|--------|--------|
| 02の楽天ベースに **すでに金丸が入っている** | **不要** |
| ベースに金丸が無く、後から貼りたい | 98に金丸PNGを保管（自動合成は後続） |

---

## 4. マスタ読取（推奨: スプシ直読）

**出品CKの正**は Google スプレッドシート。CSVダウンロードは不要（古いCSVだとレ点がずれる）。

| 方式 | コマンド | 備考 |
|------|----------|------|
| **スプシ直読（推奨）** | `--from-sheets` | C1 と同じ OAuth（`tools/c1_hpc_packaged/config.local.json` + `secrets/`） |
| 互換CSV | `--master-csv …` | 手元エクスポート。レ点はエクスポート時点 |

```text
cd tools\set_main_image
python sheets_master.py
# → rowCount / trueCk（出品CKが TRUE の件数）を確認
```

### Amazon素材・透過2点確認（必須）

Cメニュー「セットMAIN量産（Amazon・透過確認＋Cursor指示）」でも同内容を出します。

1. **① 透過PNGか** — `01.amazon白抜きベース` の素材にアルファがあるか  
2. **② Canva背景リムーバー実施か** — ダウンロードの「背景透過」だけでは不足。写真の白を消して市松になったPNGか  

```text
cd tools\set_main_image
python check_amazon_base_matte.py --require-ok
```

両方 OK（`matteOk`）になるまで量産しない。

### 親SKU自動紐付け（複数商品まとめ量産）

`01` 直下の透過PNGを、出品CKレ点親の **楽天メイン画像1** と Vision 照合し、`{親SKU}_単体.png` へ自動リネームする（人手確認なし。閾値未満はSKIP）。

```text
cd tools\set_main_image
python bind_amazon_base_to_parents.py --from-sheets
# 確認のみ: --dry-run
# 未割当が残ったら失敗終了: --require-bound
```

メタ: `05…/00.テスト出力/_meta/BASE_BIND_latest.json`

---

## 5. 実行（Agent）— Amazon Pillow 貼付（セットMAIN）

Cメニューでパターン選択: **①正方形（本線）** / **②縦長タイプ（合格固定）** / **③横長タイプ（合格固定）**

```text
cd tools\set_main_image
python check_amazon_base_matte.py --require-ok
python bind_amazon_base_to_parents.py --from-sheets
python amazon_paste_batch.py --from-sheets --checked-only --min-n 1 --loose-unit-match --aspect square
# 縦長タイプ（瓶など・合格固定）:
python amazon_paste_batch.py --from-sheets --checked-only --min-n 1 --loose-unit-match --aspect portrait
# 横長タイプ（缶など・合格固定）:
python amazon_paste_batch.py --from-sheets --checked-only --min-n 1 --loose-unit-match --aspect landscape
# バインドを量産コマンドに内包する場合:
python amazon_paste_batch.py --from-sheets --checked-only --min-n 1 --loose-unit-match --aspect portrait --auto-bind-bases
```

- 出力先（本線）: `G:\マイドライブ\04.amazonカタログ作成（CSV一括UL）\07.白抜きの置き場（人間が入れる）\{子SKU}_amazon.jpg`
- **N=1（1個セット）も対象**（中央最大化・単体配置）。`--min-n` 既定は 1
- **縦長タイプ（portrait・合格固定）**: キャンバスは正方形のまま。**N=2＝の字＋直立四隅**／**N=3＝幅広 upright扇／細長 top_arc＋Q(四隅上辺中点)→M＋緑法線上+35px**／**N=4＝直立四隅＋扇（unit3 BR枠下接触）**／**N≥5＝hero直立＋右傾け積み**。素材は `01.amazon白抜きベース` 透過PNG。見本=`05…/03.amazon見本/縦型レイアウト基本`。詳細=[SET_MAIN_LAYOUT_RULES.md §1.1](SET_MAIN_LAYOUT_RULES.md)。**Cで量産→07→D本線で消費**（Dは07画像を読むだけ・追加ロジック不要）。
- **横長タイプ（landscape・合格固定）**: **N=1中央**／**N=2縦二段**／**N=3–4階段枠ピン**／**N≥5上hero+下グリッド（右Octas帯・半端行中央）**。詳細=[SET_MAIN_LAYOUT_RULES.md §1.2](SET_MAIN_LAYOUT_RULES.md)。C③→07→D本線。
- メタJSON/サマリ: 同フォルダの `_meta\`（アップロード対象外）
- 直下を画像のみに整理: `python tidy_images_only_folder.py --also-test-out`
- 検証だけテスト出力へ: `--test-out --name-style debug`
- `--loose-unit-match` … `01` に単体が **1枚だけ** のとき、ファイル名に親SKUが無くても許可

---

## 6. 実行（Agent）— 楽天 Canva本線

マスタ **出品CKレ点** × **総個数** × **バリエーション単位**（空/素材なし→**個**）で、`02.楽天ベース` にバッジを合成。

```text
cd tools\set_main_image
python compose_set_main.py ^
  --master-csv "…\master_export.csv" ^
  --checked-only ^
  --malls rakuten ^
  --rakuten-engine layered ^
  --text-color 1 ^
  --out-dir "G:\マイドライブ\03.楽天・Yahoo!商品登録（CSV一括UL）\02.楽天アップロード画像保存場所"
```

親SKUで絞る場合は `--parent-sku sanky-…-oya` を追加。  
Cメニューから実行すると、上記相当の指示文がダイアログに出ます（GASは合成しない）。

旧・数字のみ描画: `--rakuten-engine badge`  
Amazon も同時: `--malls amazon,rakuten`  
`--work-root` 省略時は `05.画像生成（セットMAIN）`。

出力: `{子SKU}_rakuten.jpg`（JPEG）。**N=1（1個セット）もレ点があれば生成**（金丸に「1」＋単位）。  
セット数 N → 子SKUはマスタ行の総個数で解決済み（`master_sets.resolve_child_by_set_count`）。マトリクス生成時はファイル名の子SKUで楽天メイン1へ自動投入（[SKU紐付け](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)）。
量産成功後、使用した `02.楽天ベース` 直下のベース画像は `02.楽天ベース/処理済み/` へ自動移動します（スモークは対象外）。

### スモーク

```text
python compose_set_main.py --smoke-amazon --out-dir .\_smoke_out
python compose_set_main.py --smoke-rakuten --text-color 茶 --out-dir .\_smoke_out
```

---

## 7. 検収

- [ ] 05の01/02/94/95に素材がある  
- [ ] 出力が楽天アップロード画像フォルダにある（子SKU付き jpg）  
- [ ] 楽天: 金丸はそのまま・文字色が選択どおり  
- [ ] 食品: Octasが右下（Amazon時）  
- [ ] Amazon一括は `--from-sheets` で現行レ点が読める  

---

## 8. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-09 | **楽天SKU自動紐付け**: 出力名とマトリクス自動投入の接続を追記。[SKU HUMAN_RUN](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)。 |
| 2026-08-09 | **横長タイプ合格固定**: C③／`--aspect landscape`／`landscape_layout.py`。楽天N=1をC指示に明記。要 clasp push。 |
| 2026-08-09 | **楽天 N=1 も生成**（レ点時）。`compose_rakuten_badge` の N≥2 制限を撤廃。 |
| 2026-08-08 | **01白抜き自動紐付け**: 楽天メインVision照合→`{親SKU}_単体.png`（`bind_amazon_base_to_parents`／`--auto-bind-bases`）。 |
| 2026-08-08 | **N=3細長合格確定**: Q＝四隅上辺中点→ARC-MID→緑法線上+35px（缶・唐辛子）。`unit0NudgeUpAlongNormalPx=35`。C②→07→D。 |
| 2026-08-08 | **縦長タイプ合格固定**: N=2/3/4/≥5 `locked_pass`。Cメニュー②「縦長タイプ（合格固定）」。unit3 BR=枠下接触。出力07→D本線。 |
| 2026-08-05 | **縦長パターン**: `--aspect portrait` 実装。斜めファン・N≤4。Cメニュー②有効（要 clasp push）。 |
| 2026-08-04 | Amazon貼付バッチのスプシ直読（`--from-sheets` / `sheets_master.py`）。 |
| 2026-08-02 | 初版。 |
| 2026-08-02 | 05フォルダ構成・書体カタログ・数字レイヤ説明・JPEG q85。 |
| 2026-08-02 | 文字色5種・出力先を楽天ULフォルダ・Cメニュー（Cursor指示＋日本語色名）。 |
