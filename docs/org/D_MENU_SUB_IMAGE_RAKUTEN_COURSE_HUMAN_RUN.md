# サブ画像コース（楽天出口）— 人間手順

**状態**: **実装済（要 clasp push → 実機検収）**  
**承認**: [LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md](LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md)  
**要件**: [D_MENU_SUB_IMAGE_RAKUTEN_COURSE_REQUIREMENTS.md](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_REQUIREMENTS.md)  
**投入**: [D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)  
**PoC詳細**: [D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md](D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)  
**MAIN寄せ**: [B45 §4.5.7](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)（P1：閉じても前回テーマ案を再利用。直すときはレ点外し＋要望→再作成。要 clasp）

---

## 0. 人がやること（これだけ）

1. 各JANの **PACKAGE_TRUTH**（単体／N=1）を `01.amazon白抜きベース` へ保存（**ファイル名にJAN/SKU不要**）
2. B-③で **サブ採用CK**
3. **B-⑤**（関数 `menuB4*`）→ **JAN↔正本画像の紐付け**＋**ベース色3択** → **テーマ概要を複数チェック** → Agent に貼る
4. **【必須】見本ゲート**（指示文先頭）: `sub_image_food_canon_launch.py` で主／副（または正本のみ）確定 → その後テーマ別新規／compose（`--theme-ids-for`）。**既存画像だけでやり直し**するときは正本紐付け画面で「画像の生成不要」チェック（§7.13）→ テーマ案Geminiなしで指示へ → `b5_skip_generate_finish.py`
5. **確認は確認ページ**。作成が終わっても、確認フォルダへは出さない。
6. 残す絵だけ「採用する」にレ点を付け、**採用分を確認フォルダへ**を押す。その商品の古いサブは「外した画像」へ移り、採用した枚だけ直下に残る。メイン画像と「アップロード済み画像」は動かさない。
7. NGなら確認ページで、直す絵だけ再生成する。採用を押すまで確認フォルダは変わらない。
8. OKなら必要時 `--to-checked-children` → C「楽天のみ」→アップ。Amazonは U4

**やらない**: phaseOrder手編集、Amazon PT新規、Yahoo、Vision QA、テスト出力フォルダの巡回。機械自動 CANON だけで本生成に進むこと（見本ゲート未確定時）。**生成不要モードで初回未生成JANを完走扱いすること**。

### B-⑤ 正本紐付け（関数 `menuB4*`）

- Drive の `01.amazon白抜きベース`（＋`処理済み`）画像一覧を表示し、B-③ JAN とドロップダウンで対応付け
- 未選択 JAN は合成スキップ
- フォルダが見つからない場合: Script Properties `SET_MAIN_AMAZON_BASE_FOLDER_ID` にフォルダIDを設定（または手入力 `JAN=相対パス`）
- 指示文に `--package-truth "JAN=G:\\…\\ファイル"` が付く（ファイル名にJAN不要）
- **ベース色**: ドロップダウン3択 → 指示に `--base-color beige|warm_white|soft_gray`（既定ベージュ）。背景・カードのみ。正本: `tools/set_main_image/tonmana_palette.py`（GAS `b4TonmanaPalette_` と同期）
- **テーマ概要**: 紐付け確定後、出品材料で作れるテーマだけ出す（**1悩み・2ベネフィットは必須**。ほか最大8＝合計10）。Gemini が2文概要（80〜140字）。人がチェック＋文面編集 → `--theme-ids-for` と概要が指示に付く。切戻し `B5_PROPOSE_GEMINI_ENABLED=false`

### 写真実写ルール

- 正本: `tools/set_main_image/photo_realism_rules.py`（VERSION 付き。改善時はこのファイルを更新）
  - **2026-08-10.2 CAMERA_LOOK 案A**: Canon EOS R5 + 50mm f/1.8／背景・小物は creamy bokeh／切り抜き輪郭禁止／粒感
- DOF・距離ボケ・反射・陰影。複数被写体はピンボケ必須

---

## 1. フロー

```text
【人】PACKAGE_TRUTH保存
 → B-③ → サブ採用CK
 → B-⑤ → 正本紐付け → テーマ概要チェック → Cursor指示
 →【必須】見本ゲート（sub_image_food_canon_launch）で FOOD_UNIT_CANON 人間確定
 → テーマ別新規 dry_sub／または競合 compose
      → 絵は作業フォルダと確認ページ。確認フォルダへは出さない
 →【人】確認ページで残す絵だけ採用
      → その商品の古いサブを「外した画像」へ移し、採用した枚だけ確認フォルダ直下へ
 →（NG）確認ページで直す絵だけ再生成 → もう一度採用
 →（OK）必要時 --to-checked-children → C「楽天のみ」→アップ → Amazonは U4
```

### 画像の生成不要（§7.13）

```text
cd tools\set_main_image
python b5_skip_generate_finish.py --jan 4906283045119 --jan 4906283045317 --path both --export --launch-review
# テーマ別新規のみ: --path theme_new
# 競合のみ export: python sub_image_ai_compose_poc.py --export-only --jan …
# dry_sub 既存のみ: python dry_sub_overview_copy_test.py --jan … --out-dir _out/theme_new_run_… --skip-generate --auto-export
```

### 人間が見る場所

| 見る | パス |
|------|------|
| **絵の確認** | 確認ページ（残す絵にレ点） |
| **紐付けが読む場所** | `…\02.楽天アップロード画像保存場所` の直下。採用した枚だけ |
| 見ない（作業用） | `…\05.画像生成（セットMAIN）\00.テスト出力\sub_image_ai_compose\…` |
| 正本素材のみ | `…\05.画像生成（セットMAIN）\01.amazon白抜きベース` |

### B-⑤（メニュー・関数 `menuB4*`）

- **JAN↔正本画像の紐付け**（＋ベース色）→ **テーマ概要案**（競合＋穴。MAIN採用は判断材料・完全一致不要）を複数チェック → 同一ダイアログに Cursor 指示（`--package-truth`・`--base-color`・`--theme-ids-for`）。**指示先頭＝見本ゲート必須**（§7.12）
- 続けて **「compose後の目視チェック（必須）」** パネル: チェックリスト＋review_loop 用 Agent 指示＋再生成コメント・テンプレ
- ※ `google.script.run` 成功後に別 `showModalDialog` を開くと「生成中…」で止まるため、指示はパネル切替で返す（要 `clasp push`）
- ※ メニュー実行はダイアログ表示で **完了**（7〜9秒程度が正常）。「終わらない」＝画面上の紐付け待ち／テーマ案の `buildB4ThemeOverviewsForBindings`／指示の `finishB4AfterTruthBind`
- ※ `google.script.run` 対象関数に末尾 `_` を付けない（private 扱いで「送信中」のまま固まる）
- 再生成本体はローカル `sub_image_review_loop.py`（ブラウザUI＋テンプレボタン）。GASは合成しない

### 再生成コメント・テンプレ

- 正本: `tools/set_main_image/review_feedback_templates.py`（GAS側 `getB4ReviewFeedbackTemplates_` と同期）
- 例: 切り抜き輪郭／背景シャープすぎ／CGI照明／プラスチック質感／湯気線画／PACKAGE_LOCK／文字過多

### PACKAGE_TRUTH

| 項目 | 内容 |
|------|------|
| 何を | 単体／N=1相当の商品パッケージ写真 |
| 置き場 | `01.amazon白抜きベース`（ファイル名にJAN/SKU**不要**）。B-④で人間がJANと紐付け |
| Property | 任意 `SET_MAIN_AMAZON_BASE_FOLDER_ID`（Drive フォルダID） |
| 禁止 | N≥2セットMAINを正本に／正本OCR |
| 欠落 | 紐付けない／未選択JANは合成スキップ |

### 確認フォルダへ出すとき

- 作成直後は出さない。確認ページの「採用分を確認フォルダへ」だけが、直下へ出す。
- 出す前に、同じ品番のサブ画像（`__skip` 付きも含む）を直下から「外した画像」へ移す。メイン（`*_rakuten.jpg`）と「アップロード済み画像」は動かさない。
- 採用した枚は最大10枚。普通の名前で置く。`__skip` は付けない。
- 紐付けは確認フォルダの直下を読む。だから、直下に残った絵だけがシートに出る。
- 全セット子への複製はしない。C本線（親サブ→全子同一）では **親あたり1セット**で足りる。`--to-checked-children` は明示の本番複製だけ（枠外候補の元になるので常用しない）。

---

## 2. コマンド例

```text
cd tools\set_main_image

python sub_image_b3_curate.py --jan 4538872281013 --classify

python sub_image_ai_compose_poc.py --providers openai --openai-quality medium --allow-invented false --export-pick a --auto-export --jan 4538872281013

# 確認ページで残す絵だけ採用する。作成直後は確認フォルダに出ない
# レビューUI（推奨・出品CKレ点親を一括・ポート連番で複数タブ）:
python sub_image_review_launch.py --from-sheets
# または JAN 明示:
python sub_image_review_launch.py --jan 4906283045119 --jan 4906283045317 --jan 4906283047410
# 単親のみ:
python sub_image_review_loop.py --compose-dir "…\00.テスト出力\sub_image_ai_compose\4538872281013_YYYYMMDD_HHMMSS"
```
`sub_image_review_launch.py` は各JANの**最新** compose-dir を選び、8765,8766,… でタブを開く。Ctrl+C で全停止（`--no-wait` なら起動後すぐ終了）。「修正なしで終了」後も他タブは続行可。採用を押したあと、確認フォルダの直下だけが紐付けの対象です。

セット数から子SKU解決:

```text
python export_sub_images_for_rakuten_matrix.py --from-compose-dir … --pick a --parent-sku "親SKU" --set-count 3 --master-csv master.csv --out-dir …
```

MAIN も同ソースへ `{子SKU}_rakuten.jpg` を置く場合は [セットMAIN HUMAN_RUN](D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)。

---

## 3. GAS（人間）

1. `clasp push`（B-③列拡張・楽天ファイル名自動を反映）
2. C → **楽天のみ**（またはマトリクス生成）
3. ログ `rakutenFilenameAutobind` の sub 件数を確認
4. 目視 → リネーム＆アップロード → マスタ「楽天サブ画像n」URL
5. Amazon: 楽天サブが入っていれば **U4／Ama新カタログ②**（新規白抜きサブは不要）

戻し: `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false`

### 実機ステータス（2026-08-09）

| 項目 | 状態 |
|------|------|
| コード／docs | **済** |
| clasp push | **人間待ち** |
| 缶飯系 E2E | **人間待ち**（§4） |

---

## 4. 検収チェック

- [ ] B-③に `参照themeId`／`参照phaseOrder`／`参照cluster` が自動で入る（人が編集しなくてもよい）
- [ ] PACKAGE_TRUTH無しJANがスキップされ、有りJANは色／ラベル／縦横比が正本寄り
- [ ] レビューUIでチェック＋要望→再生成ができ、「修正なしで終了」できる（`review_launch --from-sheets` で複数タブ可）
- [ ] 人手は採用CK＋レビュー／A/B目視のみで手順が閉じる
- [ ] `{子SKU}_subN.jpg` が正しい子行の楽天サブNへ入り、アップ後マスタURLが埋まる
- [ ] Property false でファイル名自動が止まる
- [ ] Amazon 新規サブ作業を要求していない（U4 REUSE）
- [ ] 楽天CSV／Amazon MAIN ε 非回帰

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-10-19 | 採用したサブは **子SKU_種類_sub番号** で確認フォルダへ出す。紐付けコードは残っている。読み取り上限の既定は200（`RAKUTEN_IMAGE_MATRIX_FILE_LIMIT`）。要 clasp。 |
| 2026-10-19 | 確認ページに **個品が見本と違う** ボタン。見本の色・厚み・1片のまま描き直す指示が入る。画面の反映は確認ページの再起動。スプシ側は clasp 後。 |
| 2026-10-19 | **採用した絵だけ確認フォルダ直下**。作成直後は出さない。同じ品番の古いサブは「外した画像」へ移す。戻しは Git。clasp 不要。 |
| 2026-08-30 | **B-⑤指示先頭＝見本ゲート必須**（§7.12）。要 clasp。[ALIGN §7.12.3](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-08-30 | **競合本線**: チェック∩写真・1テーマ1枚・想像OFF・`--export-pick a`。[ALIGN §7.9](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。要 clasp。 |
| 2026-08-30 | **P-align-2b 本線**: 材料門＋Gemini 2文。1・2必須＋ほか最大8＝10。要 clasp push。 |
| 2026-08-30 | **P-align-2b**: テーマ概要チェック→`--theme-ids-for`。表示B-⑤。要 clasp push。 |
| 2026-08-10 | **ベース色 PoC 3択**（B-④＋`--base-color`／`tonmana_palette.py`）。要 clasp push。 |
| 2026-08-10 | B-④紐付け後の指示を**同一ダイアログ内表示**（`showModalDialog` 二重呼び出しハング修正）。要 clasp push。 |
| 2026-08-10 | B-④に **目視チェック必須パネル**（review_loop指示＋コメントテンプレ）。`review_feedback_templates.py`。要 clasp push。 |
| 2026-08-10 | 目視 auto-export は品番キー×最大10枚（pick=ab）。全セット子FO禁止。`--to-checked-children` は出品CK子のみ。CAMERA_LOOK 案A。 |
| 2026-08-09 | **1話完結**: compose後 auto-export で楽天フォルダへ直出し。人間目視は `02.楽天アップロード画像保存場所` のみ。 |
| 2026-08-09 | **PACKAGE_TRUTH／LOCK強化／文字量上限／SEO／人間レビュー再生成ループ**。B-④正本確認。Vision QA・ベース色メニューは対象外。 |
| 2026-08-09 | **B-④**: サブ採用CKの JAN から Python／Cursor 指示ダイアログ（GAS非合成）。 |
| 2026-08-09 | 初版。楽天出口本線。 |
