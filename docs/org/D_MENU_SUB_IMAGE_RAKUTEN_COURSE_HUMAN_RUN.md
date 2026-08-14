# サブ画像コース（楽天出口）— 人間手順

**状態**: **実装済（要 clasp push → 実機検収）**  
**承認**: [LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md](LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md)  
**要件**: [D_MENU_SUB_IMAGE_RAKUTEN_COURSE_REQUIREMENTS.md](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_REQUIREMENTS.md)  
**投入**: [D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)  
**PoC詳細**: [D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md](D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)

---

## 0. 人がやること（これだけ）

1. 各JANの **PACKAGE_TRUTH**（単体／N=1）を `01.amazon白抜きベース` へ保存（**ファイル名にJAN/SKU不要**）
2. B-③で **サブ採用CK**
3. **B-④** → **JAN↔正本画像の紐付け**＋**ベース色3択**（ベージュ／ウォームホワイト／ソフトグレー）→ Agent に貼る（compose が自動 export まで）
4. **目視フォルダは1つだけ**:  
   `G:\マイドライブ\03.楽天・Yahoo!商品登録（CSV一括UL）\02.楽天アップロード画像保存場所`  
   （`{品番キー}_{themeSlug}_subN.jpg`・最大10枚／全子複製しない）
5. NGなら `sub_image_review_loop.py` → 再目視
6. OKなら必要時 `--to-checked-children` → C「楽天のみ」→アップ。Amazonは U4

**やらない**: phaseOrder手編集、Amazon PT新規、Yahoo、Vision QA、テスト出力フォルダの巡回。

### B-④ 正本紐付け

- Drive の `01.amazon白抜きベース`（＋`処理済み`）画像一覧を表示し、B-③ JAN とドロップダウンで対応付け
- 未選択 JAN は合成スキップ
- フォルダが見つからない場合: Script Properties `SET_MAIN_AMAZON_BASE_FOLDER_ID` にフォルダIDを設定（または手入力 `JAN=相対パス`）
- 指示文に `--package-truth "JAN=G:\\…\\ファイル"` が付く（ファイル名にJAN不要）
- **ベース色**: ドロップダウン3択 → 指示に `--base-color beige|warm_white|soft_gray`（既定ベージュ）。背景・カードのみ。正本: `tools/set_main_image/tonmana_palette.py`（GAS `b4TonmanaPalette_` と同期）

### 写真実写ルール

- 正本: `tools/set_main_image/photo_realism_rules.py`（VERSION 付き。改善時はこのファイルを更新）
  - **2026-08-10.2 CAMERA_LOOK 案A**: Canon EOS R5 + 50mm f/1.8／背景・小物は creamy bokeh／切り抜き輪郭禁止／粒感
- DOF・距離ボケ・反射・陰影。複数被写体はピンボケ必須

---

## 1. フロー

```text
【人】PACKAGE_TRUTH保存
 → B-③ → サブ採用CK
 → B-④ → compose（--auto-export 既定ON）
      → 品番キー×最大10枚を楽天目視フォルダへ（全子FO禁止）
 →【人】上記フォルダのみ目視
 →（NG）review_loop → 再export → 再目視
 →（OK）必要時 --to-checked-children → C「楽天のみ」→アップ → Amazonは U4
```

### 人間が見る場所

| 見る | パス |
|------|------|
| **唯一の目視先** | `…\03.楽天・Yahoo!商品登録（CSV一括UL）\02.楽天アップロード画像保存場所` |
| 見ない（作業用） | `…\05.画像生成（セットMAIN）\00.テスト出力\sub_image_ai_compose\…` |
| 正本素材のみ | `…\05.画像生成（セットMAIN）\01.amazon白抜きベース` |

### B-④（メニュー）

- **JAN↔正本画像の紐付けダイアログ**（＋ベース色）→ 確定後、**同一ダイアログ内**に Cursor 指示を表示（`--package-truth`・`--base-color` 付き）
- 続けて **「compose後の目視チェック（必須）」** パネル: チェックリスト＋review_loop 用 Agent 指示＋再生成コメント・テンプレ
- ※ `google.script.run` 成功後に別 `showModalDialog` を開くと「生成中…」で止まるため、指示はパネル切替で返す（要 `clasp push`）
- ※ メニュー実行はダイアログ表示で **完了**（7〜9秒程度が正常）。「終わらない」＝画面上の紐付け待ち／確定後の `finishB4AfterTruthBind`
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

### auto-export

- 既定ON（`--no-auto-export` で停止可）
- **目視**: 品番キー1件あたり最大10枚（既定 `pick=ab` = A/B両出し）。キー優先: 明示product/child → **JAN** → 親SKU → 代表1子 → composeフォルダ名
- **全セット子への複製フォールバックは禁止**（出品CK0件でも全子コピーしない）
- 本番で出品CK子へ配るときだけ: `export_sub_images_for_rakuten_matrix.py --to-checked-children --pick a`
- 再生成後も `export_plan.json` から同フォルダへ再投入（plan のキーに従う）

---

## 2. コマンド例

```text
cd tools\set_main_image

python sub_image_b3_curate.py --jan 4538872281013 --classify

python sub_image_ai_compose_poc.py --providers openai --openai-quality medium --target-slots 5 --auto-export --jan 4538872281013

# ログ末尾の HUMAN_REVIEW_DIR を開いて目視
# NG時のみ:
python sub_image_review_loop.py --compose-dir "…\00.テスト出力\sub_image_ai_compose\4538872281013_YYYYMMDD_HHMMSS"
```
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
- [ ] レビューUIでチェック＋要望→再生成ができ、チェック0で完了できる
- [ ] 人手は採用CK＋レビュー／A/B目視のみで手順が閉じる
- [ ] `{子SKU}_subN.jpg` が正しい子行の楽天サブNへ入り、アップ後マスタURLが埋まる
- [ ] Property false でファイル名自動が止まる
- [ ] Amazon 新規サブ作業を要求していない（U4 REUSE）
- [ ] 楽天CSV／Amazon MAIN ε 非回帰

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-10 | **ベース色 PoC 3択**（B-④＋`--base-color`／`tonmana_palette.py`）。要 clasp push。 |
| 2026-08-10 | B-④紐付け後の指示を**同一ダイアログ内表示**（`showModalDialog` 二重呼び出しハング修正）。要 clasp push。 |
| 2026-08-10 | B-④に **目視チェック必須パネル**（review_loop指示＋コメントテンプレ）。`review_feedback_templates.py`。要 clasp push。 |
| 2026-08-10 | 目視 auto-export は品番キー×最大10枚（pick=ab）。全セット子FO禁止。`--to-checked-children` は出品CK子のみ。CAMERA_LOOK 案A。 |
| 2026-08-09 | **1話完結**: compose後 auto-export で楽天フォルダへ直出し。人間目視は `02.楽天アップロード画像保存場所` のみ。 |
| 2026-08-09 | **PACKAGE_TRUTH／LOCK強化／文字量上限／SEO／人間レビュー再生成ループ**。B-④正本確認。Vision QA・ベース色メニューは対象外。 |
| 2026-08-09 | **B-④**: サブ採用CKの JAN から Python／Cursor 指示ダイアログ（GAS非合成）。 |
| 2026-08-09 | 初版。楽天出口本線。 |
