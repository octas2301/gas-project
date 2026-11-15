# 楽天画像 SKU 自動紐付け — 人間手順

**状態**: **実装済（要 clasp push → 実機検収）**  
**承認**: [LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md](LV4_RAKUTEN_IMAGE_SKU_AUTOBIND_APPROVAL.md)  
**要件**: [D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)／[D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md](D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md)  
**Cコース**: [D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)

---

## 0. 前提

- マスタにセット数違いの**子SKU行**があり、`A.セット商品数`（またはバリエーション名の数字）が埋まっている
- 生成は **`{子SKU}_rakuten.jpg`**（MAIN）／**`{子SKU}_{theme}_subN.jpg`**（サブ・**親あたり1セット**）
- 置き場: 楽天 Drive ソース（`DRIVE_IMAGE_SOURCE_FOLDER_ID`＝未整理画像フォルダ）直下、または作業後にマトリクスが読む位置
- **レイアウト拡張（既定ON）**: メイン2〜10＝他セット昇順、サブ＝親のみ投入→全子へ**上書き**共有、親行＝最少セット子。戻し `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED=false`
- **混在／テーマ別の「採用分を出口へ」**: JAN名や全子複製はしない。親1セット（代表子SKU名）だけ出す。全子へ同じ絵を置くと未整理・候補へ溢れる

---

## 1. 生成（ローカル・Agent 可）

### MAIN（セット数→子SKUはマスタから）

```text
cd tools\set_main_image
python compose_set_main.py --checked-only --malls rakuten --rakuten-engine layered --text-color 1 --out-dir "（楽天アップロード or ソースへコピー）"
```

出力名は既定で `{子SKU}_rakuten.jpg`（セット数はマスタ行の総個数）。

セット数 N だけ指定して子SKUを解決する例:

```text
python -c "from pathlib import Path; from master_sets import load_set_children_for_parent, resolve_child_by_set_count; ch,_=load_set_children_for_parent(Path('master.csv'),'親SKU'); print(resolve_child_by_set_count(ch, 3).child_sku)"
```

### サブ（AI compose 等 → マトリクス用名）

```text
python export_sub_images_for_rakuten_matrix.py --child-sku "子SKU" --src-dir "…\openai" --out-dir "（楽天ソース）"
```

`S01_….jpg` 等を **phaseOrder／ファイル順**で `{子SKU}_sub1.jpg` … にコピーする。

---

## 2. GAS（人間）

1. **`clasp push`**（Agent からの force push は承認ゲートあり → **人間がローカルで実行**）
2. 画像を楽天ソース Drive へ配置（例: セット違い2子以上の `{子SKU}_rakuten.jpg`）
3. トップ **C → 楽天のみ**（または従来マトリクス生成）
4. 完了ダイアログ／ログで `rakutenFilenameAutobind` の MAIN/サブ件数を確認
5. sheet を目視（誤りだけ手修正）
6. **リネーム＆アップロード** → マスタ URL 確認
7. 非回帰: Amazon C① の MAIN 自動（既存ε）が従来どおり動くこと

戻し: Script Property `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false`

### 実機検収ステータス（2026-08-09）

| 項目 | 状態 |
|------|------|
| コード／docs | **済** |
| `clasp push` | **人間待ち** |
| セット数違い JAN での目視・アップ | **人間待ち**（§3 チェックリスト） |

---

## 3. 検収チェック

- [ ] セット数違いの子が2つ以上で、各 `{子SKU}_rakuten.jpg` が正しい行のメイン1
- [ ] 既存メインあり行は非上書き
- [ ] `{子SKU}_subN.jpg` → サブN
- [ ] メイン2〜10に他セットMAIN（昇順）が入る
- [ ] 全子＋親のサブが同一／親＝最少セット子と同じメイン・サブ
- [ ] Property false でファイル名自動／レイアウト拡張が止まる
- [ ] アップ後マスタ URL・Amazon ε 非回帰

---

## 4. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-10-26 | 出口は親1セット（子SKU名）。JAN／全子複製をやめる（枠外候補対策）。clasp不要 |
| 2026-09-05 | レイアウト拡張 L1〜L3（[LAYOUT](D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md)） |
| 2026-08-09 | 初版。 |
