# サブ画像コース（楽天出口）— 要件

**文書種別**: 要件定義  
**最終更新**: 2026-08-09  
**状態**: **実装済（コード）**／実機検収は HUMAN_RUN  
**承認**: [LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md](LV4_SUB_IMAGE_RAKUTEN_COURSE_APPROVAL.md)  
**手順**: [D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)  
**投入レイヤ**: [D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)  
**PoC参照**: [D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md](D_MENU_SUB_IMAGE_POC_HUMAN_RUN.md)

---

## 1. ゴール一文

B-③で採用した競合サブを心理プロセス順に AI 合成し、**子SKU付きファイル名**で楽天マトリクスへ自動投入→R-Cabinet まで載せる。人手は採用レ点と短い目視に限定する。

---

## 2. 確定スコープ

| 項目 | 決定 |
|------|------|
| 出口 | 楽天まで（生成→`{子SKU}_subN`→マトリクス自動→R-Cabinet） |
| Amazon | 既存 **REUSE_RAKUTEN（U4）**。本コースで PT 新規生成しない |
| Yahoo | 別チケット |
| 人手 | **サブ採用CK＋PACKAGE_TRUTH準備＋レビューUI（チェック／要望）＋短い目視**。テーマ／phaseOrder は自動参照列（編集しない） |
| Vision セット数→SKU | 対象外 |
| Vision 自動QA | 対象外（人間レビューループで代替） |
| ベース色メニュー | **PoC 3択**（beige / warm_white / soft_gray）。B-④で選択 → `--base-color`。背景・カードのみ（PACKAGE_LOCK） |
| パーツ単位テーマ分類 | 対象外（将来メモ） |

---

## 3. 作るもの

| ID | 要件 |
|----|------|
| S1 | B-③自動参照列: `参照themeId`／`参照phaseOrder`／`参照cluster`。LPページ分類結果。人手編集対象外 |
| S2 | 人手はサブ採用CKと合成 A/B の短い目視のみ |
| S3 | AI compose v4: OpenAI medium・PACKAGE_TRUTH必須・PACKAGE_LOCK・正本OCR禁止・文字量上限・SEO |
| S3b | **1話完結 auto-export**: 目視は品番キー×最大10枚（`{キー}_{themeSlug}_subN`）。全セット子FO禁止。出品CK子への複製は `--to-checked-children` |
| S3c | B-④でJAN↔正本画像の人間紐付け（ファイル名にJAN不要）。写真実写は `photo_realism_rules.py` |
| S3d | B-④ベース色3択（`tonmana_palette.py`）→ compose `--base-color`／`--tonmana`。再生成は `run_meta.baseColor` を継承 |
| S4 | export: phaseOrder 順に `{子SKU}_{themeSlug}_subN.jpg`（N≤10） |
| S5 | 配置先=楽天 Drive ソース。未使用退避は人手（破壊的整理しない） |
| S6 | マトリクスファイル名自動（SKU紐付け要件 R1–R4） |
| S7 | 永続の正=マスタ楽天サブURL。Amazonは U4 REUSE のみ |
| S8 | ログ: runId／参照列／export／autobind 件数 |
| S9 | E2E HUMAN_RUN＋検収 |

---

## 4. 作らないもの

- Yahoo サブ／Amazon PT 新規・ONLY 本線化
- Vision セット数マッチ／パーツ単位分類／Vision自動QA
- ベース色の自由パレット（3択以外）／商品パッケージの色替え
- `generateRakutenCSV`・アップ時リネーム再設計
- Drive `03\02` 破壊的整理の自動化

---

## 5. 受け入れ条件

- [ ] B-③に参照列が自動で入り、人が phaseOrder を編集しなくても compose→export が動く
- [ ] 人手は採用CK＋短い目視のみで缶飯系 JAN が通る手順になっている
- [ ] `{子SKU}_subN.jpg` が正しい子行サブN→アップ後マスタURL
- [ ] Property off でファイル名自動停止
- [ ] HUMAN_RUN に U4 REUSE のみと明記（Amazon新規画像作業を要求しない）
- [ ] 楽天CSV／Amazon MAIN ε 非回帰

---

## 6. バックログ

- Yahoo サブ命名・配置
- Amazon ε セット個数マッチ共通化
- パーツ単位テーマ分類

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-10 | **S3d ベース色 PoC 3択**（beige/warm_white/soft_gray）。B-④＋`--base-color`。商品色は PACKAGE_LOCK。 |
| 2026-08-09 | 1話完結 auto-export（楽天フォルダ直出し・目視1フォルダ）。 |
| 2026-08-09 | PACKAGE_TRUTH必須・LOCK強化・文字量上限・SEO・人間レビュー再生成。Vision QA／ベース色は当時対象外→S3dでPoC化。 |
| 2026-08-09 | 初版。楽天出口・人手最小化・U4 REUSE。 |
