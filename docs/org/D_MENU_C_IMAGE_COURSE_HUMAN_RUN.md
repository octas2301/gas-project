# C. 画像コース — 人間手順

**状態**: **実装済**（D型選択モーダル＋一時Property ON／復元。要シート再読込）  
**承認**: [LV4_C_COURSE_CONSOLIDATION_APPROVAL.md](LV4_C_COURSE_CONSOLIDATION_APPROVAL.md)  
**親**: [D_MENU_U2_HUMAN_RUN.md](D_MENU_U2_HUMAN_RUN.md)／[D_MENU_U4_HUMAN_RUN.md](D_MENU_U4_HUMAN_RUN.md)／[D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)

---

## 0. 日常フロー（Amazon新規・最小）

```text
（常設）AMAZON_IMAGE_CANDIDATE_FOLDER_ID = Drive07
 → トップ「C. 画像コース（選択可）」
 → 「Ama新カタログ①：メイン画像を07保存⇒MAIN自動（目視）」
 →【人間】07へ `{子SKU}_amazon.jpg` 配置（未配置なら）／sheetの Amazon MAIN を目視（子SKU一致は自動投入済み。誤りだけ修正）
 → もう一度Cを開き「Ama新カタログ②：楽天サブ→マスタ反映→URLをR2へUP」
   （U4は URL充足スキップ・失敗SKU継続・長時間は自動再開。途中失敗後の再実行も可）
 → 以降は D（**Drive 02 は触らない**。手ZIPはSC用のみ）
```

楽天のみ: モーダルで **「楽天のみ」**（U4なし・一時ONなし）。

**候補の並び（2026-08-01 変更）**: Amazon 白抜きは **1子SKU＝1枚**。候補は全行タイルではなく、**子SKU行に1枚ずつ**置く。ファイル名に子SKU（無ければ親SKU）を含むものはその行へ、残りは候補が無い行へ上から順に入る。他SKUの画像を手で消す必要はない。旧・全行タイルに戻すときだけ `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED=true`。

**MAIN自動（2026-08-08・U2-ε）**: ファイル名に**子SKU**を含む候補は、当該行の **Amazon MAIN（BX）が空なら自動投入**。親SKUのみ一致／順置き余りは MAIN 自動しない。既存 MAIN は上書きしない。戻しは `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED=false`。

**楽天 MAIN/サブ自動（2026-08-09・SKU紐付け転用）**: 楽天ソースのファイル名に**子SKU**を含む画像は、マトリクス生成時に当該子行の **楽天メイン画像1**（空なら）へ自動投入。`{子SKU}_subN.jpg`（または `_pN`）は空の **楽天サブ画像N** へ。親SKUのみ一致は候補のみ。セット数はマスタで子SKUを決めてからファイル名に載せる（VisionでN推定しない）。戻しは `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false`。詳細 [要件](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)／[HUMAN_RUN](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_HUMAN_RUN.md)。

**サブ画像本線（楽天出口）**: B-③→compose v3→export→上記自動投入→R-Cabinet。Amazonは U4 REUSE。[E2E HUMAN_RUN](D_MENU_SUB_IMAGE_RAKUTEN_COURSE_HUMAN_RUN.md)。

手動で U2/U4 を true にする必要はない（C本線）。終わったら false にする作業も不要（自動復元）。
P1: 07配置後の Drive File は GAS のみ（[P1 HUMAN_RUN](D_MENU_P1_HUMAN_RUN.md)）。

---

## 1. メニュー配置

| 場所 | 内容 |
|------|------|
| トップ **C. 画像コース（選択可）** | モーダルで「Ama新カタログ①／②／楽天のみ／前提チェック」を選択 |
| **Z → Cコース（互換・分割）** | C-0／C-1楽天／C-1 Amazon／C-2 |
| **Z → C-Amazon（互換・分割）** | ①〜④＋旧マトリクス生成（**単独は従来どおり Property 要**） |
| D補助／Z → E | E-1/E-2 は C互換（一時ON込み）。E-3〜E-5 は承認〜GENERATED用 |

「Ama新カタログ②」のモーダル表示には **書込あり**（MAINのマスタ保存・Drive 02コピー・楽天サブのR2アップロード・Amazon PT URL反映・兄弟PT補完）を明記している。「実行」ボタンを開始前確認とするため、追加のOKダイアログは出さない。Z／E互換からC-2を単独実行した場合は従来の開始前確認を出す。

同一親内で今回U4対象の子SKUに `Amazon PT URL` の空欄がある場合、親の下から見て最初の `Amazon PT URL` 非空子SKUをコピー元として補完する。既存URLおよび今回の対象外SKUは変更しない。

---

## 2. Script Properties（必要最小限）

| Key | 扱い |
|-----|------|
| `AMAZON_IMAGE_CANDIDATE_FOLDER_ID` | **常設必須**（Drive 07） |
| `AMAZON_IMAGE_U2_ENABLED` | C本線では**常設不要**。C-1/C-2（E互換含む）実行中だけ一時 `true` → finally で元へ（未設定なら削除） |
| `AMAZON_U4_URL_EMBED_ENABLED` | C本線では**常設不要**。C-2実行中だけ一時 `true` → 復元 |
| `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED` | 未設定＝**1行1枚**（既定）。`true` で旧・全行タイル表示（戻し用） |
| `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED` | 未設定＝**true**（子SKU名一致→MAIN自動）。`false` で旧・手ドラッグ運用 |
| `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED` | 未設定＝**true**（子SKU名一致→楽天メイン1／`_subN`→サブN）。`false` でファイル名自動オフ（旧 Vision／手ドラッグ中心） |

Z単独の C-Amazon①〜④／21-⑦ は従来どおり各トグルが必要。

---

## 3. レ点規則（変更なし）

| マスタのレ点 | C に出る行 |
|--------------|------------|
| 子に1つ以上レ点 | その親＋レ点付き子だけ |
| 親のみレ点 | 親＋全子（楽天互換） |

Amazon は **対象セットの子だけにレ点**推奨。

---

## 4. 検収チェック

- [ ] トップCが1項目で、選択モーダルが開く
- [ ] 「Ama新カタログ①」が U2 未設定でも動く（実行後 U2 復元）
- [ ] 「Ama新カタログ②」が書込警告付きで、U4未設定でも動く。実行後 U2/U4 復元
- [ ] 「楽天のみ」が従来どおり
- [ ] ZにC-0〜C-2とC-Amazon①〜④が残る
- [ ] Z単独①〜④は Property 無しだと従来どおり止まる
- [ ] 同時実行時は「別のCコース実行中」で停止
- [ ] 候補が子SKU行に1枚ずつ並ぶ（他SKUの画像が同じ行に出ない）
- [ ] ファイル名に子SKUを含む画像はその行に入り、**Amazon MAIN へ自動投入**される（既存MAINは非上書き）
- [ ] `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED=false` で自動投入が止まる
- [ ] `{子SKU}_rakuten.jpg` が正しい子行の**楽天メイン1**へ自動投入される（既存非上書き）
- [ ] `{子SKU}_subN.jpg` が空の**楽天サブN**へ入る
- [ ] `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED=false` で楽天ファイル名自動が止まる
- [ ] `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED=true` で旧表示に戻る
- [ ] 今回U4対象でPT空の子SKUだけ、同親の最上位非空子SKUから `Amazon PT URL` が補完される
- [ ] 既存 `Amazon PT URL` と今回U4対象外のSKUは上書きされない

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-09 | **楽天 SKU 自動紐付け**: ファイル名子SKU→楽天メイン1／`_subN`→サブ。Property `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED`。[要件](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)。 |
| 2026-08-08 | **U2-ε**: 子SKU名一致で Amazon MAIN（BX）自動投入。人間は目視のみ。Property `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED`。 |
| 2026-08-02 | Cの表示名を「Ama新カタログ①／②」へ変更。U4末尾に対象SKU限定の兄弟PT URL補完を追加。 |
| 2026-08-01 | **候補を1子SKU＝1枚に変更**（名前一致→残りは順に配置）。戻しは `AMAZON_IMAGE_CANDIDATE_TILE_ALL_ENABLED=true`。 |
| 2026-08-01 | 初版。Cコース実装に合わせた手順。 |
| 2026-08-01 | C本線の U2/U4 一時ON→復元。常設は07フォルダIDのみ。 |
| 2026-08-01 | トップCをD型選択モーダル1本へ。C-0〜C-2はZ互換に残置。 |
| 2026-08-01 | P1: 日常で02を触らない旨を追記（[P1](D_MENU_P1_HUMAN_RUN.md)）。 |
