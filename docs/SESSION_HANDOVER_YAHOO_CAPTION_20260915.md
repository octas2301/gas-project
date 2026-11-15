# セッション引継ぎ — Yahoo商品説明テンプレ（2026-09-15）

**用途**: このチャットをエージェント一覧から開けない／再開できないときの入口。新 Agent は本ファイルを最初に読む。  
**旧チャット**: Cursor agent-transcripts  
`35d4a68e-0550-4e01-9d5c-4166753a7d61`（タイトル目安: Yahoo caption P5/P6）

---

## 0. 新 Agent への起動文（コピペ）

```
gas-project の docs/AGENT_HANDOVER.md に従い、まず docs/CURRENT_PHASE.md と
docs/SESSION_HANDOVER_YAHOO_CAPTION_20260915.md を読め。
担当は Yahoo商品説明テンプレ（案A'・P5/P6）。実装は承認後のみ。clasp push は人間。
```

---

## 1. 目的・方針（崩さない）

- **案A'**: 生成HTMLを editItem に**毎回同送**（省略＝Yahooデフォルトで装飾消える）
- **手貼り・SKUごとテンプレID増殖は本線外**（見本は `katsuo_odashi`＝おだし共通骨格）
- **P5固定文は仮本線。** 完成形はマスタ `楽天サブ画像1〜10` の当選サブの catch／overview。画と文は同一 themeId。載せる枚数＝当選の中をAI売れそう順の**上位3**（空欄・落選は入れない。人手で型を選ばせない）
- 出品 clasp はリポ直下 `C:\Users\takuy\Desktop\gas-project`（① `purchase_research` と混ぜない）
- Agent は `clasp push` しない。実装は承認後のみ

正本: [org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)

---

## 2. 実装済み（コード／運用）

| 段階 | 内容 | 状態 |
|------|------|------|
| P3/P4c | かつおシリーズ API反映・inject | 済 |
| P5 | 共通骨格＋魚種・内容量・セットリンク変数 | 済（API例 `ycr_20260915_021123` 等） |
| P6a | 本文画像自動差し込み。**楽天CDN直載せ禁止**（欠画） | 済 |
| P6b | 魚種lib命名＋GAS `fish_lib_{slug}`＋**公式 `uploadLibImage` API UP** | UP 6/0 済 |
| Yahoo.js | `uploadLib` / siblings（未チェック子もセット表） / A' inject | 済 |
| マスタ列 | `▼マスタ(Yahoo説明テンプレID)`（親に `katsuo_odashi` 可）／説明画像1..3 | 人間作成済 |

主要コード: `YahooCaptionRich.js` / `Yahoo.js`  
手順: `tools/yahoo_page_mock/paste/HUMAN_RUN_YAHOO_CAPTION_*.md`

---

## 3. 画像まわりの事実（重要）

### 3.1 選定ロジック（現行＝第1スライス。完成形は §4）
- **紹介文を読んで合う画を選ぶ処理はない**
- 骨格固定3枠: `t03_sizzle` → `t04_aji` → `t11_arrange`
- テキストは `かつお`→魚種置換、画像は同じ `fishJa` から slug で **同枠ファイル名差し替え**のみ
- 優先: 説明画像列 → 魚種lib(UP済) → かつおlib。楽天サブ直載せなし

### 3.2 素材ミスと修正
- 初回 P6b prep が `theme_new_full_*`（**古い／非本採用**）→ 人間が「採用されていない古いサブ」と指摘
- **修正済**: Drive `sub_image_ai_compose` の OpenAI 本採用候補へ差し替え再UP
  - まぐろ: `…/4906283045317_20260830_160211/03_openai/` の T03/T04/T11
  - ぶり: `…/4906283047410_20260831_123235/03_openai/` の T03/T04/T11
- `prep_odashi_lib_bodies.py` は上記パスが正（`theme_new_full` 禁止）
- UP: `python tools/yahoo_page_mock/upload_lib_images_api.py`（認証は出品SS Yahooマッピング B12–B15）
- **同一ファイル名上書き**のため CDN キャッシュで古い画が残る可能性あり → だめならファイル名リネーム＋GAS＋再inject

### 3.3 人間確認待ち
- Pro でまぐろ／ぶり本文3枚が**本採用画**になっているか（Ctrl+F5）
- 説明HTML再送はURL同一なら原則不要（キャッシュ問題時のみ再送／改名）

---

## 4. 完成形（要件ロック・実装未）

正本 [YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md](org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md) **§0.1**。

- 本文の種＝マスタ `楽天サブ画像1〜10` の入っている枚（catch＝見出し、salesTalk／subs＝補足、紹介はそれら＋魚種／主原料。内容量の数字は紹介に使わない。写真用overviewと商品説明文AIは種にしない）
- 画と文は同一 themeId
- 載せる枚数＝当選の中をAIが「売れそう」順に並べた **上位3**（当選が3未満ならある分だけ。落選は入れない。人が埋めない）
- 人手で型名・文案を増やさない。社長に型を選ばせない
- 固定 T03/T04/T11 は第1スライスの仮。完成形の枠契約ではない

旧「対策A–Dを社長が選べ」は **取り下げ**（正本の禁止に抵触）。

---

## 5. 別課題（このレーンで直さない）

| ID | 内容 |
|----|------|
| AI-DESC-FISH-MIX | Yahoo **商品情報**（explanation）に複合魚種が載る。captionテンプレ不合格ではない。**2026-09-17 コード済**（一意キーJOIN＋生成1SKU固定。`AI_DESC_FISH_MIX_JOIN_ENABLED` 未設定=ON）。要出品 clasp→おだしAI再生成→同期→説明API |
| YAHOO-PARTIAL-UPDATE | **出品済みの任意項目だけ直す器が無い**（2026-09-16 メモ。実装しない）。公式は `getItem`（XML）→変えたい欄だけ差し替え→`editItem` に現行値も同送。省略＝デフォルト上書き。マッピングシート形式では返らない。現行の説明API反映は画像再UPしないが、マッピング TRUE の data をマスタから再送する。おだし以外はテンプレスキップ。要件化・実装は別承認 |

---

## 6. 次アクション候補（承認後）

1. **今すぐ（人間）**: マッピング `point-immediate`＝`[[固定値:1]]` 済。出品CK付きかつお・まぐろ・ぶりに **Yahoo説明をAPI反映を1回**（`point_immediate=1` が乗る。説明HTMLも再送。出品実行は画像まで再送するので使わない）
2. Proで「獲得ポイント今すぐ利用」が対象か確認。未反映なら反映画面を一度
3. Pro目視で本採用画OKか（だめなら lib リネーム＋再UP＋説明再反映）
4. **YAHOO-PARTIAL-UPDATE は作らない**（§5）。未コミットの Yahoo説明差分は人間指示があるまで commit しない

---

## 7. 参照パス早見

- 要件: `docs/org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md`（§0.1）
- Agent規則: `.cursor/rules/yahoo-caption-auto.mdc`
- GAS: `YahooCaptionRich.js` / `Yahoo.js`
- prep: `tools/yahoo_page_mock/prep_odashi_lib_bodies.py` / `prep_lib_images.py`
- UP: `tools/yahoo_page_mock/upload_lib_images_api.py`
- 素材: `assets/yahoo_lib_src/` → `tools/yahoo_page_mock/out/lib_ready/`
- Property: `YAHOO_CAPTION_RICH_LOCK_ENABLED` / `YAHOO_CAPTION_RICH_P6_IMG_ENABLED` / `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED`
- P6c: `tools/yahoo_page_mock/build_yahoo_caption_themes.py` / `paste/HUMAN_RUN_YAHOO_CAPTION_P6C_THEMES.md`

---

## 8. 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-09-15 | 初版。チャット復活不可時の新Agent入口。画像本採用差し替え・選定ロジック限界・対策A–Dを固定 |
| 2026-09-16 | **P6c-1 本番前**: `build --prep-lib`＋`upload --p6c` 9/0・公開200。次は clasp＋Property ON＋説明API反映 |
| 2026-09-16 | **P6c pool＝当選のみ**: `_export_meta` の exports。messages 全件は禁止 |
| 2026-09-16 | **P6c pool＝マスタ楽天サブ1〜10**: 入っている枠が正。`_export_meta` 単独は使わない |
| 2026-09-16 | **P6c 文字数合わせ**: PC・スマホ同じ仕様。`--fit-only` で既存上位3を枠内。要 clasp＋説明API再反映 |
| 2026-09-16 | **P6c 紹介組立**: catch＋salesTalk／subs＋安全事実。写真用overviewは種にしない |
| 2026-09-16 | **紹介から内容量数字を除外**: 単位不問。大容量は残す。要 clasp＋説明API再反映 |
| 2026-09-16 | **YAHOO-PARTIAL-UPDATE 課題メモのみ**: getItem＋現行値同送の部分更新は未着手。今は説明API反映1回で point_immediate=1 |
