# セットMAIN画像 — AI PoC 手順（Nano Banana）

**承認**: [LV4_SET_MAIN_IMAGE_PHASE_AI_APPROVAL.md](LV4_SET_MAIN_IMAGE_PHASE_AI_APPROVAL.md)  
**レイアウトルール**: [SET_MAIN_LAYOUT_RULES.md](SET_MAIN_LAYOUT_RULES.md)／`tools/set_main_image/layout_rules.json`  
**ツール**: `tools/set_main_image/ai_compose_poc.py`  
**作業ルート**: `G:\マイドライブ\05.画像生成（セットMAIN）\`  
**出力（PoC）**: `00.テスト出力\`

Pillow Phase A（`compose_set_main.py`）は本線ではない。比較用に残置。

---

## 0. フォルダ

```text
05.画像生成（セットMAIN）/
  01.amazon白抜きベース/   … 商品ベース
      （任意）○○ヒーロー.jpg / ○○単体.jpg で使い分け
  02.楽天ベース/           … デザイン固定ベース（数字なし推奨）
  03.amazon見本/           … 配置・大きさ・重なりの見本
  04.楽天見本/             … 書体・金丸内バランスの見本
  06.競合事実参照/         … 任意手置き／_cache（自動）
  99.octas…/               … 食品時（浮遊シール・軽い回転）
  00.テスト出力/           … AI PoC の出力先
```

**01のファイル名タグ（任意）**

| タグ | 用途 |
|------|------|
| `ヒーロー` | メイン（開封・スプーン等。**Octas焼き込みなし**推奨） |
| `単体` | 在庫個（セットの他の個体） |
| （なし）＋`単体`あり | タグ無し最新をヒーロー、`単体`を在庫に使用 |
| どちらもなし | 従来どおり1枚を全個体に使用 |

---

## 1. 準備（一度だけ）

```text
cd tools\set_main_image
python -m pip install -r requirements-ai.txt
```

### APIキー（重要）

GAS の **Script Properties**（`GEMINI_API_KEY` / `OPENAI_API_KEY`）はローカル Python から読めません。  
PoC 実行前に、同じ値を次へコピーしてください（フォルダは gitignore）:

```text
tools\set_main_image\secrets\gemini_api_key.txt   … 1行
tools\set_main_image\secrets\openai_api_key.txt   … 1行
```

または OS のユーザー環境変数 `GEMINI_API_KEY` / `OPENAI_API_KEY`。

**モデル**

| エンジン | 既定 | 禁止・注意 |
|----------|------|------------|
| Gemini（Nano Banana） | 通常 Flash Image の最新（list／fallback `gemini-3.1-flash-image`） | **PRO禁止** |
| OpenAI（ChatGPT画像） | `gpt-image-1`（通常帯・精度比較用） | `SET_MAIN_OPENAI_IMAGE_MODEL` で上書き可 |

---

## 2. PoC実行（Agent）

### Amazon（見本レイアウト拘束＋トレース）— 量産バッチ

```text
cd tools\set_main_image
python amazon_ai_batch.py ^
  --master-csv "…\master_export.csv" ^
  --parent-sku sanky-…-oya ^
  --engine gemini
```

- 出力先既定: Drive **07**（`04.amazonカタログ作成…/07.白抜きの置き場…`）`{子SKU}_amazon.jpg`
- 見本選定: `amazon_blueprint.py`（N帯×patternHint×ink密度）。理由は各 `*_amazon_trace.json` と `AMAZON_AI_LOGIC_*.json`
- **競合事実参照**（プルタブ等の幻覚抑制）: バッチ先頭で1回解決。トレースに `IMAGE_FACT_COMPETITOR_REALITY`
- 成功後: `01.amazon白抜きベース` 直下のベース → `処理済み/`（既に `処理済み` なら `--no-move-base`）

#### 競合事実参照のソース（優先順）

| 順 | ソース | 備考 |
|----|--------|------|
| 0 | `06.競合事実参照/{親SKU}/` 手置き画像 | 任意 |
| 1 | マスタ `▼マスタ(参考情報(画像URL))` | 卸サイト等。物理事実向き |
| 2 | `競合店ASINコード` → Keepa | `secrets/keepa_api_key.txt` または `KEEPA_API_KEY` |
| 3 | `競合AmazonページURL` → og:image / Keepa | スクレイプ失敗時あり |
| — | **使わない** | `ASIN貼り付け（Keepa用）`（別商品混入リスク） |

01ベースとの Vision 一致（overall≥70 等）を満たさない場合は **参照なし**で続行。

単発 PoC:

```text
python ai_compose_poc.py --engine gemini --mall amazon --set-count 4 --food --stem POC2
```

出力: `POC2_gemini_amazon_set4_ai.jpg` ＋ `.json` ＋ `*_trace.json`  
トレースに `IMAGE_2_LAYOUT_BLUEPRINT` のファイル名・sha・選定理由・プロンプト全文・`apiPath` が入ります。

### 楽天（Python 3段組版・デザイン変更なし）

```text
python ai_compose_poc.py --mall rakuten --set-count 4 --unit 缶 --font-id tsukushi_mincho_like --stem POC4
python ai_compose_poc.py --mall rakuten --set-count 12 --unit 缶 --stem POC4
```

マスタCSVから単位・セット数（ヘッダ名解決）:

```text
python ai_compose_poc.py --mall rakuten --master-csv "…\master.csv" --parent-sku sanky-…-oya --child-sku … --stem POC4
```

- 単位ヘッダ正本: **`バリエーション単位`**
- 金丸内: **透過PNGグリフ**を重ねるのみ（黄土色クリア塗りはOFF）
- Canva書き出しは `05…/96.楽天透過文字/`（`digit_*.png` / `unit_*` / `text_set.png`）
- 無い場合は太字ゴシックで透過グリフを自動生成してテスト
- AIフル生成は呼ばない

---

## 2.1 検証バックログ（Pillow貼付チューニング）

| # | 項目 | 状態 |
|---|------|------|
| V2 | 大きさ＝本体縦×横 | **OK** |
| **V0** | **余白・縁際配置（左右上下を消す）** | **いま検証**（`--unit-only --layout edge_fill`） |
| V3 | 重なり上限 | V0の後 |
| V4 | N=4〜 占有 | V0連動 |
| V5 | motion / static | 未 |
| V6 | 量産→07 | 未 |

### 縁際配置の提案（N≤4）／N≥5

| N | パターン | 寄せる辺 |
|---|----------|----------|
| 2 | 左=unit / 右=hero（手前） | **底＋左＋右** |
| 3 | 左〜中央 unit×2 / 右 hero | **底＋左＋右** |
| 4 | 四隅寄り 2×2（hero=右下） | **四辺** |
| ≥5 | **hero=左下固定**、他は右・上へ | 底＋左を優先、残り埋め |

余白目標: 各辺 **≤2%**（`EDGE_MARGIN_RATIO`）。テスト中はヒーロー原画を外し **単体のみ**で全個体を作成。

## 3. 検収チェック

- [ ] Amazon: 見本の重なり・大小に近い。コピペ格子でない  
- [ ] Amazon: **縦横比が01と同一**（缶が縦伸び／横伸びしていない）  
- [ ] Amazon: **蓋・ラベル文字が捏造されていない**（単体／事実画像どおり。読めない偽字なし）  
- [ ] Amazon: N≤4で **単体がヒーローより小さく見えない**（同一スケール）  
- [ ] Amazon: **個数がセット数と完全一致**（裏積みに見えない・数えられる）  
- [ ] Amazon: 列が規則的。縦長なら斜め等の動きあり、画面が埋まっている  
- [ ] Amazon: 蓋・プルタブ等が **事実参照と矛盾しない**（二重タブ等なし）。`competitorFact.used=true` を確認  
- [ ] 楽天: ベースが同じに見える。金丸の数字だけが N  
- [ ] json の `modelId` が `*-pro-image` でない  
- [ ] 不合格なら見本の差し替え or プロンプト調整（別承認で本線化判断）

---

## 4. 人間／Agent 分担

| 担当 | 内容 |
|------|------|
| 人間 | 見本配置、AI Studioキー発行、目視検収、GAS/SC |
| Agent | `pip`・PoC実行・ログ確認・docs更新 |

---

## 5. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | AI PoC 初版。通常Flash最新・PRO禁止。 |
| 2026-08-02 | 見本ルール文書化・Amazon透過前処理・楽天1/2桁・筑紫/源柔近似。 |
| 2026-08-02 | Amazon見本＋AI量産バッチ（`amazon_ai_batch.py`）・見本選定ロジック明示・07出力。 |
| 2026-08-02 | 硬制約追加: 個数完全一致／裏積み禁止／列の規則性／動き＋画面埋め。焼鳥丼再生成。 |
| 2026-08-02 | ヒーロー＝01ベース保持（縦横比・デザイン変更禁止）。焼鳥丼再生成。 |
| 2026-08-03 | 競合事実参照（マスタ参考画像URL／競合ASIN・Keepa／06手置き）＋一致ゲート。焼鳥丼再生成。 |
| 2026-08-03 | チューニング開始: 第1閾値 **inkFillMin**（N=2–3: 0.48）。`00.テスト出力/TUNING_fill048_*` で検証。 |
| 2026-08-03 | N≤4同サイズ＋強め重なり／Octas浮遊+|θ|=6–10°／01の`ヒーロー``単体`タグ使い分け。`TUNING_n2_sameSize_heroUnit_*`。 |
| 2026-08-03 | 強化: 縦横比厳禁／蓋・ラベル捏造禁止（提供＋事実のみ）／N≤4は奥行き縮小禁止。`TUNING_n2_aspectLock_*`。 |
| 2026-08-03 | Pillow等倍貼付PoC（`amazon_paste_poc.py`）。AI描き直しなし。`TUNING_paste_n2_*`。 |
| 2026-08-03 | 焼き込み無しヒーロー＋浮遊Octas。単体タグ＋タグ無しヒーロー解決。`TUNING_paste_n*_cleanHero_*`。 |
| 2026-08-03 | 大きさ=本体縦横contain。処理済みは衝突時リネーム。`TUNING_paste_n*_bodyWH_*`。 |
| 2026-08-03 | 検証V3: pairOverlapMax=0.35 / heroVisibleMin=0.70。`TUNING_paste_n*_ov35_*`。 |
| 2026-08-03 | V0縁際配置＋unit-only。N≤4提案／N≥5左下。`TUNING_edge_unitOnly_n*`。 |
| 2026-08-03 | edge_fill: 重なり硬制約（≤0.35）下でスケール最大化。`TUNING_edge_ov35_n*`。 |
